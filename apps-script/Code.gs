/**
 * Aljamea Security Control: backend.
 * Lives in a Google Sheet you own. Run setup() once, then deploy as a Web app.
 * Sheets: Posts, Guards, Checks, Issues, Log, Settings, Users (one row per person: Name, PIN, Role, Active). Photos go to a Drive folder.
 */
var TZ = 'Asia/Kolkata';
var HEAD = {
  Posts: ['Post'],
  Guards: ['ID', 'Name', 'Phone', 'PhotoId', 'AddedBy', 'AddedAt', 'RegularPost', 'RegularShift', 'Status', 'Agency'],
  Checks: ['Time', 'Date', 'Shift', 'Post', 'GuardId', 'Guard', 'Presence', 'Type', 'Related', 'Remark', 'PhotoId', 'By', 'Flags', 'Role', 'Lat', 'Lng', 'Acc', 'Hash'],
  Issues: ['ID', 'Created', 'Post', 'Category', 'Priority', 'Description', 'PhotoId', 'RaisedBy', 'Status', 'LastRemark', 'Updated', 'SolvedAt', 'To'],
  Log: ['Time', 'IssueID', 'By', 'Role', 'Action', 'Remark', 'PhotoId'],
  Settings: ['Key', 'Value'],
  Users: ['Name', 'PIN', 'Role', 'Active'],
  Alerts: ['Time', 'Type', 'Key', 'Level', 'Message', 'Audience'],
  Checkpoints: ['Time', 'Date', 'Id', 'Post', 'Result', 'Remark', 'PhotoId', 'By', 'Role']
};
var POSTS = ['Gate 1', 'Gate 2', 'Gate 4', 'Gate 6 (Rabvat)', 'Gate 8 & 9', 'MVD Gate 10', 'Mahad Gate',
  'Masakin front gate', 'Masakin back gate', 'Raj Manzil', 'Yusufi Makan', 'Ezzy Makan', 'Funun-ul-Quran',
  'Sarela Makan', 'Nuzul', 'Utility', 'Control Room'];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function ts_(t) { return new Date(String(t).replace(' ', 'T').replace(/T(\d):/, 'T0$1:') + '+05:30').getTime(); }
function clockMin_(time) {
  var m = String(time || '').match(/(\d+):(\d+)/); return m ? Number(m[1]) * 60 + Number(m[2]) : -1;
}
/** Shifts: Day 8 AM to 8 PM, Night 8 PM to 8 AM. */
function shiftByClock_(time) { var x = clockMin_(time); return x >= 480 && x < 1200 ? 'Day' : 'Night'; }
/** A guard is "outside his shift" only when clearly inside the OTHER shift. One hour either side of each handover is ignored. */
function outsideShift_(regShift, time) {
  var x = clockMin_(time);
  if (x < 0) return false;
  if (regShift === 'Day') return x >= 1260 || x < 420;
  if (regShift === 'Night') return x >= 540 && x < 1140;
  return false;
}
function dayBefore_(d) { var x = new Date(d + 'T12:00:00+05:30'); x.setDate(x.getDate() - 1); return Utilities.formatDate(x, TZ, 'yyyy-MM-dd'); }
function distM_(a, b, c, d) {
  var R = 6371000, r = Math.PI / 180, dLa = (c - a) * r, dLo = (d - b) * r;
  var h = Math.sin(dLa / 2) * Math.sin(dLa / 2) + Math.cos(a * r) * Math.cos(c * r) * Math.sin(dLo / 2) * Math.sin(dLo / 2);
  return 2 * R * Math.asin(Math.sqrt(h));
}
function median_(a) { a = a.slice().sort(function (x, y) { return x - y; }); var n = a.length; return n % 2 ? a[(n - 1) / 2] : (a[n / 2 - 1] + a[n / 2]) / 2; }
function roleOfName_(name, map) {
  var n = String(name || '').trim().toLowerCase();
  if (map && map[n]) return map[n];
  if (n === 'control room') return 'control';
  if (n === 'supervisor') return 'supervisor';
  return '';
}
function roleMap_() {
  var o = {}; rows_('Users').forEach(function (u) { o[String(u.Name).trim().toLowerCase()] = String(u.Role).trim().toLowerCase(); }); return o;
}
function isAbsent_(p) { return /^(Absent|अनुपस्थित)$/.test(String(p || '')); }
/** Start time (ms) of the shift a check belongs to. Night runs 8 PM to 8 AM, so a night check before noon belongs to the previous date's night. */
function shiftStart_(c) {
  var d = c.Date, x = clockMin_(c.Time);
  if (c.Shift === 'Night') { if (x >= 0 && x < 720) d = dayBefore_(d); return ts_(d + ' 20:00:00'); }
  return ts_(d + ' 08:00:00');
}
/** Handover hour: day staff such as bouncers work 9 AM to 9 PM, and night staff often stay till 9 AM.
 *  A scan in the first hour of a shift (8-9 PM night, 8-9 AM day) belongs to the PREVIOUS duty when the same man
 *  was also seen at the same gate in that previous shift. Otherwise it is the start of the new shift. */
function effStart_(c, all) {
  var s = shiftStart_(c), x = clockMin_(c.Time);
  var hand = (c.Shift === 'Night' && x >= 1200 && x < 1260) || (c.Shift === 'Day' && x >= 480 && x < 540);
  if (!hand) return s;
  var before = s - 12 * 3600 * 1000;
  for (var i = 0; i < all.length; i++) {
    var q = all[i];
    if (q.GuardId === c.GuardId && q.Post === c.Post && q.Time !== c.Time && !isAbsent_(q.Presence) && shiftStart_(q) === before) return before;
  }
  return s;
}
function hm_(t) { return String(t).slice(11, 16).replace(/^(\d):/, '0$1:'); }
/** Double duty = the guard was ALSO on post in the shift just before this one (night then day, or day then night).
 *  Working a different shift from his registered one, with the previous shift off, is only a shift change. */
function dutyFlag_(g, c, all) {
  if (!g || isAbsent_(c.Presence)) return '';
  c.GuardId = c.GuardId || g.ID;
  var start = effStart_(c, all), prev = start - 12 * 3600 * 1000;
  for (var i = all.length - 1; i >= 0; i--) {
    var p = all[i];
    if (p.GuardId !== g.ID || isAbsent_(p.Presence) || p.Time === c.Time) continue;
    if (effStart_(p, all) === prev)
      return 'DOUBLE DUTY: ' + g.Name + ' was also on post in the previous ' + (new Date(prev + 5.5 * 3600 * 1000).getUTCHours() === 8 ? 'Day' : 'Night') + ' shift (' + p.Post + ', ' + String(p.Date).slice(5) + ' at ' + hm_(p.Time) + ') and is on post again in this ' + c.Shift + ' shift.';
  }
  if (g.Agency === 'Inhouse' && g.RegularShift && g.RegularShift !== c.Shift)
    return 'SHIFT CHANGE: ' + g.Name + ' is registered for the ' + g.RegularShift + ' shift but is on the ' + c.Shift + ' shift now. He was not on post in the shift before, so this is not double duty. If we moved him, the Control Room should update his shift.';
  return '';
}
/** Run once from the editor: adds the new columns, creates the Alerts sheet, seeds each guard\'s shift and agency, repairs old flags. */
function migrate() {
  var ss = ss_();
  ['Guards', 'Checks', 'Alerts'].forEach(function (name) {
    var fresh = !ss.getSheetByName(name), sh = ss.getSheetByName(name) || ss.insertSheet(name), cols = HEAD[name].length;
    if (fresh) sh.getRange(1, 1, sh.getMaxRows(), cols).setNumberFormat('@');
    sh.getRange(1, 1, 1, cols).setValues([HEAD[name]]).setFontWeight('bold').setBackground('#1F3A5F').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });
  var st = settings_();
  if (!st.min_cover) setSetting_('min_cover', '0.9');
  if (!st.alert_email) setSetting_('alert_email', '');
  var rm = roleMap_(), checks = rows_('Checks'), gsh = ss.getSheetByName('Guards'), gs = rows_('Guards'), seeded = 0;
  var ch = ss.getSheetByName('Checks'), iRole = HEAD.Checks.indexOf('Role') + 1;
  checks.forEach(function (c) { if (!c.Role) { var r = roleOfName_(c.By, rm); if (r) ch.getRange(c._row, iRole).setValue(r); } });
  var iSh = HEAD.Guards.indexOf('RegularShift') + 1, iPo = HEAD.Guards.indexOf('RegularPost') + 1, iAg = HEAD.Guards.indexOf('Agency') + 1;
  gs.forEach(function (g) {
    var mine = checks.filter(function (c) { return c.GuardId === g.ID && c.Presence !== 'Absent'; });
    if (!g.RegularShift && mine.length) {
      var first = mine.filter(function (c) { return roleOfName_(c.By, rm) === 'supervisor'; })[0] || mine[0];
      gsh.getRange(g._row, iSh).setValue(first.Shift); gsh.getRange(g._row, iPo).setValue(first.Post); seeded++;
    }
    if (!g.Agency) gsh.getRange(g._row, iAg).setValue(/bouncer/i.test(g.Name) ? 'Avenger' : 'Sainik');
  });
  return { seeded: seeded, repaired: repairFlags() };
}
/** Recompute the DOUBLE DUTY flag on every old check, using each guard\'s registered shift. */
function repairFlags() {
  var ss = ss_(), sh = ss.getSheetByName('Checks'), gmap = {};
  rows_('Guards').forEach(function (g) { gmap[g.ID] = g; });
  var iF = HEAD.Checks.indexOf('Flags') + 1, n = 0, all = rows_('Checks'), byG = {};
  all.forEach(function (c) { (byG[c.GuardId] = byG[c.GuardId] || []).push(c); });
  /* Agency guards: their usual post and shift = where they were last seen on duty. */
  var gsh = ss.getSheetByName('Guards'), iSh = HEAD.Guards.indexOf('RegularShift') + 1, iPo = HEAD.Guards.indexOf('RegularPost') + 1;
  Object.keys(gmap).forEach(function (id) {
    var g = gmap[id]; if (g.Agency === 'Inhouse') return;
    var last = (byG[id] || []).filter(function (c) { return !isAbsent_(c.Presence); }).sort(function (x, y) { return ts_(x.Time) - ts_(y.Time); }).pop();
    if (last && (g.RegularShift !== last.Shift || g.RegularPost !== last.Post)) {
      gsh.getRange(g._row, iSh).setValue(last.Shift); gsh.getRange(g._row, iPo).setValue(last.Post); g.RegularShift = last.Shift; g.RegularPost = last.Post;
    }
  });
  all.forEach(function (c) {
    var parts = String(c.Flags || '').split(' | ').filter(function (x) { return x && x.indexOf('DOUBLE DUTY') !== 0 && x.indexOf('SHIFT CHANGE') !== 0; });
    var g = gmap[c.GuardId], df = g ? dutyFlag_(g, c, byG[c.GuardId] || []) : '';
    if (df) parts.push(df);
    var nf = parts.join(' | ');
    if (nf !== String(c.Flags || '')) { sh.getRange(c._row, iF).setValue(nf); n++; }
  });
  return n;
}
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

/** Run ONCE from the editor. Creates sheets, default settings and the photo folder. */
function setup() {
  var ss = ss_();
  Object.keys(HEAD).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var cols = HEAD[name].length;
    sh.getRange(1, 1, sh.getMaxRows(), cols).setNumberFormat('@');
    sh.getRange(1, 1, 1, cols).setValues([HEAD[name]]).setFontWeight('bold').setBackground('#1F3A5F').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });
  var posts = ss.getSheetByName('Posts');
  if (posts.getLastRow() < 2) posts.getRange(2, 1, POSTS.length, 1).setValues(POSTS.map(function (p) { return [p]; }));
  var st = ss.getSheetByName('Settings');
  if (st.getLastRow() < 2) {
    st.getRange(2, 1, 5, 2).setValues([
      ['supervisor_pin', '4721'], ['control_pin', '5832'], ['head_pin', '6943'],
      ['overdue_hours', '24'], ['folder_id', '']
    ]);
  }
  folder_();
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  SpreadsheetApp.getUi().alert('Setup done. Open the Settings sheet and change the three PINs, then deploy as a Web app.');
}

function settings_() {
  var rows = ss_().getSheetByName('Settings').getDataRange().getValues(), o = {};
  for (var i = 1; i < rows.length; i++) o[String(rows[i][0])] = String(rows[i][1]);
  return o;
}
function setSetting_(key, val) {
  var sh = ss_().getSheetByName('Settings'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if (String(rows[i][0]) === key) { sh.getRange(i + 1, 2).setValue(val); return; }
  sh.appendRow([key, val]);
}
function folder_() {
  var id = settings_().folder_id;
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('Security Control Photos');
  setSetting_('folder_id', f.getId());
  return f;
}
function rows_(name) {
  var sh = ss_().getSheetByName(name), v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  var h = v[0], out = [];
  for (var i = 1; i < v.length; i++) {
    var o = { _row: i + 1 };
    for (var j = 0; j < h.length; j++) {
      var cell = v[i][j];
      if (cell instanceof Date) cell = Utilities.formatDate(cell, TZ, h[j] === 'Date' ? 'yyyy-MM-dd' : 'yyyy-MM-dd HH:mm:ss');
      o[h[j]] = cell === '' ? '' : String(cell);
    }
    out.push(o);
  }
  return out;
}
function add_(name, obj) {
  var sh = ss_().getSheetByName(name);
  sh.appendRow(HEAD[name].map(function (h) { return obj[h] === undefined ? '' : String(obj[h]); }));
}
function setCells_(name, row, obj) {
  var sh = ss_().getSheetByName(name), h = HEAD[name];
  Object.keys(obj).forEach(function (k) {
    var c = h.indexOf(k);
    if (c > -1) sh.getRange(row, c + 1).setValue(String(obj[k]));
  });
}
function photo_(dataUrl, prefix) {
  if (!dataUrl) return '';
  var m = /^data:image\/\w+;base64,(.+)$/.exec(dataUrl);
  if (!m) return '';
  var blob = Utilities.newBlob(Utilities.base64Decode(m[1]), 'image/jpeg', prefix + '_' + Date.now() + '.jpg');
  var f = folder_().createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f.getId();
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function userOf_(pin) {
  pin = String(pin || '').trim();
  if (!pin) return null;
  var sh = ss_().getSheetByName('Users');
  if (!sh) return null;
  var us = rows_('Users');
  for (var i = 0; i < us.length; i++) {
    var u = us[i], role = String(u.Role || '').trim().toLowerCase();
    if (String(u.PIN).trim() === pin && String(u.Active || 'Yes').trim().toLowerCase() !== 'no' && { supervisor: 1, control: 1, head: 1 }[role]) {
      return { name: String(u.Name).trim(), role: role };
    }
  }
  return null;
}

function doGet() { return json_({ ok: true, app: 'Aljamea Security Control' }); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var q = JSON.parse(e.postData.contents);
    if (q.action === 'card') return json_(card_(q.t));
    if (q.action === 'board') return json_(board_(q.t));
    if (q.action === 'boardReport') return json_(boardReport_(q));
    var user = userOf_(q.pin);
    if (!user) return json_({ ok: false, error: 'Wrong PIN' });
    var role = user.role, by = user.name;
    switch (q.action) {
      case 'bootstrap': return json_(bootstrap_(role, by));
      case 'resetBoard': return json_(resetBoard_(role, q));
      case 'addGuard': return json_(addGuard_(q, by, role));
      case 'updateGuard': return json_(updateGuard_(q, by, role));
      case 'rounds': return json_(roundsOut_(role, q));
      case 'alerts': return json_(alertsOut_(role));
      case 'checkpoints': return json_(cpOut_(q));
      case 'submitCheckpoint': return json_(cpSubmit_(q, by, role));
      case 'submitCheck': return json_(submitCheck_(q, by, role));
      case 'raiseIssue': return json_(raiseIssue_(q, by, role));
      case 'issues': return json_(listIssues_(role, q));
      case 'updateIssue': return json_(updateIssue_(q, by, role));
      case 'checks': return json_(listChecks_(q.from || q.date || today_(), q.to || q.from || q.date || today_()));
      default: return json_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function bootstrap_(role, name) {
  var s = settings_();
  try { alertTick_(); } catch (e) {}
  return {
    ok: true, role: role, name: name, nightOff: NIGHT_OFF, overdueHours: Number(s.overdue_hours) || 24, rounds: ROUNDS, minCover: Number(s.min_cover) || 0.9,
    posts: rows_('Posts').map(function (r) { return r.Post; }).filter(String),
    guards: rows_('Guards').filter(function (g) { return g.Status !== 'Left'; }).map(guardOut_),
    boardToken: role === 'supervisor' ? '' : boardToken_(''),
    boardLinks: role === 'supervisor' ? [] : boardLinks_()
  };
}
function guardOut_(g) {
  return { id: g.ID, name: g.Name, phone: g.Phone, photo: g.PhotoId, post: g.RegularPost, shift: g.RegularShift, agency: g.Agency, share: shareToken_(g.ID) };
}
/* ---- Share link for one guard: an unguessable token, so one link shows only that guard. ---- */
var SHARE_SECRET_ = null;
function shareSecret_() {
  if (SHARE_SECRET_) return SHARE_SECRET_;
  var s = settings_().share_secret;
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); setSetting_('share_secret', s); }
  return (SHARE_SECRET_ = s);
}
function shareToken_(id) {
  if (!id) return '';
  var b = Utilities.computeHmacSha256Signature(String(id), shareSecret_());
  return Utilities.base64EncodeWebSafe(b).replace(/=+$/, '').slice(0, 18);
}
/* ---- Live "on duty" board: one secret link shows every gate and who is on duty this shift. Head Office can reset the link. ---- */
/* Zones: a link that shows only some gates, for that area's in-charge. '' = all gates. */
var ZONES = {
  '': { name: 'All gates', posts: null },
  masakin: { name: 'Masakin Zone', posts: ['Masakin front gate', 'Masakin back gate', 'Raj Manzil', 'Yusufi Makan'] }
};
function boardToken_(zone) {
  var k = 'board_token' + (zone ? '_' + zone : ''), t = settings_()[k];
  if (!t) { t = Utilities.getUuid().replace(/-/g, '').slice(0, 20); setSetting_(k, t); }
  return t;
}
function boardLinks_() {
  return Object.keys(ZONES).map(function (z) { return { zone: z, name: ZONES[z].name, token: boardToken_(z) }; });
}
function zoneOfToken_(t) {
  t = String(t || ''); if (!t) return null;
  var zs = Object.keys(ZONES);
  for (var i = 0; i < zs.length; i++) if (boardToken_(zs[i]) === t) return zs[i];
  return null;
}
function resetBoard_(role, q) {
  if (role !== 'head') return { ok: false, error: 'Only Head Office can reset the link' };
  var z = String((q && q.zone) || ''); if (!ZONES[z]) return { ok: false, error: 'Unknown zone' };
  var t = Utilities.getUuid().replace(/-/g, '').slice(0, 20); setSetting_('board_token' + (z ? '_' + z : ''), t);
  return { ok: true, boardToken: t, boardLinks: boardLinks_() };
}
/** Public report from the shared board link: the person must give a name; a photo is required. */
var REPORT_CATS = ['Guard sleeping / away from post', 'Guard absent', 'Uniform / dress', 'Complaint about guard', 'Other'];
function boardReport_(q) {
  var z = zoneOfToken_(q.t);
  if (z === null) return { ok: false, error: 'This link is not valid or has been reset.' };
  var posts = ZONES[z].posts, post = String(q.post || '');
  if (posts && posts.indexOf(post) < 0) return { ok: false, error: 'This gate is not in your zone.' };
  var name = String(q.name || '').trim().slice(0, 60), ph = String(q.phone || '').replace(/\D/g, '').slice(0, 13);
  if (name.length < 3) return { ok: false, error: 'Please write your name.' };
  if (!q.photo) return { ok: false, error: 'Please take a photo.' };
  var cat = REPORT_CATS.indexOf(q.category) > -1 ? q.category : 'Other';
  var cache = CacheService.getScriptCache(), n = Number(cache.get('brep') || 0);
  if (n >= 30) return { ok: false, error: 'Too many reports in the last hour. Please call the Control Room.' };
  cache.put('brep', String(n + 1), 3600);
  var who = name + (ph ? ' (' + ph + ')' : '') + ' · via ' + ZONES[z].name + ' link';
  var desc = (q.guard ? 'Watchman: ' + String(q.guard).slice(0, 60) + '. ' : '') + String(q.description || '').trim();
  if (desc.length < 3) desc = cat;
  return raiseIssue_({ post: post, category: cat, priority: q.priority === 'Urgent' ? 'Urgent' : 'Normal', description: desc, photo: q.photo }, who, 'admin');
}
function board_(t) {
  var zone = zoneOfToken_(t);
  if (zone === null) return { ok: false, error: 'This link is not valid or has been reset.' };
  var now = new Date(), nowS = now_(), h = Number(nowS.slice(11, 13)), shift = (h >= 8 && h < 20) ? 'Day' : 'Night';
  var startD = shift === 'Day' ? today_() : (h < 8 ? dateAdd_(today_(), -1) : today_());
  var start = ts_(startD + ' ' + (shift === 'Day' ? '08:00:00' : '20:00:00'));
  var guards = rows_('Guards').filter(function (g) { return g.Status !== 'Left'; }), byId = {};
  guards.forEach(function (g) { byId[g.ID] = g; });
  var from = dateAdd_(today_(), -1);
  var checks = rows_('Checks').filter(function (c) { return c.Date >= from && c.Shift === shift && c.Presence !== 'Absent' && ts_(c.Time) >= start; });
  /* Handover hour: day staff such as bouncers work till 9 pm, night staff till 9 am. A check of them in that hour is the end of
     their own duty, not the next shift, so leave them out of the current shift's list. */
  checks = checks.filter(function (c) {
    var g = byId[c.GuardId], m = clockMin_(c.Time); if (!g || !g.RegularShift || g.RegularShift === shift) return true;
    return !((shift === 'Night' && m >= 1200 && m < 1260) || (shift === 'Day' && m >= 480 && m < 540));
  });
  checks.sort(function (a, b) { return a.Time < b.Time ? 1 : -1; });
  var posts = rows_('Posts').map(function (r) { return r.Post; }).filter(function (p) { return p && p !== 'Control Room'; });
  if (ZONES[zone].posts) posts = posts.filter(function (p) { return ZONES[zone].posts.indexOf(p) > -1; });
  if (shift === 'Night') posts = posts.filter(function (p) { return NIGHT_OFF.indexOf(p) < 0; });
  var card = function (g, seen) { return { name: g.Name, phone: g.Phone, photo: g.PhotoId, agency: g.Agency, seen: seen || '' }; };
  /* Who is at a gate now = every guard seen there this shift (a gate can have several: gate guard, SO, agency supervisor).
     A guard seen at two gates is shown only at the gate where he was seen last. */
  var lastAt = {};
  checks.forEach(function (c) { if (c.GuardId && !lastAt[c.GuardId]) lastAt[c.GuardId] = c.Post; });
  var shown = {};
  var out = posts.map(function (p) {
    var mine = checks.filter(function (c) { return c.Post === p; }), list = [], seen = {};
    if (mine.length) {
      mine.forEach(function (c) {
        if (!c.GuardId || seen[c.GuardId] || lastAt[c.GuardId] !== p) return; seen[c.GuardId] = 1;
        var g = byId[c.GuardId]; if (g) { list.push(card(g, String(c.Time).slice(11, 16))); shown[c.GuardId] = 1; }
      });
    }
    return { post: p, checked: list.length > 0, moved: mine.length > 0 && !list.length, guards: list };
  });
  out.forEach(function (o) {
    if (o.checked) return;
    guards.filter(function (g) { return g.RegularPost === o.post && g.RegularShift === shift && !shown[g.ID] && !lastAt[g.ID]; }).forEach(function (g) { o.guards.push(card(g, '')); });
  });
  return { ok: true, shift: shift, updated: nowS, zone: ZONES[zone].name, all: !ZONES[zone].posts, reportCats: REPORT_CATS, gates: out };
}
/** Public: no PIN. Returns only name, photo, phone, gate, shift and agency of the one guard the link is for. */
function card_(t) {
  t = String(t || '').trim();
  if (t.length < 12) return { ok: false, error: 'This link is not valid.' };
  var gs = rows_('Guards');
  for (var i = 0; i < gs.length; i++) {
    var g = gs[i];
    if (shareToken_(g.ID) === t) {
      if (g.Status === 'Left') return { ok: false, error: 'This watchman is no longer on duty here.' };
      return { ok: true, card: { name: g.Name, phone: g.Phone, photo: g.PhotoId, post: g.RegularPost, shift: g.RegularShift, agency: g.Agency } };
    }
  }
  return { ok: false, error: 'This link is not valid.' };
}

function addGuard_(q, by, role) {
  if (role === 'head') return { ok: false, error: 'Not allowed' };
  var name = String(q.name || '').trim(), phone = String(q.phone || '').replace(/\D/g, '');
  if (name.length < 3) return { ok: false, error: 'Enter the full name' };
  var shift = q.shift === 'Night' ? 'Night' : q.shift === 'Day' ? 'Day' : '';
  if (!shift) return { ok: false, error: 'Choose his shift: Day or Night' };
  var agency = q.agency === 'Avenger' ? 'Avenger' : q.agency === 'Inhouse' ? 'Inhouse' : 'Sainik';
  if (role === 'supervisor' && !q.photo) return { ok: false, error: 'Take a photo of the guard with the camera' };
  var gs = rows_('Guards');
  for (var i = 0; i < gs.length; i++) {
    if (gs[i].Name.toLowerCase() === name.toLowerCase() || (phone && gs[i].Phone === phone))
      return { ok: false, error: 'Already registered: ' + gs[i].Name };
  }
  var id = 'G' + ('000' + (gs.length + 1)).slice(-4);
  var g = { ID: id, Name: name, Phone: phone, PhotoId: photo_(q.photo, 'guard_' + id), AddedBy: by, AddedAt: now_(), RegularPost: String(q.post || ''), RegularShift: shift, Status: 'Active', Agency: agency };
  add_('Guards', g);
  return { ok: true, guard: guardOut_(g) };
}
function updateGuard_(q, by, role) {
  if (role === 'supervisor') return { ok: false, error: 'Only the Control Room or Head Office can change a guard\'s shift or agency' };
  var gs = rows_('Guards'), g = null;
  for (var i = 0; i < gs.length; i++) if (gs[i].ID === q.id) g = gs[i];
  if (!g) return { ok: false, error: 'Guard not found' };
  var upd = {};
  if (q.shift === 'Day' || q.shift === 'Night') upd.RegularShift = q.shift;
  if (q.agency === 'Sainik' || q.agency === 'Avenger' || q.agency === 'Inhouse') upd.Agency = q.agency;
  if (q.post !== undefined) upd.RegularPost = String(q.post);
  if (q.name !== undefined) { var nm = String(q.name).trim(); if (nm.length < 3) return { ok: false, error: 'Enter the full name' }; if (nm !== g.Name) upd.Name = nm; }
  if (q.phone !== undefined) upd.Phone = String(q.phone).replace(/\D/g, '');
  if (q.status === 'Left' || q.status === 'Active') upd.Status = q.status;
  if (q.photo) { var pid = photo_(q.photo, 'guard_' + g.ID); if (pid) upd.PhotoId = pid; }
  if (!Object.keys(upd).length) return { ok: false, error: 'Nothing to change' };
  setCells_('Guards', g._row, upd);
  addAlert_('guard', 'guard:' + g.ID + ':' + now_(), 'info', by + ' updated ' + g.Name + ': ' + Object.keys(upd).map(function (k) { return k === 'PhotoId' ? 'new photo' : k + ' = ' + upd[k]; }).join(', '), 'head');
  var ng = rows_('Guards').filter(function (x) { return x.ID === g.ID; })[0];
  return { ok: true, guard: guardOut_(ng) };
}

function submitCheck_(q, by, role) {
  if (role === 'head') return { ok: false, error: 'Not allowed' };
  var shift = q.shift === 'Night' ? 'Night' : 'Day', presence = String(q.presence || 'Present');
  var post = String(q.post || '').trim();
  if (!post) return { ok: false, error: 'Choose a post' };
  var type = String(q.type || 'Regular');
  var guard = null, flags = [], date = today_(), t = now_();
  if (q.guardId) {
    var gs = rows_('Guards');
    for (var i = 0; i < gs.length; i++) if (gs[i].ID === q.guardId) guard = gs[i];
    if (!guard) return { ok: false, error: 'Guard not found. Register him first.' };
  } else if (presence !== 'Absent') {
    return { ok: false, error: 'Choose the guard, or mark the post Absent' };
  }
  var onPost = presence !== 'Absent';
  if (role === 'supervisor' && onPost && !q.photo) return { ok: false, error: 'Take a photo at the post with the camera. It is required.' };
  var all = rows_('Checks'), hash = '', lat = '', lng = '', acc = '';
  /* 1. Is the photo fresh and unique? */
  if (q.photo) {
    var mm = /^data:image\/\w+;base64,(.+)$/.exec(q.photo);
    if (mm) {
      hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, Utilities.base64Decode(mm[1])).map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
      for (var h = all.length - 1; h >= 0 && h > all.length - 800; h--) {
        if (all[h].Hash && all[h].Hash === hash) { flags.push('SAME PHOTO REUSED: this photo was already used at ' + all[h].Post + ' on ' + all[h].Date + '. Not a fresh photo.'); break; }
      }
    }
  }
  /* 2. Was the supervisor physically at the gate? */
  if (role === 'supervisor' && onPost) {
    var la = Number(q.lat), lo = Number(q.lng);
    if (q.lat === undefined || q.lat === '' || isNaN(la) || isNaN(lo)) {
      flags.push('NO LOCATION: location was not shared, so presence at the gate cannot be confirmed.');
    } else {
      lat = la; lng = lo; acc = Math.round(Number(q.acc) || 0);
      var pl = [], po = [];
      all.forEach(function (c) { if (c.Post === post && c.Lat !== '' && c.Lng !== '' && c.Role === 'supervisor') { pl.push(Number(c.Lat)); po.push(Number(c.Lng)); } });
      if (pl.length >= 3) {
        var d = Math.round(distM_(la, lo, median_(pl), median_(po)));
        if (d > 150) flags.push('LOCATION: scanned ' + d + ' m away from where ' + post + ' is usually checked.');
      }
    }
  }
  /* 3. Is the guard on his own shift? */
  if (guard && onPost) {
    var seen = all.some(function (c) { return c.GuardId === guard.ID && c.Presence !== 'Absent'; });
    if (!seen) flags.push('NEW GUARD: first time on record. Ask his agency for his ID proof.');
    var df = dutyFlag_(guard, { Date: date, Shift: shift, Time: t, Presence: presence }, all);
    if (df) flags.push(df);
    if (guard.Agency !== 'Inhouse' && (guard.RegularShift !== shift || guard.RegularPost !== post)) {
      /* Agencies move their men between gates and shifts as they like; keep his usual post and shift up to date without a warning. */
      setCells_('Guards', guard._row, { RegularPost: post, RegularShift: shift }); guard.RegularShift = shift; guard.RegularPost = post;
    }
    if (!guard.RegularShift) flags.push('SHIFT NOT SET: this guard has no registered shift yet. Control Room should set it.');
  }
  if (type === 'Temporary' || type === 'Permanent') {
    if (!String(q.related || '').trim()) return { ok: false, error: 'Write whom he is replacing' };
  }
  if (type === 'Rotation' && !String(q.related || '').trim()) return { ok: false, error: 'Write where he was moved from' };
  if (presence === 'Absent') flags.push('ABSENT: no guard at ' + post + ' (' + shift + ')');
  if (presence === 'Late') flags.push('LATE arrival');
  if (type === 'Temporary') flags.push('Temporary replacement for ' + q.related);
  if (type === 'Permanent' && guard) {
    setCells_('Guards', guard._row, { RegularPost: post, RegularShift: shift });
    flags.push('Permanent replacement for ' + q.related + '. Regular post and shift updated.');
  }
  add_('Checks', {
    Time: t, Date: date, Shift: shift, Post: post, GuardId: guard ? guard.ID : '', Guard: guard ? guard.Name : '',
    Presence: presence, Type: type, Related: q.related || '', Remark: String(q.remark || '').slice(0, 500),
    PhotoId: photo_(q.photo, 'check'), By: by, Flags: flags.join(' | '), Role: role, Lat: lat, Lng: lng, Acc: acc, Hash: hash
  });
  flags.forEach(function (f) {
    var lvl = /^ABSENT/.test(f) ? 'high' : /^(DOUBLE|SAME PHOTO|NO LOCATION|LOCATION)/.test(f) ? 'medium' : '';
    if (lvl) addAlert_('scan', 'scan:' + t + ':' + post + ':' + f.slice(0, 12), lvl, post + ' (' + shift + '), ' + by + ': ' + f, 'control,head');
  });
  try { alertTick_(); } catch (e) {}
  return { ok: true, flags: flags };
}

function raiseIssue_(q, by, role) {
  var desc = String(q.description || '').trim();
  if (desc.length < 3) return { ok: false, error: 'Write what the issue is' };
  issueToCol_();
  var to = role === 'control' ? (q.to === 'supervisor' ? 'supervisor' : 'head') : 'control';
  var n = rows_('Issues').length + 1, id = 'I-' + ('0000' + n).slice(-4), t = now_();
  add_('Issues', { To: to,
    ID: id, Created: t, Post: q.post || '', Category: q.category || 'Other', Priority: q.priority === 'Urgent' ? 'Urgent' : 'Normal',
    Description: desc.slice(0, 800), PhotoId: photo_(q.photo, 'issue_' + id), RaisedBy: by, Status: 'Open', LastRemark: '', Updated: t, SolvedAt: ''
  });
  add_('Log', { Time: t, IssueID: id, By: by, Role: role, Action: 'Raised', Remark: desc.slice(0, 200) });
  addAlert_('issue', 'issue:' + id, q.priority === 'Urgent' ? 'high' : 'medium', (q.priority === 'Urgent' ? 'URGENT ' : '') + 'Issue ' + id + ' at ' + (q.post || '') + ': ' + (q.category || 'Other') + ' (raised by ' + by + ')', 'control,head');
  return { ok: true, id: id, to: to };
}
/** Adds the To header to the Issues sheet once (issues raised before it existed have it blank). */
function issueToCol_() {
  var sh = ss_().getSheetByName('Issues'), c = HEAD.Issues.length, r = sh.getRange(1, c);
  if (r.getValue() !== 'To') r.setValue('To').setFontWeight('bold').setBackground('#1F3A5F').setFontColor('#ffffff');
}
/** Who raised an issue (role) and whom it is for. Old issues: Control Room reports went to Head Office, the rest to the Control Room. */
function raisedRoles_(logs) { var o = {}; logs.forEach(function (l) { if (l.Action === 'Raised' && !o[l.IssueID]) o[l.IssueID] = l.Role; }); return o; }
function toOf_(r, rr) { return r.To || ((rr[r.ID] || '') === 'control' ? 'head' : 'control'); }

/** Camera sightings are for Control Room and Head Office only. */
var PRIVATE_CAT = 'Suspicious activity (seen on camera)';
function listIssues_(role, q) {
  var days = Math.min(Math.max(Number(q && q.days) || 7, 1), 62);
  issueToCol_();
  var logs = rows_('Log'), byId = {}, rr = raisedRoles_(logs), cutoff = new Date().getTime() - days * 24 * 3600 * 1000;
  logs.forEach(function (l) { (byId[l.IssueID] = byId[l.IssueID] || []).push({ time: l.Time, by: l.By, role: l.Role, action: l.Action, remark: l.Remark, photo: l.PhotoId }); });
  var out = rows_('Issues').filter(function (r) {
    if (role === 'supervisor' && (toOf_(r, rr) === 'head' || (r.Category === PRIVATE_CAT && toOf_(r, rr) !== 'supervisor'))) return false;
    return r.Status !== 'Solved' || new Date((r.SolvedAt || r.Updated).replace(' ', 'T') + '+05:30').getTime() >= cutoff;
  }).map(function (r) {
    return { id: r.ID, created: r.Created, post: r.Post, category: r.Category, priority: r.Priority, description: r.Description,
      photo: r.PhotoId, raisedBy: r.RaisedBy, status: r.Status, lastRemark: r.LastRemark, updated: r.Updated, solvedAt: r.SolvedAt,
      to: toOf_(r, rr), raisedRole: rr[r.ID] || '', log: (byId[r.ID] || []).slice(-8) };
  });
  return { ok: true, issues: out.slice(days > 7 ? -1500 : -300) };
}

function updateIssue_(q, by, role) {
  var op = String(q.op || ''), remark = String(q.remark || '').trim();
  var allowed = {
    supervisor: ['Comment', 'Acknowledged', 'Solved'],
    control: ['Acknowledged', 'In progress', 'Solved', 'Reopened', 'Comment'],
    head: ['Query', 'Reopened', 'Comment', 'Instruction', 'Noted']
  };
  if (allowed[role].indexOf(op) < 0) return { ok: false, error: 'Not allowed for your role' };
  if ((op === 'Solved' || op === 'Query' || op === 'Comment' || op === 'Instruction') && remark.length < 2) return { ok: false, error: 'Write a remark' };
  var rs = rows_('Issues'), is = null;
  for (var i = 0; i < rs.length; i++) if (rs[i].ID === q.id) is = rs[i];
  if (!is) return { ok: false, error: 'Issue not found' };
  var to = toOf_(is, raisedRoles_(rows_('Log')));
  if (role === 'supervisor' && (to === 'head' || (is.Category === PRIVATE_CAT && to !== 'supervisor'))) return { ok: false, error: 'Issue not found' };
  if (role === 'supervisor' && op !== 'Comment' && to !== 'supervisor') return { ok: false, error: 'Only the Control Room can change this issue' };
  if (op === 'Instruction' || op === 'Noted') {
    /* Head Office can reply to or close only reports the Control Room sent to Head Office. */
    if (to !== 'head') return { ok: false, error: 'Only for reports the Control Room sent to Head Office' };
    if (is.Status === 'Solved') return { ok: false, error: 'Already closed' };
    if (op === 'Noted' && !remark) remark = 'Noted by Head Office';
  }
  var t = now_(), upd = { Updated: t };
  if (op === 'Acknowledged' || op === 'In progress') upd.Status = op;
  if (op === 'Instruction') upd.Status = 'In progress';
  if (op === 'Solved' || op === 'Noted') { upd.Status = 'Solved'; upd.SolvedAt = t; }
  if (op === 'Reopened') { upd.Status = 'Open'; upd.SolvedAt = ''; }
  if (remark) upd.LastRemark = remark.slice(0, 300);
  setCells_('Issues', is._row, upd);
  add_('Log', { Time: t, IssueID: is.ID, By: by, Role: role, Action: op, Remark: remark.slice(0, 500), PhotoId: photo_(q.photo, 'log_' + is.ID) });
  return { ok: true };
}

function listChecks_(from, to) {
  var span = (new Date(to) - new Date(from)) / 864e5;
  if (isNaN(span) || span < 0 || span > 62) return { ok: false, error: 'Choose a period of up to 62 days' };
  var rs = rows_('Checks').filter(function (c) { return c.Date >= from && c.Date <= to; }).map(function (c) {
    return { time: c.Time, date: c.Date, shift: c.Shift, post: c.Post, guardId: c.GuardId, guard: c.Guard, presence: c.Presence,
      type: c.Type, related: c.Related, remark: c.Remark, photo: c.PhotoId, by: c.By, flags: c.Flags, role: c.Role, lat: c.Lat, lng: c.Lng, acc: c.Acc };
  });
  return { ok: true, from: from, to: to, checks: rs };
}

/* ================= Rounds and alerts ================= */
/** Round windows (IST). A round is complete when at least min_cover of the gates were checked by that role inside the window. */
var ROUNDS = [
  { id: 'S1', role: 'supervisor', name: 'Morning round', from: '09:00', to: '11:30' },
  { id: 'S2', role: 'supervisor', name: 'Day round (any time)', from: '11:30', to: '19:00' },
  { id: 'S4', role: 'supervisor', name: 'Night watchmen round (evening)', from: '19:30', to: '21:00', night: true },
  { id: 'C4', role: 'control', name: 'Control Room round (once a day, any time)', from: '00:00', to: '23:59' }
];
/* Gates with no watchman at night: not expected in night rounds or on the night board. */
var NIGHT_OFF = ['Gate 2', 'Sarela Makan'];
function requiredPosts_(night) { if (night) return requiredPosts_().filter(function (p) { return NIGHT_OFF.indexOf(p) < 0; }); return rows_('Posts').map(function (r) { return r.Post; }).filter(function (p) { return p && p !== 'Control Room'; }); }
function dateAdd_(d, n) { var x = new Date(d + 'T12:00:00+05:30'); x.setDate(x.getDate() + n); return Utilities.formatDate(x, TZ, 'yyyy-MM-dd'); }
/** Status of every round window on the given dates, newest first. */
function roundList_(dates, checks, rm) {
  var firstT = Infinity; rows_('Checks').forEach(function (c) { var x = ts_(c.Time); if (x < firstT) firstT = x; });
  var minCover = Number(settings_().min_cover) || 0.9, req = requiredPosts_(), now = new Date().getTime(), out = [];
  dates.forEach(function (d) {
    ROUNDS.forEach(function (R) {
      var st = ts_(d + ' ' + R.from + ':00'), en = ts_(d + ' ' + R.to + ':00'), lim = en + (R.grace || 0) * 60000, posts = {}, late = false;
      if (en < firstT) return;
      checks.forEach(function (c) {
        if ((c.Role || roleOfName_(c.By, rm)) !== R.role) return;
        var t = ts_(c.Time); if (t >= st && t <= lim) { posts[c.Post] = 1; if (t > en) late = true; }
      });
      var rq = R.night ? req.filter(function (p) { return NIGHT_OFF.indexOf(p) < 0; }) : req;
      var got = rq.filter(function (p) { return posts[p]; }), status = now < st ? 'upcoming' : now <= en ? 'open' : (got.length >= Math.ceil(rq.length * minCover) ? 'done' : (now <= lim ? 'open' : 'missed'));
      out.push({ id: R.id, role: R.role, name: R.name, date: d, from: R.from, to: R.to, status: status, late: late && status === 'done', done: got.length, total: rq.length,
        missing: rq.filter(function (p) { return !posts[p]; }), start: st });
    });
  });
  out.sort(function (a, b) { return b.start - a.start; });
  out.forEach(function (r) { delete r.start; });
  return out;
}
function roundsOut_(role, q) {
  var days = Math.min(Math.max(Number(q && q.days) || 1, 1), 31), dates = [], today = today_();
  for (var i = 0; i < days; i++) dates.push(dateAdd_(today, -i));
  var from = dateAdd_(today, -days), checks = rows_('Checks').filter(function (c) { return c.Date >= from; });
  var list = roundList_(dates, checks, roleMap_());
  /* Head Office sees all rounds; the Control Room also sees the supervisor's rounds; the supervisor sees only his own. */
  if (role === 'supervisor') list = list.filter(function (r) { return r.role === role; });
  return { ok: true, rounds: list, minCover: Number(settings_().min_cover) || 0.9 };
}

function addAlert_(type, key, level, msg, aud) {
  var sh = ss_().getSheetByName('Alerts'); if (!sh) return false;
  var last = sh.getLastRow(), from = Math.max(2, last - 400);
  if (last >= 2 && sh.getRange(from, 3, last - from + 1, 1).getValues().some(function (r) { return String(r[0]) === key; })) return false;
  add_('Alerts', { Time: now_(), Type: type, Key: key, Level: level, Message: msg, Audience: aud });
  if (level === 'high' && type === 'issue' && /^(unack|overdue):/.test(key)) emailAlert_(msg); // email only unsolved issues, not rounds or scans
  return true;
}
/** Emails important alerts to the address(es) in Settings > alert_email. Never lets an email problem break the app. */
function emailAlert_(msg) {
  try {
    var st = settings_(), to = String(st.alert_email || '').trim();
    if (!to) return;
    var day = today_(), cnt = (st.mail_day === day ? Number(st.mail_count) || 0 : 0);
    if (cnt >= 40) return;
    MailApp.sendEmail({ to: to, subject: 'Security alert: ' + msg.slice(0, 90), body: msg + '\n\nOpen the Security Control app to see everything:\nhttps://urzhusain.github.io/gate-rosters/supervisor/\n\n(Automatic message from Aljamea Security Control.)' });
    setSetting_('mail_day', day); setSetting_('mail_count', String(cnt + 1));
  } catch (e) {}
}
/** Run ONCE from the editor. Approves email + timer permission, sets the alert address and starts the 10-minute check. */
function installAlerts() {
  if (!settings_().alert_email) setSetting_('alert_email', 'urzhusain@gmail.com');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'alertTickAuto') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('alertTickAuto').timeBased().everyMinutes(10).create();
  MailApp.sendEmail({ to: settings_().alert_email, subject: 'Security alerts are switched on', body: 'This is a test message. From now on you will get an email here when a round is missed, an issue is not acknowledged or is overdue, or a guard is found absent.\n\nApp: https://urzhusain.github.io/gate-rosters/supervisor/' });
  return 'ok';
}
/** Runs every 10 minutes by itself, even when nobody has the app open. */
/* ================= Checkpoints (timed duties at a gate) ================= */
/** A checkpoint is a duty that must be confirmed once a day inside a time window. type 'photo' = photo of the register or gate; type 'check' = spot check, Matches or Problem. */
var CHECKPOINTS = [
  { id: 'R1', by: 'supervisor', post: 'Gate 1', name: 'Register photo (visitors, deliveries, outgoing material)', type: 'photo', from: '17:00', to: '20:45' },
  { id: 'R2', by: 'supervisor', post: 'Gate 4', name: 'Register photo (visitors, deliveries, outgoing material)', type: 'photo', from: '17:00', to: '20:45' },
  { id: 'R3', by: 'supervisor', post: 'Gate 8 & 9', name: 'Register photo (visitors, deliveries, outgoing material)', type: 'photo', from: '17:00', to: '20:45' },
  { id: 'M1', post: 'Gate 8 & 9', name: 'Incoming material vs bill', type: 'check', from: '09:00', to: '19:00', hint: 'Take the latest incoming entry. Check the goods against the bill. Does the guard\'s tally match?' },
  { id: 'M2', post: 'Gate 8 & 9', name: 'Outgoing material vs gate pass', type: 'check', from: '09:00', to: '19:00', hint: 'Take the latest outgoing entry. Check the goods against the gate pass. Does the guard\'s tally match?' }
];
function cpSheet_() {
  var ss = ss_(), sh = ss.getSheetByName('Checkpoints');
  if (!sh) {
    sh = ss.insertSheet('Checkpoints'); var cols = HEAD.Checkpoints.length;
    sh.getRange(1, 1, sh.getMaxRows(), cols).setNumberFormat('@');
    sh.getRange(1, 1, 1, cols).setValues([HEAD.Checkpoints]).setFontWeight('bold').setBackground('#1F3A5F').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}
function cpRows_() { cpSheet_(); return rows_('Checkpoints'); }
/** Status of every checkpoint on the given dates, newest first. */
function cpList_(dates, recs) {
  var now = new Date().getTime(), out = [];
  dates.forEach(function (d) {
    CHECKPOINTS.forEach(function (C) {
      var st = ts_(d + ' ' + C.from + ':00'), en = ts_(d + ' ' + C.to + ':00'), mine = recs.filter(function (r) { return r.Date === d && r.Id === C.id; });
      var rec = mine.length ? mine[mine.length - 1] : null, status;
      if (rec) status = rec.Result === 'Problem' ? 'problem' : (ts_(rec.Time) > en ? 'late' : 'done');
      else status = now < st ? 'upcoming' : now <= en ? 'open' : 'missed';
      out.push({ id: C.id, hint: C.hint || '', by_role: C.by || '', post: C.post, name: C.name, type: C.type, date: d, from: C.from, to: C.to, status: status,
        by: rec ? rec.By : '', time: rec ? rec.Time : '', result: rec ? rec.Result : '', remark: rec ? rec.Remark : '', photo: rec ? rec.PhotoId : '' });
    });
  });
  out.sort(function (a, b) { return a.date === b.date ? (a.from < b.from ? -1 : a.from > b.from ? 1 : 0) : (a.date < b.date ? 1 : -1); });
  return out;
}
function cpOut_(q) {
  var days = Math.min(Math.max(Number(q && q.days) || 1, 1), 14), dates = [], t = today_();
  for (var i = 0; i < days; i++) dates.push(dateAdd_(t, -i));
  return { ok: true, checkpoints: cpList_(dates, cpRows_().filter(function (r) { return r.Date >= dates[dates.length - 1]; })) };
}
function cpSubmit_(q, by, role) {
  if (role !== 'supervisor' && role !== 'control') return { ok: false, error: 'Only the supervisor or the Control Room can confirm' };
  var C = CHECKPOINTS.filter(function (x) { return x.id === q.id; })[0];
  if (!C) return { ok: false, error: 'Unknown checkpoint' };
  if (C.by && C.by !== role) return { ok: false, error: 'This checkpoint is confirmed by the ' + (C.by === 'control' ? 'Control Room' : C.by) + '.' };
  var t = today_(), now = new Date().getTime();
  if (now < ts_(t + ' ' + C.from + ':00')) return { ok: false, error: 'Not open yet. It opens at ' + C.from + '.' };
  var result = C.type === 'check' ? (q.result === 'Problem' ? 'Problem' : 'Matches') : 'Done';
  if (C.type === 'photo' && role === 'supervisor' && !q.photo) return { ok: false, error: 'Take the photo with the camera. It is required.' };
  if (result === 'Problem' && String(q.remark || '').trim().length < 3) return { ok: false, error: 'Write what the problem is' };
  var photo = photo_(q.photo, 'cp_' + C.id);
  cpSheet_();
  add_('Checkpoints', { Time: now_(), Date: t, Id: C.id, Post: C.post, Result: result, Remark: String(q.remark || '').slice(0, 300), PhotoId: photo, By: by, Role: role });
  if (result === 'Problem') raiseIssue_({ post: C.post, category: 'Gate pass / material', priority: 'Normal', description: C.name + ': ' + String(q.remark).trim(), photo: q.photo }, by, role);
  return { ok: true };
}

/* ================= Daily summary email (9 AM) ================= */
function digestText_() {
  var t = today_(), y = dateAdd_(t, -1), st = settings_(), rm = roleMap_(), req = requiredPosts_();
  var all = rows_('Checks'), from = dateAdd_(t, -15);
  var checks = all.filter(function (c) { return c.Date === y || c.Date === t; });
  var dayC = checks.filter(function (c) { return c.Date === y && c.Shift === 'Day'; });
  var nightC = checks.filter(function (c) { return c.Shift === 'Night' && ((c.Date === y && clockMin_(c.Time) >= 720) || (c.Date === t && clockMin_(c.Time) < 720)); });
  var nightPosts = {}; all.forEach(function (c) { if (c.Date >= from && c.Shift === 'Night' && c.Post !== 'Control Room') nightPosts[c.Post] = 1; });
  var dayPosts = {}; all.forEach(function (c) { if (c.Date >= from && c.Shift === 'Day' && c.Post !== 'Control Room') dayPosts[c.Post] = 1; });
  var seenPost = function (list, p) { return list.some(function (c) { return c.Post === p; }); };
  var noDay = req.filter(function (p) { return dayPosts[p] && !seenPost(dayC, p); }), noNight = req.filter(function (p) { return nightPosts[p] && NIGHT_OFF.indexOf(p) < 0 && !seenPost(nightC, p); });
  var both = dayC.concat(nightC);
  var absent = both.filter(function (c) { return c.Presence === 'Absent'; }), late = both.filter(function (c) { return c.Presence === 'Late'; });
  var newF = both.filter(function (c) { return /NEW GUARD/.test(c.Flags || ''); }), dbl = both.filter(function (c) { return /DOUBLE/.test(c.Flags || ''); });
  var L = [];
  L.push('SECURITY SUMMARY for ' + y + ' (day shift) and the night that followed');
  L.push('');
  L.push('GUARDS');
  L.push('- Absent: ' + absent.length + (absent.length ? ' -> ' + absent.map(function (c) { return (c.Guard || 'no name') + ' (' + c.Post + ', ' + c.Shift + ')'; }).join('; ') : ''));
  L.push('- Late: ' + late.length + (late.length ? ' -> ' + late.map(function (c) { return (c.Guard || 'no name') + ' (' + c.Post + ')'; }).join('; ') : ''));
  L.push('- New faces: ' + newF.length + (newF.length ? ' -> ' + newF.map(function (c) { return (c.Guard || 'no name') + ' (' + c.Post + ', ' + c.Shift + ')'; }).join('; ') : ''));
  L.push('- Double duty: ' + dbl.length + (dbl.length ? ' -> ' + dbl.map(function (c) { return (c.Guard || 'no name') + ' (' + c.Post + ')'; }).join('; ') : ''));
  L.push('- Gates not checked, day shift: ' + (noDay.length ? noDay.join(', ') : 'none'));
  L.push('- Gates not checked, night shift: ' + (noNight.length ? noNight.join(', ') : 'none'));
  L.push('');
  var rl = roundList_([y, t], all.filter(function (c) { return c.Date >= dateAdd_(y, -1); }), rm).filter(function (r) { return r.date === y || r.id === 'C4'; }).reverse();
  var cpm = cpList_([y], cpRows_().filter(function (c) { return c.Date === y; })).filter(function (c) { return c.status === 'missed' || c.status === 'problem'; });
  L.push('CHECKPOINTS');
  L.push('- ' + (cpm.length ? cpm.map(function (c) { return c.name + ' at ' + c.post + ': ' + (c.status === 'missed' ? 'MISSED' : 'PROBLEM'); }).join('; ') : 'All confirmed yesterday'));
  L.push('');
  L.push('ROUNDS');
  rl.forEach(function (r) {
    var who = r.role === 'control' ? 'Control Room' : 'Supervisor';
    L.push('- ' + (r.date === t ? 'This morning' : 'Yesterday') + ' ' + r.name + ' (' + who + ', ' + r.from + ' to ' + r.to + '): ' + (r.status === 'done' ? 'DONE' : r.status === 'missed' ? 'MISSED' : r.status) + ', ' + r.done + ' of ' + r.total + ' gates' + (r.status === 'missed' && r.missing.length ? '. Not checked: ' + r.missing.join(', ') : ''));
  });
  L.push('');
  var issues = rows_('Issues'), logs = rows_('Log'), od = Number(st.overdue_hours) || 24, nowT = new Date().getTime();
  var raised = issues.filter(function (i) { return String(i.Created).slice(0, 10) === y; }), solvedY = issues.filter(function (i) { return i.Status === 'Solved' && String(i.SolvedAt).slice(0, 10) === y; });
  var open = issues.filter(function (i) { return i.Status !== 'Solved'; });
  var resp = [];
  raised.forEach(function (i) {
    if (i.Category === PRIVATE_CAT) return;
    var first = logs.filter(function (l) { return l.IssueID === i.ID && l.Action !== 'Raised' && l.Role === 'control'; })[0];
    if (first) resp.push((ts_(first.Time) - ts_(i.Created)) / 60000);
  });
  L.push('ISSUES');
  L.push('- Raised yesterday: ' + raised.length + '. Solved yesterday: ' + solvedY.length + '. Open now: ' + open.length + '.');
  if (resp.length) L.push('- Control Room first response: average ' + Math.round(resp.reduce(function (a, b) { return a + b; }, 0) / resp.length) + ' minutes (' + resp.length + ' issues).');
  var unack = open.filter(function (i) { return i.Status === 'Open'; });
  if (unack.length) L.push('- Not yet acknowledged: ' + unack.map(function (i) { return i.ID + ' (' + i.Post + ')'; }).join(', '));
  var over = open.filter(function (i) { return (nowT - ts_(i.Created)) / 3600000 >= od; });
  if (over.length) L.push('- Unsolved for more than ' + od + ' hours: ' + over.map(function (i) { return i.ID + ' ' + i.Post + ', ' + i.Category + ' (' + Math.round((nowT - ts_(i.Created)) / 3600000) + ' h)'; }).join('; '));
  if (!unack.length && !over.length) L.push('- Nothing waiting for acknowledgement and nothing overdue.');
  L.push('');
  L.push('Open the app: https://urzhusain.github.io/gate-rosters/supervisor/');
  L.push('(Automatic daily message from Aljamea Security Control.)');
  return { subject: 'Security summary ' + y + ': ' + absent.length + ' absent, ' + (rl.filter(function (r) { return r.status === 'missed'; }).length) + ' rounds missed, ' + open.length + ' issues open', body: L.join('\n') };
}
function dailyDigest() {
  var st = settings_(), to = String(st.alert_email || '').trim();
  if (!to) return 'no email set';
  var d = digestText_();
  MailApp.sendEmail({ to: to, subject: d.subject, body: d.body });
  return 'sent';
}
/** Run once in the editor: sends the summary every day at about 9 AM (India time). */
function installDigest() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyDigest') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyDigest').timeBased().atHour(9).nearMinute(0).everyDays(1).inTimezone(TZ).create();
  return dailyDigest();
}
function alertTickAuto() { alertTick_(); }
/** Looks for problems nobody has reported: missed rounds, unacknowledged and overdue issues. Runs at most every 4 minutes, whenever the app is used. */
function alertTick_() {
  var st = settings_(), nowT = new Date().getTime();
  if (st.last_tick && nowT - Number(st.last_tick) < 4 * 60000) return;
  setSetting_('last_tick', String(nowT));
  var today = today_(), rm = roleMap_(), from = dateAdd_(today, -2);
  var list = roundList_([today, dateAdd_(today, -1)], rows_('Checks').filter(function (c) { return c.Date >= from; }), rm);
  list.forEach(function (r) {
    if (r.status !== 'missed' || nowT - ts_(r.date + ' ' + r.to + ':00') > 6 * 3600000) return;
    addAlert_('round', 'round:' + r.id + ':' + r.date, 'high', 'Missed round: ' + r.name + ' (' + (r.role === 'control' ? 'Control Room' : 'Supervisor') + '), ' + r.date + ' ' + r.from + ' to ' + r.to + '. Only ' + r.done + ' of ' + r.total + ' gates checked. Not checked: ' + r.missing.slice(0, 6).join(', ') + (r.missing.length > 6 ? '…' : ''), 'control,head');
  });
  try {
    cpList_([today, dateAdd_(today, -1)], cpRows_().filter(function (c) { return c.Date >= from; })).forEach(function (c) {
      if (c.status !== 'missed' || nowT - ts_(c.date + ' ' + c.to + ':00') > 6 * 3600000) return;
      addAlert_('checkpoint', 'cp:' + c.id + ':' + c.date, 'medium', 'Checkpoint missed: ' + c.name + ' at ' + c.post + ' (' + c.date + ', by ' + c.to + ')', 'control,head');
    });
  } catch (e) {}
  var od = Number(st.overdue_hours) || 24;
  rows_('Issues').forEach(function (i) {
    if (i.Status === 'Solved') return;
    var ageH = (nowT - ts_(i.Created)) / 36e5;
    if (i.Status === 'Open' && i.Category !== PRIVATE_CAT && ageH >= 0.5) addAlert_('issue', 'unack:' + i.ID, 'high', 'Nobody has acknowledged ' + i.ID + ' (' + i.Post + ', ' + i.Category + ') for ' + Math.round(ageH * 60) + ' minutes.', 'control,head');
    if (ageH >= od) addAlert_('issue', 'overdue:' + i.ID + ':' + today, 'high', i.ID + ' (' + i.Post + ', ' + i.Category + ') is still unsolved after ' + Math.round(ageH) + ' hours.', 'control,head');
  });
}
function alertsOut_(role) {
  try { alertTick_(); } catch (e) {}
  if (role === 'supervisor') return { ok: true, alerts: [] };
  var out = rows_('Alerts').filter(function (a) { return String(a.Audience).split(',').indexOf(role) > -1; }).map(function (a) {
    return { time: a.Time, type: a.Type, level: a.Level, message: a.Message };
  }).reverse().slice(0, 60);
  return { ok: true, alerts: out };
}
