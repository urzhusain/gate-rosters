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
  Issues: ['ID', 'Created', 'Post', 'Category', 'Priority', 'Description', 'PhotoId', 'RaisedBy', 'Status', 'LastRemark', 'Updated', 'SolvedAt'],
  Log: ['Time', 'IssueID', 'By', 'Role', 'Action', 'Remark', 'PhotoId'],
  Settings: ['Key', 'Value'],
  Users: ['Name', 'PIN', 'Role', 'Active'],
  Alerts: ['Time', 'Type', 'Key', 'Level', 'Message', 'Audience']
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
function doubleDutyMsg_(g, time) {
  return 'DOUBLE DUTY: ' + g.Name + ' is registered for the ' + g.RegularShift + ' shift but was on post at ' + String(time).slice(11, 16).replace(/^(\d):/, '0$1:') + ' (' + shiftByClock_(time) + ' hours).';
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
  var iF = HEAD.Checks.indexOf('Flags') + 1, n = 0;
  rows_('Checks').forEach(function (c) {
    var parts = String(c.Flags || '').split(' | ').filter(function (x) { return x && x.indexOf('DOUBLE DUTY') !== 0; });
    var g = gmap[c.GuardId];
    if (g && c.Presence !== 'Absent' && g.RegularShift && outsideShift_(g.RegularShift, c.Time)) parts.push(doubleDutyMsg_(g, c.Time));
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
    var q = JSON.parse(e.postData.contents), user = userOf_(q.pin);
    if (!user) return json_({ ok: false, error: 'Wrong PIN' });
    var role = user.role, by = user.name;
    switch (q.action) {
      case 'bootstrap': return json_(bootstrap_(role, by));
      case 'addGuard': return json_(addGuard_(q, by, role));
      case 'updateGuard': return json_(updateGuard_(q, by, role));
      case 'rounds': return json_(roundsOut_(role, q));
      case 'alerts': return json_(alertsOut_(role));
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
    ok: true, role: role, name: name, overdueHours: Number(s.overdue_hours) || 24, rounds: ROUNDS, minCover: Number(s.min_cover) || 0.9,
    posts: rows_('Posts').map(function (r) { return r.Post; }).filter(String),
    guards: rows_('Guards').filter(function (g) { return g.Status !== 'Left'; }).map(guardOut_)
  };
}
function guardOut_(g) {
  return { id: g.ID, name: g.Name, phone: g.Phone, photo: g.PhotoId, post: g.RegularPost, shift: g.RegularShift, agency: g.Agency };
}

function addGuard_(q, by, role) {
  if (role === 'head') return { ok: false, error: 'Not allowed' };
  var name = String(q.name || '').trim(), phone = String(q.phone || '').replace(/\D/g, '');
  if (name.length < 3) return { ok: false, error: 'Enter the full name' };
  var shift = q.shift === 'Night' ? 'Night' : q.shift === 'Day' ? 'Day' : '';
  if (!shift) return { ok: false, error: 'Choose his shift: Day or Night' };
  var agency = q.agency === 'Avenger' ? 'Avenger' : 'Sainik';
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
  if (q.agency === 'Sainik' || q.agency === 'Avenger') upd.Agency = q.agency;
  if (q.post !== undefined) upd.RegularPost = String(q.post);
  if (q.phone !== undefined) upd.Phone = String(q.phone).replace(/\D/g, '');
  if (q.status === 'Left' || q.status === 'Active') upd.Status = q.status;
  setCells_('Guards', g._row, upd);
  addAlert_('guard', 'guard:' + g.ID + ':' + now_(), 'info', by + ' updated ' + g.Name + ': ' + Object.keys(upd).map(function (k) { return k + ' = ' + upd[k]; }).join(', '), 'head');
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
    if (guard.RegularShift && outsideShift_(guard.RegularShift, t)) flags.push(doubleDutyMsg_(guard, t));
    else if (!guard.RegularShift) flags.push('SHIFT NOT SET: this guard has no registered shift yet. Control Room should set it.');
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
  var n = rows_('Issues').length + 1, id = 'I-' + ('0000' + n).slice(-4), t = now_();
  add_('Issues', {
    ID: id, Created: t, Post: q.post || '', Category: q.category || 'Other', Priority: q.priority === 'Urgent' ? 'Urgent' : 'Normal',
    Description: desc.slice(0, 800), PhotoId: photo_(q.photo, 'issue_' + id), RaisedBy: by, Status: 'Open', LastRemark: '', Updated: t, SolvedAt: ''
  });
  add_('Log', { Time: t, IssueID: id, By: by, Role: role, Action: 'Raised', Remark: desc.slice(0, 200) });
  addAlert_('issue', 'issue:' + id, q.priority === 'Urgent' ? 'high' : 'medium', (q.priority === 'Urgent' ? 'URGENT ' : '') + 'Issue ' + id + ' at ' + (q.post || '') + ': ' + (q.category || 'Other') + ' (raised by ' + by + ')', 'control,head');
  return { ok: true, id: id };
}

/** Camera sightings are for Control Room and Head Office only. */
var PRIVATE_CAT = 'Suspicious activity (seen on camera)';
function listIssues_(role, q) {
  var days = Math.min(Math.max(Number(q && q.days) || 7, 1), 62);
  var logs = rows_('Log'), byId = {}, cutoff = new Date().getTime() - days * 24 * 3600 * 1000;
  logs.forEach(function (l) { (byId[l.IssueID] = byId[l.IssueID] || []).push({ time: l.Time, by: l.By, role: l.Role, action: l.Action, remark: l.Remark, photo: l.PhotoId }); });
  var out = rows_('Issues').filter(function (r) {
    if (role === 'supervisor' && r.Category === PRIVATE_CAT) return false;
    return r.Status !== 'Solved' || new Date((r.SolvedAt || r.Updated).replace(' ', 'T') + '+05:30').getTime() >= cutoff;
  }).map(function (r) {
    return { id: r.ID, created: r.Created, post: r.Post, category: r.Category, priority: r.Priority, description: r.Description,
      photo: r.PhotoId, raisedBy: r.RaisedBy, status: r.Status, lastRemark: r.LastRemark, updated: r.Updated, solvedAt: r.SolvedAt,
      log: (byId[r.ID] || []).slice(-8) };
  });
  return { ok: true, issues: out.slice(days > 7 ? -1500 : -300) };
}

function updateIssue_(q, by, role) {
  var op = String(q.op || ''), remark = String(q.remark || '').trim();
  var allowed = {
    supervisor: ['Comment'],
    control: ['Acknowledged', 'In progress', 'Solved', 'Reopened', 'Comment'],
    head: ['Query', 'Reopened', 'Comment']
  };
  if (allowed[role].indexOf(op) < 0) return { ok: false, error: 'Not allowed for your role' };
  if ((op === 'Solved' || op === 'Query' || op === 'Comment') && remark.length < 2) return { ok: false, error: 'Write a remark' };
  var rs = rows_('Issues'), is = null;
  for (var i = 0; i < rs.length; i++) if (rs[i].ID === q.id) is = rs[i];
  if (!is) return { ok: false, error: 'Issue not found' };
  if (role === 'supervisor' && is.Category === PRIVATE_CAT) return { ok: false, error: 'Issue not found' };
  var t = now_(), upd = { Updated: t };
  if (op === 'Acknowledged' || op === 'In progress') upd.Status = op;
  if (op === 'Solved') { upd.Status = 'Solved'; upd.SolvedAt = t; }
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
  { id: 'S1', role: 'supervisor', name: 'Morning round', from: '09:30', to: '11:30' },
  { id: 'S2', role: 'supervisor', name: 'Midday round', from: '12:30', to: '14:30' },
  { id: 'S3', role: 'supervisor', name: 'Afternoon round', from: '15:30', to: '17:30' },
  { id: 'S4', role: 'supervisor', name: 'Night watchmen round (evening)', from: '19:30', to: '21:00' },
  { id: 'C4', role: 'control', name: 'Night watchmen round (morning)', from: '06:30', to: '08:00' }
];
function requiredPosts_() { return rows_('Posts').map(function (r) { return r.Post; }).filter(function (p) { return p && p !== 'Control Room'; }); }
function dateAdd_(d, n) { var x = new Date(d + 'T12:00:00+05:30'); x.setDate(x.getDate() + n); return Utilities.formatDate(x, TZ, 'yyyy-MM-dd'); }
/** Status of every round window on the given dates, newest first. */
function roundList_(dates, checks, rm) {
  var firstT = Infinity; rows_('Checks').forEach(function (c) { var x = ts_(c.Time); if (x < firstT) firstT = x; });
  var minCover = Number(settings_().min_cover) || 0.9, req = requiredPosts_(), now = new Date().getTime(), out = [];
  dates.forEach(function (d) {
    ROUNDS.forEach(function (R) {
      var st = ts_(d + ' ' + R.from + ':00'), en = ts_(d + ' ' + R.to + ':00'), posts = {};
      if (en < firstT) return;
      checks.forEach(function (c) {
        if ((c.Role || roleOfName_(c.By, rm)) !== R.role) return;
        var t = ts_(c.Time); if (t >= st && t <= en) posts[c.Post] = 1;
      });
      var got = req.filter(function (p) { return posts[p]; }), status = now < st ? 'upcoming' : now <= en ? 'open' : (got.length >= Math.ceil(req.length * minCover) ? 'done' : 'missed');
      out.push({ id: R.id, role: R.role, name: R.name, date: d, from: R.from, to: R.to, status: status, done: got.length, total: req.length,
        missing: req.filter(function (p) { return !posts[p]; }), start: st });
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
  if (role !== 'head') list = list.filter(function (r) { return r.role === role; });
  return { ok: true, rounds: list, minCover: Number(settings_().min_cover) || 0.9 };
}

function addAlert_(type, key, level, msg, aud) {
  var sh = ss_().getSheetByName('Alerts'); if (!sh) return false;
  var last = sh.getLastRow(), from = Math.max(2, last - 400);
  if (last >= 2 && sh.getRange(from, 3, last - from + 1, 1).getValues().some(function (r) { return String(r[0]) === key; })) return false;
  add_('Alerts', { Time: now_(), Type: type, Key: key, Level: level, Message: msg, Audience: aud });
  return true;
}
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
