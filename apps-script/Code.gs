/**
 * Aljamea Security Control: backend.
 * Lives in a Google Sheet you own. Run setup() once, then deploy as a Web app.
 * Sheets: Posts, Guards, Checks, Issues, Log, Settings, Users (one row per person: Name, PIN, Role, Active). Photos go to a Drive folder.
 */
var TZ = 'Asia/Kolkata';
var HEAD = {
  Posts: ['Post'],
  Guards: ['ID', 'Name', 'Phone', 'PhotoId', 'AddedBy', 'AddedAt', 'RegularPost', 'RegularShift', 'Status'],
  Checks: ['Time', 'Date', 'Shift', 'Post', 'GuardId', 'Guard', 'Presence', 'Type', 'Related', 'Remark', 'PhotoId', 'By', 'Flags'],
  Issues: ['ID', 'Created', 'Post', 'Category', 'Priority', 'Description', 'PhotoId', 'RaisedBy', 'Status', 'LastRemark', 'Updated', 'SolvedAt'],
  Log: ['Time', 'IssueID', 'By', 'Role', 'Action', 'Remark', 'PhotoId'],
  Settings: ['Key', 'Value'],
  Users: ['Name', 'PIN', 'Role', 'Active']
};
var POSTS = ['Gate 1', 'Gate 2', 'Gate 4', 'Gate 6 (Rabvat)', 'Gate 8 & 9', 'MVD Gate 10', 'Mahad Gate',
  'Masakin front gate', 'Masakin back gate', 'Raj Manzil', 'Yusufi Makan', 'Ezzy Makan', 'Funun-ul-Quran',
  'Sarela Makan', 'Nuzul', 'Utility', 'Control Room'];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
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
      case 'submitCheck': return json_(submitCheck_(q, by, role));
      case 'raiseIssue': return json_(raiseIssue_(q, by, role));
      case 'issues': return json_(listIssues_());
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
  return {
    ok: true, role: role, name: name, overdueHours: Number(s.overdue_hours) || 24,
    posts: rows_('Posts').map(function (r) { return r.Post; }).filter(String),
    guards: rows_('Guards').filter(function (g) { return g.Status !== 'Left'; }).map(guardOut_)
  };
}
function guardOut_(g) {
  return { id: g.ID, name: g.Name, phone: g.Phone, photo: g.PhotoId, post: g.RegularPost, shift: g.RegularShift };
}

function addGuard_(q, by, role) {
  if (role === 'head') return { ok: false, error: 'Not allowed' };
  var name = String(q.name || '').trim(), phone = String(q.phone || '').replace(/\D/g, '');
  if (name.length < 3) return { ok: false, error: 'Enter the full name' };
  var gs = rows_('Guards');
  for (var i = 0; i < gs.length; i++) {
    if (gs[i].Name.toLowerCase() === name.toLowerCase() || (phone && gs[i].Phone === phone))
      return { ok: false, error: 'Already registered: ' + gs[i].Name };
  }
  var id = 'G' + ('000' + (gs.length + 1)).slice(-4);
  var g = { ID: id, Name: name, Phone: phone, PhotoId: photo_(q.photo, 'guard_' + id), AddedBy: by, AddedAt: now_(), RegularPost: '', RegularShift: '', Status: 'Active' };
  add_('Guards', g);
  return { ok: true, guard: guardOut_(g) };
}

function submitCheck_(q, by, role) {
  if (role === 'head') return { ok: false, error: 'Not allowed' };
  var shift = q.shift === 'Night' ? 'Night' : 'Day', presence = String(q.presence || 'Present');
  var post = String(q.post || '').trim();
  if (!post) return { ok: false, error: 'Choose a post' };
  var type = String(q.type || 'Regular');
  var guard = null, flags = [], date = today_();
  if (q.guardId) {
    var gs = rows_('Guards');
    for (var i = 0; i < gs.length; i++) if (gs[i].ID === q.guardId) guard = gs[i];
    if (!guard) return { ok: false, error: 'Guard not found. Register him first.' };
  } else if (presence !== 'Absent') {
    return { ok: false, error: 'Choose the guard, or mark the post Absent' };
  }
  if (guard && presence !== 'Absent') {
    var all = rows_('Checks'), cutoff = new Date().getTime() - 24 * 3600 * 1000, seen = false;
    for (var k = 0; k < all.length; k++) {
      var c = all[k];
      if (c.GuardId !== guard.ID || c.Presence === 'Absent') continue;
      seen = true;
      var t = new Date(c.Time.replace(' ', 'T') + '+05:30').getTime();
      if (t >= cutoff && (c.Date !== date || c.Shift !== shift)) flags.push('DOUBLE DUTY: ' + guard.Name + ' was also at ' + c.Post + ' (' + c.Shift + ', ' + c.Date + ')');
    }
    if (!seen) flags.push('NEW GUARD: first time on record. Ask Sainik for his ID proof.');
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
    flags.push('Permanent replacement for ' + q.related + '. Regular post updated.');
  }
  add_('Checks', {
    Time: now_(), Date: date, Shift: shift, Post: post, GuardId: guard ? guard.ID : '', Guard: guard ? guard.Name : '',
    Presence: presence, Type: type, Related: q.related || '', Remark: String(q.remark || '').slice(0, 500),
    PhotoId: photo_(q.photo, 'check'), By: by, Flags: flags.join(' | ')
  });
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
  return { ok: true, id: id };
}

function listIssues_() {
  var logs = rows_('Log'), byId = {}, cutoff = new Date().getTime() - 7 * 24 * 3600 * 1000;
  logs.forEach(function (l) { (byId[l.IssueID] = byId[l.IssueID] || []).push({ time: l.Time, by: l.By, role: l.Role, action: l.Action, remark: l.Remark, photo: l.PhotoId }); });
  var out = rows_('Issues').filter(function (r) {
    return r.Status !== 'Solved' || new Date((r.SolvedAt || r.Updated).replace(' ', 'T') + '+05:30').getTime() >= cutoff;
  }).map(function (r) {
    return { id: r.ID, created: r.Created, post: r.Post, category: r.Category, priority: r.Priority, description: r.Description,
      photo: r.PhotoId, raisedBy: r.RaisedBy, status: r.Status, lastRemark: r.LastRemark, updated: r.Updated, solvedAt: r.SolvedAt,
      log: (byId[r.ID] || []).slice(-8) };
  });
  return { ok: true, issues: out.slice(-300) };
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
      type: c.Type, related: c.Related, remark: c.Remark, photo: c.PhotoId, by: c.By, flags: c.Flags };
  });
  return { ok: true, from: from, to: to, checks: rs };
}
