/* Minerva Daily Priorities - Google Sheet backup (version 1)
   1. Paste into the Google Sheet: Extensions > Apps Script (replace everything) and Save.
   2. Deploy > New deployment > gear > Web app. Execute as "Me", Who has access "Anyone". Deploy, allow access.
   3. Copy the Web app URL (ends with /exec) into the app: Data & settings > Google Sheet backup.
   Updating later: paste the new script, Save, then Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy.
   The full backup is a file in your Google Drive (next to this sheet), with a daily snapshot for 30 days.
   The tabs "Open tasks", "Completed" and "Updates log" are readable copies, rebuilt on every save. */
const KEY = '__KEY__';
const SCRIPT_VERSION = 1;
const KEEP_DAYS = 30;
const BASE = 'minerva-priorities';
const PRIO = { 1: 'Important and Urgent', 2: 'Important and Not Urgent', 3: 'Not Important but Urgent', 4: 'Not Important, Not Urgent' };
const STATUS = { ns: 'Not Started', on: 'Ongoing', wait: 'Waiting', done: 'Completed' };

function doPost(e) {
  let res;
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.key !== KEY) res = { ok: false, error: 'key' };
    else if (req.action === 'ping') res = Object.assign({ ok: true, script: SCRIPT_VERSION }, info_());
    else if (req.action === 'restore') res = restore_(req.day);
    else if (req.action === 'snapshots') res = { ok: true, list: snapList_() };
    else if (req.action === 'save') res = locked_(function () { return save_(req); });
    else res = { ok: false, error: 'action' };
  } catch (err) {
    res = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() { return ContentService.createTextOutput('Minerva Daily Priorities backup is running.'); }

// Optional: run once from the editor to grant access before deploying.
function setup() { folder_(); return 'Ready'; }

function locked_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(45000)) return { ok: false, error: 'busy' };
  try { return fn(); } finally { lock.releaseLock(); }
}
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function props_() { return PropertiesService.getScriptProperties(); }
function folder_() {
  const parents = DriveApp.getFileById(ss_().getId()).getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}
function file_(name, create) {
  const it = folder_().getFilesByName(name);
  if (it.hasNext()) return it.next();
  return create ? folder_().createFile(name, '', 'application/json') : null;
}
function info_() {
  const p = props_().getProperties();
  return { url: ss_().getUrl(), saved: Number(p.saved) || null, open: Number(p.open) || 0, total: Number(p.total) || 0, v: Number(p.v) || 0 };
}
function counts_(st) {
  const t = (st && st.tasks) || [];
  return { open: t.filter(function (x) { return x.status !== 'done'; }).length, total: t.length };
}

function save_(req) {
  const st = req.state || {};
  const c = counts_(st), prev = info_();
  if (req.expect != null && Number(req.expect) !== prev.v) return { ok: false, error: 'stale', v: prev.v };
  if (!req.force) {
    if (c.total === 0 && prev.total > 0) return { ok: false, error: 'empty', prev: prev };
    if (prev.total >= 10 && c.total < prev.total * 0.5) return { ok: false, error: 'shrink', prev: prev };
  }
  const now = new Date(), json = JSON.stringify(st), ver = prev.v + 1;
  file_(BASE + '-backup.json', true).setContent(json);
  const day = Utilities.formatDate(now, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const snap = file_(BASE + '-snapshot-' + day + '.json', true);
  snap.setContent(json);
  snap.setDescription(c.open + ' open, ' + c.total + ' total');
  props_().setProperties({ v: String(ver), saved: String(now.getTime()), open: String(c.open), total: String(c.total) });
  pruneSnapshots_();
  try { readable_(st); } catch (x) {}
  return { ok: true, saved: now.getTime(), open: c.open, total: c.total, v: ver, url: ss_().getUrl() };
}

function restore_(day) {
  const f = day ? file_(BASE + '-snapshot-' + day + '.json', false) : file_(BASE + '-backup.json', false);
  if (day && !f) return { ok: false, error: 'none' };
  const json = f ? f.getBlob().getDataAsString() : '';
  return { ok: true, state: json ? JSON.parse(json) : null, v: info_().v };
}

function snapList_() {
  const out = [], it = folder_().getFiles(), re = new RegExp('^' + BASE + '-snapshot-(\\d{4}-\\d{2}-\\d{2})\\.json$');
  while (it.hasNext()) {
    const f = it.next(), m = f.getName().match(re);
    if (m) out.push({ day: m[1], info: f.getDescription() || '' });
  }
  out.sort(function (a, b) { return a.day < b.day ? 1 : -1; });
  return out;
}
function pruneSnapshots_() {
  snapList_().slice(KEEP_DAYS).forEach(function (x) { const f = file_(BASE + '-snapshot-' + x.day + '.json', false); if (f) f.setTrashed(true); });
}

/* ---------- readable tabs (rebuilt on every save, so edits here are overwritten) ---------- */
function d_(k) { if (!k) return ''; const p = String(k).split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
function owners_(t) { return (t.owners && t.owners.length ? t.owners.slice().sort() : ['Unassigned']).join(' / '); }
function ups_(t) { return (t.updates || []).slice().sort(function (a, b) { return (b.date + (b.at || '')) < (a.date + (a.at || '')) ? -1 : 1; }); }
function due_(t, today) {
  if (!t.target) return 'No target date';
  if (t.target === today) return 'Due today';
  return t.target < today ? 'Overdue' : 'Due ' + t.target.slice(8, 10) + '/' + t.target.slice(5, 7);
}
function readable_(st) {
  const tasks = st.tasks || [], today = Utilities.formatDate(new Date(), ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const byOwner = function (a, b) { const x = owners_(a).toLowerCase(), y = owners_(b).toLowerCase(); return x < y ? -1 : x > y ? 1 : (a.prio - b.prio); };
  const open = tasks.filter(function (t) { return t.status !== 'done'; }).sort(byOwner);
  write_('Open tasks', ['Responsible', 'Priority', 'Task', 'Partner', 'End-user', 'Country', 'BDM', 'Status', 'Start date', 'Target date', 'Due', 'Latest update', 'All updates'],
    open.map(function (t) { const u = ups_(t); return [owners_(t), PRIO[t.prio] || '', t.title, t.partner || '', t.endUser || '', t.country || '', t.bdm || '', STATUS[t.status] || '', d_(t.start), d_(t.target), due_(t, today),
      u[0] ? u[0].date + ' ' + u[0].text : '', u.map(function (x) { return x.date + ' ' + x.text; }).join('\n')]; }),
    [9, 10], [150, 170, 260, 140, 130, 80, 90, 95, 95, 95, 110, 320, 420]);
  const done = tasks.filter(function (t) { return t.status === 'done'; }).sort(function (a, b) { return (b.doneAt || '') < (a.doneAt || '') ? -1 : 1; });
  write_('Completed', ['Completed on', 'How', 'Responsible', 'Task', 'Partner', 'End-user', 'Country', 'Latest update'],
    done.map(function (t) { const u = ups_(t); return [d_(t.doneAt), t.doneHow === 'removed' ? 'Removed' : 'Completed', owners_(t), t.title, t.partner || '', t.endUser || '', t.country || '', u[0] ? u[0].date + ' ' + u[0].text : '']; }),
    [1], [100, 90, 150, 260, 140, 130, 80, 420]);
  const log = [];
  tasks.forEach(function (t) { (t.updates || []).forEach(function (u) { log.push([d_(u.date), owners_(t), t.title, u.text, STATUS[t.status] || '']); }); });
  log.sort(function (a, b) { return b[0] - a[0]; });
  write_('Updates log', ['Date', 'Responsible', 'Task', 'Update', 'Task status'], log, [1], [100, 150, 260, 480, 100]);
  const s = ss_(), blank = s.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && s.getSheets().length > 1) s.deleteSheet(blank);
}
function write_(name, head, data, dateCols, widths) {
  const s = ss_().getSheetByName(name) || ss_().insertSheet(name);
  s.clear();
  const w = head.length, n = data.length;
  if (s.getMaxRows() < n + 1) s.insertRowsAfter(s.getMaxRows(), n + 1 - s.getMaxRows());
  if (s.getMaxColumns() < w) s.insertColumnsAfter(s.getMaxColumns(), w - s.getMaxColumns());
  s.getRange(1, 1, 1, w).setValues([head]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#004679');
  s.setFrozenRows(1);
  if (n) {
    s.getRange(2, 1, n, w).setValues(data).setVerticalAlignment('top').setWrap(true);
    dateCols.forEach(function (c) { s.getRange(2, c, n, 1).setNumberFormat('dd mmm yyyy'); });
  }
  widths.forEach(function (px, i) { if (i < w) s.setColumnWidth(i + 1, px); });
}
