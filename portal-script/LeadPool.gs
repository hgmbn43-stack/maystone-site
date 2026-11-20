/* ---------------- lead pool: reps pull fresh leads when they want to call ----------------
 * "Get my leads" on the Leads page: the rep's untouched leads go back into the "Lead pool" tab,
 * then they get up to POOL_MAX fresh ones (best first). Anything they've worked stays with them.
 * Every night returnAllUntouched() sweeps untouched leads off rep tabs back into the pool,
 * so nobody sits on a pile. Harvey's and Fin's tabs are left alone by the nightly sweep.
 */
var POOL_TAB = 'Lead pool';
var POOL_MAX = 50;
var POOL_DEFAULT = 30;
var POOL_SKIP_NIGHTLY = ['Harvey', 'Fin'];
// A lead counts as worked if any of these has something in it.
var WORKED_COLS = ['Called (tick)', 'Outcome', 'Client email', 'Call notes', 'Preview link', 'Sent to client', 'Next contact',
                   'Paid', 'Booked by', 'Callback time', 'Interested in', 'Details', 'Closer', 'Package'];

function headOf_(sh) {
  var lc = sh.getLastColumn(); if (lc < 1) return [];
  var h = sh.getRange(1, 1, 1, lc).getValues()[0].map(function (x) { return String(x).trim(); });
  while (h.length && !h[h.length - 1]) h.pop();
  return h;
}
function lastIdRowOf_(sh) {
  var lr = sh.getLastRow(); if (lr < 2) return 1;
  var ids = sh.getRange(1, 1, lr, 1).getValues();
  for (var i = ids.length - 1; i >= 1; i--) if (String(ids[i][0]).trim()) return i + 1;
  return 1;
}
function isUntouched_(row, head) {
  for (var i = 0; i < WORKED_COLS.length; i++) {
    var c = head.indexOf(WORKED_COLS[i]); if (c === -1) continue;
    var v = row[c];
    if (v instanceof Date) return false;
    if (v === true) return false;
    if (String(v == null ? '' : v).trim() && v !== false) return false;
  }
  var d = head.indexOf('Do not call'); if (d !== -1 && row[d] === true) return false;
  var t = head.indexOf('Tries'); if (t !== -1 && Number(row[t]) > 0) return false;
  return true;
}
function ensurePool_(ss, template) {
  var pool = ss.getSheetByName(POOL_TAB);
  if (pool) return pool;
  pool = ss.insertSheet(POOL_TAB, ss.getSheets().length);
  var h = headOf_(template);
  pool.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
  pool.setFrozenRows(1);
  return pool;
}
/** Copies rows (as objects keyed by heading) onto the end of a sheet, matching columns by heading. */
function writeObjs_(sh, objs) {
  if (!objs.length) return;
  var h = headOf_(sh);
  var vals = objs.map(function (o) { return h.map(function (k) { return o[k] === undefined ? '' : o[k]; }); });
  var start = lastIdRowOf_(sh) + 1;
  if (sh.getMaxRows() < start + vals.length - 1) sh.insertRowsAfter(sh.getMaxRows(), start + vals.length - 1 - sh.getMaxRows());
  sh.getRange(start, 1, vals.length, h.length).setValues(vals);
}
/** Deletes the given 1-based row numbers, bottom first, in contiguous blocks. */
function deleteRowsList_(sh, rows) {
  rows = rows.slice().sort(function (a, b) { return b - a; });
  var i = 0;
  while (i < rows.length) {
    var end = rows[i], start = end;
    while (i + 1 < rows.length && rows[i + 1] === start - 1) { i++; start = rows[i]; }
    sh.deleteRows(start, end - start + 1);
    i++;
  }
}
/** Moves untouched leads from a rep tab into the pool. Returns how many moved. */
function returnUntouchedFrom_(sh, pool) {
  var lr = lastIdRowOf_(sh); if (lr < 2) return 0;
  var h = headOf_(sh), v = sh.getRange(2, 1, lr - 1, h.length).getValues(), objs = [], rows = [];
  v.forEach(function (r, i) {
    if (!String(r[0]).trim() || !isUntouched_(r, h)) return;
    var o = {}; h.forEach(function (k, c) { if (k) o[k] = r[c]; }); objs.push(o); rows.push(i + 2);
  });
  if (!objs.length) return 0;
  writeObjs_(pool, objs);
  deleteRowsList_(sh, rows);
  return objs.length;
}
function poolRank_(o) {
  var p = { A: 0, B: 1, C: 2 }[String(o['Priority'] || '').trim()];
  return [p === undefined ? 3 : p, -(Number(o['Google reviews']) || 0)];
}
/** Deals the best n leads from the pool onto a rep tab. Returns how many were dealt. */
function dealFromPool_(pool, sh, n) {
  var lr = lastIdRowOf_(pool); if (lr < 2 || n < 1) return 0;
  var h = headOf_(pool), v = pool.getRange(2, 1, lr - 1, h.length).getValues(), items = [];
  v.forEach(function (r, i) {
    if (!String(r[0]).trim()) return;
    var o = {}; h.forEach(function (k, c) { if (k) o[k] = r[c]; }); items.push({ o: o, row: i + 2, rank: poolRank_(o) });
  });
  items.sort(function (a, b) { return a.rank[0] - b.rank[0] || a.rank[1] - b.rank[1]; });
  var pick = items.slice(0, n);
  if (!pick.length) return 0;
  writeObjs_(sh, pick.map(function (x) { return x.o; }));
  deleteRowsList_(pool, pick.map(function (x) { return x.row; }));
  return pick.length;
}
function poolSize_(pool) { return Math.max(0, lastIdRowOf_(pool) - 1); }

/** Portal action: rep presses "Get my leads". */
function getLeads_(me, count) {
  var n = Math.max(1, Math.min(POOL_MAX, Math.floor(Number(count) || POOL_DEFAULT)));
  var tab = String(me['Call list tab'] || '').trim();
  if (!tab) return { ok: false, error: 'You do not have a call list yet. Ask Harvey to set one up.' };
  var ss = SpreadsheetApp.openById(CALL_LISTS_ID), sh = ss.getSheetByName(tab);
  if (!sh) return { ok: false, error: 'Your call list tab "' + tab + '" was not found.' };
  var pool = ensurePool_(ss, sh);
  var back = returnUntouchedFrom_(sh, pool);
  var got = dealFromPool_(pool, sh, n);
  SpreadsheetApp.flush();
  log_(me.Email, 'got leads', got + ' new, ' + back + ' untouched sent back to the pool');
  var msg = got ? 'You\'ve got ' + got + ' fresh leads.' : 'The lead pool is empty right now. New leads are added every weekday morning.';
  if (back) msg += ' ' + back + ' you hadn\'t called went back to the pool.';
  return { ok: true, got: got, returned: back, pool: poolSize_(pool), message: msg };
}

/** Nightly: every rep's untouched leads go back into the pool. Run by a time trigger. */
function returnAllUntouched() {
  var lock = LockService.getScriptLock(); if (!lock.tryLock(60000)) return;
  try {
    var ss = SpreadsheetApp.openById(CALL_LISTS_ID), total = 0, pool = null;
    table_('Users').forEach(function (u) {
      var tab = String(u['Call list tab'] || '').trim();
      if (!tab || tab === POOL_TAB || POOL_SKIP_NIGHTLY.indexOf(tab) !== -1) return;
      var sh = ss.getSheetByName(tab); if (!sh) return;
      if (!pool) pool = ensurePool_(ss, sh);
      try { total += returnUntouchedFrom_(sh, pool); } catch (e) { log_('system', 'pool sweep failed', tab + ': ' + e.message); }
    });
    log_('system', 'nightly pool sweep', total + ' untouched leads back in the pool');
  } finally { lock.releaseLock(); }
}

/** Run once from the editor: switches on the nightly sweep at about 2am. */
function switchOnLeadPool() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'returnAllUntouched') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('returnAllUntouched').timeBased().atHour(2).everyDays(1).inTimezone('Europe/London').create();
}
