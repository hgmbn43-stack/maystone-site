/* ---------------- lead pool: reps pull fresh leads when they want to call ----------------
 * "Get my leads" on the Leads page: the rep's untouched leads go back into the "Lead pool" tab,
 * then they get up to POOL_MAX fresh ones (best first). Anything they've worked stays with them.
 * Every night returnAllUntouched() sends back leads that have sat untouched for POOL_KEEP_DAYS days,
 * and leads that got POOL_NA_TRIES no-answers with nothing else logged (so someone else tries them).
 * Harvey's and Fin's tabs are left alone by the nightly sweep.
 */
var POOL_TAB = 'Lead pool';
var POOL_MAX = 50;
var POOL_DEFAULT = 30;
var POOL_KEEP_DAYS = 3;          // untouched leads stay with a rep this long before the nightly sweep takes them
var POOL_NA_TRIES = 3;           // no-answers before a lead goes back for someone else to try
var POOL_SKIP_NIGHTLY = ['Harvey', 'Fin'];
var POOL_SKIP_RESET = ['Harvey', 'Fin', 'Gabriel'];
var DEALT_COL = 'Dealt', HISTORY_COL = 'History';
// A lead counts as worked if any of these has something in it.
var WORKED_COLS = ['Called (tick)', 'Outcome', 'Client email', 'Call notes', 'Preview link', 'Sent to client', 'Next contact',
                   'Paid', 'Booked by', 'Callback time', 'Interested in', 'Details', 'Closer', 'Package'];
// A no-answer lead only goes back if none of these has anything in it.
var NA_BLOCK_COLS = ['Outcome', 'Client email', 'Preview link', 'Sent to client', 'Next contact', 'Paid', 'Booked by', 'Callback time',
                     'Interested in', 'Details', 'Closer', 'Package'];

function headOf_(sh) {
  var lc = sh.getLastColumn(); if (lc < 1) return [];
  var h = sh.getRange(1, 1, 1, lc).getValues()[0].map(function (x) { return String(x).trim(); });
  while (h.length && !h[h.length - 1]) h.pop();
  return h;
}
function ensureCol_(sh, name) {
  var h = headOf_(sh), i = h.indexOf(name); if (i !== -1) return i;
  sh.getRange(1, h.length + 1).setValue(name).setFontWeight('bold');
  return h.length;
}
function lastIdRowOf_(sh) {
  var lr = sh.getLastRow(); if (lr < 2) return 1;
  var ids = sh.getRange(1, 1, lr, 1).getValues();
  for (var i = ids.length - 1; i >= 1; i--) if (String(ids[i][0]).trim()) return i + 1;
  return 1;
}
function filled_(v) {
  if (v instanceof Date) return !isNaN(v.getTime());
  if (v === true) return true;
  if (v === false || v == null) return false;
  return !!String(v).trim();
}
function isUntouched_(row, head) {
  for (var i = 0; i < WORKED_COLS.length; i++) { var c = head.indexOf(WORKED_COLS[i]); if (c !== -1 && filled_(row[c])) return false; }
  var d = head.indexOf('Do not call'); if (d !== -1 && row[d] === true) return false;
  var t = head.indexOf('Tries'); if (t !== -1 && Number(row[t]) > 0) return false;
  return true;
}
/** A leaver's lead with no outcome yet (only ticks/tries/notes): goes back to the pool with a History note. */
function isOpenLead_(row, head) {
  var d = head.indexOf('Do not call'); if (d !== -1 && row[d] === true) return false;
  for (var i = 0; i < NA_BLOCK_COLS.length; i++) { var k = head.indexOf(NA_BLOCK_COLS[i]); if (k !== -1 && filled_(row[k])) return false; }
  return true;
}
/** Tried POOL_NA_TRIES times with no answer and nothing else logged. */
function isNoAnswerDone_(row, head) {
  var c = head.indexOf('Called (tick)'), t = head.indexOf('Tries'), d = head.indexOf('Do not call');
  if (c === -1 || t === -1) return false;
  if (!/^\s*(no answer|voicemail|busy)/i.test(String(row[c] || ''))) return false;
  if ((Number(row[t]) || 0) < POOL_NA_TRIES) return false;
  if (d !== -1 && row[d] === true) return false;
  for (var i = 0; i < NA_BLOCK_COLS.length; i++) { var k = head.indexOf(NA_BLOCK_COLS[i]); if (k !== -1 && filled_(row[k])) return false; }
  return true;
}
function ensurePool_(ss, template) {
  var pool = ss.getSheetByName(POOL_TAB);
  if (!pool) {
    pool = ss.insertSheet(POOL_TAB, ss.getSheets().length);
    var h = headOf_(template);
    pool.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    pool.setFrozenRows(1);
  }
  ['Tries', HISTORY_COL].forEach(function (n) { ensureCol_(pool, n); });
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
function dmy_(d) { return Utilities.formatDate(d, 'Europe/London', 'dd/MM'); }
/**
 * Moves leads from a rep tab into the pool.
 * mode 'all': every untouched lead plus no-answer leads that are done (Get my leads, the one-off reset).
 * mode 'nightly': untouched leads dealt POOL_KEEP_DAYS+ days ago, plus no-answer leads that are done.
 *   Untouched leads with no Dealt date get today's date, so the clock starts now.
 * Returns { untouched, noAnswer }.
 */
function returnFrom_(sh, pool, mode, tab) {
  var res = { untouched: 0, noAnswer: 0 };
  var lr = lastIdRowOf_(sh); if (lr < 2) return res;
  var dc = mode === 'nightly' ? ensureCol_(sh, DEALT_COL) : headOf_(sh).indexOf(DEALT_COL);
  var h = headOf_(sh), v = sh.getRange(2, 1, lr - 1, h.length).getValues(), objs = [], rows = [], now = new Date();
  var stamp = false, dv = dc !== -1 ? v.map(function (r) { return [r[dc]]; }) : null;
  v.forEach(function (r, i) {
    if (!String(r[0]).trim()) return;
    var o = null;
    if (isUntouched_(r, h)) {
      if (mode === 'nightly') {
        var dealt = r[dc];
        if (!(dealt instanceof Date) || isNaN(dealt.getTime())) { dv[i][0] = now; stamp = true; return; }
        if (now - dealt < POOL_KEEP_DAYS * 86400000) return;
      }
      o = {}; h.forEach(function (k, c) { if (k) o[k] = r[c]; });
      res.untouched++;
    } else if (isNoAnswerDone_(r, h) || (mode === 'leaver' && isOpenLead_(r, h))) {
      o = {}; h.forEach(function (k, c) { if (k) o[k] = r[c]; });
      var notes = String(o['Call notes'] || '').replace(/\s+/g, ' ').trim().slice(0, 200);
      var line = dmy_(now) + ': ' + (Number(o['Tries']) || 0) + ' no answers (' + tab + ')' + (notes ? ' – ' + notes : '');
      o[HISTORY_COL] = (String(o[HISTORY_COL] || '').trim() ? String(o[HISTORY_COL]).trim() + ' | ' : '') + line;
      o['Called (tick)'] = ''; o['Tries'] = ''; o['Call notes'] = '';
      res.noAnswer++;
    }
    if (!o) return;
    o[DEALT_COL] = '';
    objs.push(o); rows.push(i + 2);
  });
  if (stamp) sh.getRange(2, dc + 1, dv.length, 1).setValues(dv);
  if (!objs.length) return res;
  writeObjs_(pool, objs);
  deleteRowsList_(sh, rows);
  return res;
}
function poolRank_(o) {
  var p = { A: 0, B: 1, C: 2 }[String(o['Priority'] || '').trim()];
  return [p === undefined ? 3 : p, String(o[HISTORY_COL] || '').trim() ? 1 : 0, -(Number(o['Google reviews']) || 0)];
}
/** Deals the best n leads from the pool onto a rep tab (never one this rep already had no answer on). Returns how many. */
function dealFromPool_(pool, sh, n, tab) {
  var lr = lastIdRowOf_(pool); if (lr < 2 || n < 1) return 0;
  var h = headOf_(pool), v = pool.getRange(2, 1, lr - 1, h.length).getValues(), items = [];
  v.forEach(function (r, i) {
    if (!String(r[0]).trim()) return;
    var o = {}; h.forEach(function (k, c) { if (k) o[k] = r[c]; });
    if (tab && String(o[HISTORY_COL] || '').indexOf('(' + tab + ')') !== -1) return;
    items.push({ o: o, row: i + 2, rank: poolRank_(o) });
  });
  items.sort(function (a, b) { return a.rank[0] - b.rank[0] || a.rank[1] - b.rank[1] || a.rank[2] - b.rank[2]; });
  var pick = items.slice(0, n);
  if (!pick.length) return 0;
  [DEALT_COL, HISTORY_COL].forEach(function (c) { ensureCol_(sh, c); });
  var now = new Date();
  writeObjs_(sh, pick.map(function (x) { x.o[DEALT_COL] = now; return x.o; }));
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
  var r = returnFrom_(sh, pool, 'all', tab), back = r.untouched + r.noAnswer;
  var got = dealFromPool_(pool, sh, n, tab);
  SpreadsheetApp.flush();
  log_(me.Email, 'got leads', got + ' new, ' + back + ' sent back to the pool');
  var msg = got ? 'You\'ve got ' + got + ' fresh leads.' : 'The lead pool is empty right now. New leads are added every weekday morning.';
  if (r.untouched) msg += ' ' + r.untouched + ' you hadn\'t called went back to the pool.';
  if (r.noAnswer) msg += ' ' + r.noAnswer + ' that didn\'t answer ' + POOL_NA_TRIES + ' times went back for someone else to try.';
  return { ok: true, got: got, returned: back, pool: poolSize_(pool), message: msg };
}

function sweepTabs_(skip, mode, label) {
  var lock = LockService.getScriptLock(); if (!lock.tryLock(60000)) return 'busy';
  try {
    var ss = SpreadsheetApp.openById(CALL_LISTS_ID), pool = null, out = [], tot = 0, seen = {};
    table_('Users').forEach(function (u) {
      var tab = String(u['Call list tab'] || '').trim();
      if (!tab || seen[tab] || tab === POOL_TAB || skip.indexOf(tab) !== -1) return; seen[tab] = 1;
      var sh = ss.getSheetByName(tab); if (!sh) return;
      if (!pool) pool = ensurePool_(ss, sh);
      try { var r = returnFrom_(sh, pool, u.Status === 'Active' ? mode : 'leaver', tab); tot += r.untouched + r.noAnswer; if (r.untouched + r.noAnswer) out.push(tab + ' ' + r.untouched + '+' + r.noAnswer); }
      catch (e) { log_('system', label + ' failed', tab + ': ' + e.message); }
    });
    SpreadsheetApp.flush();
    var summary = tot + ' leads back in the pool' + (out.length ? ' (untouched+no answer: ' + out.join(', ') + ')' : '');
    log_('system', label, summary);
    return summary;
  } finally { lock.releaseLock(); }
}
/** Nightly (time trigger). */
function returnAllUntouched() { return sweepTabs_(POOL_SKIP_NIGHTLY, 'nightly', 'nightly pool sweep'); }
/** One-off: every rep's untouched leads back to the pool now (not Harvey, Fin or Gabriel). Run from the editor. */
function resetLeadsNow() { var s = sweepTabs_(POOL_SKIP_RESET, 'all', 'lead reset'); Logger.log(s); return s; }

/** Run once from the editor: switches on the nightly sweep at about 2am and the evening report at about 8:30pm. */
function switchOnLeadPool() {
  ScriptApp.getProjectTriggers().forEach(function (t) { var f = t.getHandlerFunction(); if (f === 'returnAllUntouched' || f === 'eveningReport') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('returnAllUntouched').timeBased().atHour(2).everyDays(1).inTimezone('Europe/London').create();
  ScriptApp.newTrigger('eveningReport').timeBased().atHour(20).nearMinute(30).everyDays(1).inTimezone('Europe/London').create();
}

/* ---------------- evening report: each rep's day, emailed to Harvey ---------------- */
var REPORT_TO = ['harvey@maystone.uk'];

function dayStart_(d) { return new Date(Utilities.formatDate(d, 'Europe/London', "yyyy-MM-dd'T'00:00:00XXX")).getTime(); }
/** Breaks one Log row into a stat bucket, or ''. */
function statKind_(act, det) {
  var val = String(det).split(' · ').slice(1).join(' · ');
  if (act === 'call logged') return /^\s*(no answer|voicemail|busy)/i.test(val) ? 'na' : /not in service/i.test(val) ? 'dead' : 'ans';
  if (act === 'outcome') { if (/call\s*back/i.test(val)) return 'cb'; if (/not interested|already has|not trading/i.test(val)) return 'no'; return ''; }
  if (act === 'preview booked') return 'yes';
  return '';
}
function eveningReport() {
  var tz = 'Europe/London', now = new Date(), from = dayStart_(now);
  var users = table_('Users').filter(function (u) { return u.Status === 'Active' && String(u['Call list tab'] || '').trim(); });
  var byEmail = {}; users.forEach(function (u) { byEmail[cleanEmail_(u.Email)] = { u: u, calls: 0, na: 0, ans: 0, dead: 0, no: 0, cb: 0, yes: 0, ids: {}, wkCalls: 0 }; });
  var sh = sheet_('Log'), weekFrom = from - 6 * 86400000;
  if (sh && sh.getLastRow() > 1) {
    var n = Math.min(8000, sh.getLastRow() - 1), rows = sh.getRange(sh.getLastRow() - n + 1, 1, n, 4).getValues();
    rows.forEach(function (r) {
      var t = new Date(r[0]).getTime(); if (!t || t < weekFrom) return;
      var s = byEmail[cleanEmail_(r[1])]; if (!s) return;
      var act = String(r[2]), det = String(r[3]), k = statKind_(act, det);
      if (act === 'call logged') { s.wkCalls++; if (t >= from) s.calls++; }
      if (t < from) return;
      if (k) s[k]++;
      if (act === 'updated lead' && /Call notes/.test(det)) { var pre = det.split(': ')[0].split(' '); s.ids[pre[pre.length - 1]] = 1; }
    });
  }
  var cl = SpreadsheetApp.openById(CALL_LISTS_ID), blocks = [], team = { calls: 0, yes: 0, cb: 0 };
  users.forEach(function (u) {
    var e = cleanEmail_(u.Email), s = byEmail[e], tab = String(u['Call list tab']).trim(), notes = [];
    var tsh = cl.getSheetByName(tab);
    if (tsh && Object.keys(s.ids).length) {
      var h = headOf_(tsh), lr = lastIdRowOf_(tsh);
      if (lr >= 2) tsh.getRange(2, 1, lr - 1, h.length).getValues().forEach(function (r) {
        var id = String(r[0]).trim(); if (!s.ids[id]) return;
        var nt = String(r[h.indexOf('Call notes')] || '').replace(/\s+/g, ' ').trim(); if (!nt) return;
        notes.push(String(r[h.indexOf('Business')] || id) + ' (' + (String(r[h.indexOf('Outcome')] || r[h.indexOf('Called (tick)')] || '').trim() || 'no outcome') + '): ' + nt.slice(0, 220));
      });
    }
    team.calls += s.calls; team.yes += s.yes; team.cb += s.cb;
    var picked = s.ans, tip;
    if (!s.calls && !s.yes) tip = s.wkCalls ? 'Nothing logged today. Check in: did they call and not log it?' : 'No calls logged this week. Worth a friendly check-in.';
    else if (s.calls >= 15 && s.na / s.calls > 0.7) tip = 'Most calls went unanswered. Try a different time of day (late morning or 4–6pm).';
    else if (picked >= 8 && !s.yes) tip = picked + ' people picked up but no previews. Listen to a call together and work on the first 10 seconds and the preview offer.';
    else if (s.no > s.yes * 3 && s.no >= 4) tip = 'Lots of "not interested". Practise the busy-owner and "already have someone" answers.';
    else if (s.yes) tip = 'Booked ' + s.yes + ' today. Ask what worked and share it with the group.';
    else tip = 'Keep the volume up and log every call, even no answers.';
    blocks.push({ name: String(u.Name || tab), html:
      '<h3 style="margin:18px 0 4px;font:600 16px Arial">' + esc_(String(u.Name || tab)) + '</h3>' +
      '<p style="margin:0 0 6px;font:14px Arial">' + s.calls + ' calls · ' + s.na + ' no answer · ' + s.ans + ' picked up · ' + s.no + ' said no · ' + s.cb + ' callbacks · <b>' + s.yes + ' yes</b>' + (s.dead ? ' · ' + s.dead + ' dead numbers' : '') + ' <span style="color:#777">(' + s.wkCalls + ' calls in 7 days)</span></p>' +
      (notes.length ? '<ul style="margin:0 0 6px;padding-left:18px;font:13px Arial;color:#333">' + notes.slice(0, 8).map(function (x) { return '<li>' + esc_(x) + '</li>'; }).join('') + '</ul>' : '') +
      '<p style="margin:0;font:14px Arial"><b>Coaching point:</b> ' + esc_(tip) + '</p>', calls: s.calls });
  });
  blocks.sort(function (a, b) { return b.calls - a.calls; });
  var html = '<div style="max-width:640px"><p style="font:15px Arial">Team today: <b>' + team.calls + ' calls, ' + team.yes + ' previews booked, ' + team.cb + ' callbacks</b>.</p>' +
    blocks.map(function (b) { return b.html; }).join('') +
    '<p style="margin-top:20px;font:12px Arial;color:#777">Only calls logged on the portal are counted. Notes are from leads updated today.</p></div>';
  var subject = 'Maystone day report – ' + Utilities.formatDate(now, tz, 'EEE d MMM') + ' (' + team.calls + ' calls, ' + team.yes + ' yes)';
  REPORT_TO.forEach(function (to) {
    try { GmailApp.sendEmail(to, subject, 'Open in an email app that shows HTML.', { htmlBody: html, from: FROM_ALIAS, name: FROM_NAME }); }
    catch (e) { GmailApp.sendEmail(to, subject, 'Open in an email app that shows HTML.', { htmlBody: html, name: FROM_NAME }); }
  });
  log_('system', 'evening report', subject);
}


/** Run from the editor: every lead on a leaver's tab (portal Status not Active) that is untouched or has no outcome yet goes back to the pool now. Leads with a callback/yes stay for Harvey or Fin. */
function returnLeaversNow() {
  var ss = SpreadsheetApp.openById(CALL_LISTS_ID), pool = null, out = [];
  table_('Users').forEach(function (u) {
    if (u.Status === 'Active') return;
    var tab = String(u['Call list tab'] || '').trim(); if (!tab || tab === POOL_TAB) return;
    var sh = ss.getSheetByName(tab); if (!sh) return;
    if (!pool) pool = ensurePool_(ss, sh);
    var r = returnFrom_(sh, pool, 'leaver', tab); out.push(tab + ' ' + r.untouched + '+' + r.noAnswer);
  });
  SpreadsheetApp.flush(); Logger.log(out.join(', ') || 'nothing to return'); return out.join(', ');
}
