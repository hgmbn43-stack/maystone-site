/**
 * Maystone – closer core (2nd calls). 24 Nov 2026.
 * The SAME file sits in both scripts (the sheet script and the portal script), so the Callbacks tab,
 * the morning email and the portal Closer page always agree.
 *
 * A "closer lead" is any lead whose Outcome starts with Yes (a preview was booked). A rep's own
 * "Call back" leads are the rep's job and are not listed here.
 *
 * Columns this adds to the end of a rep tab the first time it needs them (nothing is moved or renamed):
 *   CB agreed       Yes = the customer asked for this day and time
 *   CB tries        closer's tries in a row with no answer (goes back to 0 when someone speaks to them)
 *   CB last tried   when a closer last tried
 *   CB last result  No answer / Voicemail / Texted / Spoke – ...
 *   CB status       blank = open, Link sent, Not interested
 *   Last viewed     when the customer last opened the preview (written by the view tracker)
 *
 * Sections:
 *   notsent   yes, but the preview has not been sent yet
 *   overdue   they asked for a time and it has passed with nobody trying since
 *   today     they asked for a time today, or no time was agreed (call any time)
 *   chase     said yes and got the pay link 24h+ ago, still not paid: call again
 *   tried     no answer: next try is spaced out (1, 1, 2, 4, 6 days), so 6 tries cover about 2 weeks
 *   cold      no time agreed, preview sent over a week ago and never opened: lowest priority
 *   tomorrow / later   an agreed time on another day
 *   awaiting  said yes on the second call, pay link sent under 24h ago
 *   quiet     6 tries in a row with no answer (comes back if they open the preview again)
 *   done      paid, or not interested
 *
 * The call queue (o.due, in o.prio order): overdue, agreed time now, opened since we last tried,
 * unpaid yes, then any-time leads (opened recently first, then fewest tries), then old unopened ones.
 * Each lead carries o.why: one plain line saying why to call it now.
 * 25 Nov: 6 tries (was 3), spaced tries, chase, queue + why, "Not interested" ticks Do not call.
 */

var CZ_COLS = ['CB agreed', 'CB tries', 'CB last tried', 'CB last result', 'CB status', 'Last viewed'];
var CZ_QUIET_TRIES = 6;
var CZ_GAPS_DAYS = [1, 1, 2, 4, 6];   // days to wait after try 1, 2, 3, 4, 5+
var CZ_CHASE_HOURS = 24;
var CZ_HOT_HOURS = 48;
var CZ_TZ = 'Europe/London';
var CZ_PAY = { monthly: 'https://buy.stripe.com/8x2dR97zO4uwacB7EueQM04', oneoff: 'https://buy.stripe.com/28EdR98DSe56ckJaQGeQM01' };
var CZ_COLD_DAYS = 7;
var CZ_ORDER = ['overdue', 'today', 'chase', 'tried', 'cold', 'awaiting', 'tomorrow', 'later', 'notsent', 'quiet', 'done'];
var CZ_LABEL = { overdue: 'Overdue', today: 'Today', chase: 'Said yes, not paid', tried: 'Next try later', cold: 'Never opened (old)', awaiting: 'Waiting on payment', tomorrow: 'Tomorrow', later: 'Later', notsent: 'Preview not sent yet', quiet: 'Gone quiet', done: 'Done' };
var CZ_ACTS = { na: 'No answer', vm: 'Voicemail', text: 'Texted', cb: 'Spoke – call back', sold: 'Spoke – said yes, link sent', ni: 'Spoke – not interested', reopen: 'Reopened' };

var CZ_NOT_REP = ['Lead pool', 'Summary', 'Binned', 'Callbacks', 'AI receptionist list', 'Leads'];
function czRepTabs_(ss) {
  // The list of call-list tabs is remembered for 10 minutes, so we don't check every tab's A1 each time.
  var all = ss.getSheets(), c = CacheService.getScriptCache(), known = c.get('czTabs');
  if (known) { try { var names = JSON.parse(known), hit = all.filter(function (s) { return names.indexOf(s.getName()) !== -1; }); if (hit.length === names.length) return hit; } catch (e) {} }
  var out = all.filter(function (s) { var n = s.getName(); return CZ_NOT_REP.indexOf(n) === -1 && !/^backup/i.test(n) && String(s.getRange(1, 1).getValue()).trim() === 'Lead ID'; });
  try { c.put('czTabs', JSON.stringify(out.map(function (s) { return s.getName(); })), 600); } catch (e) {}
  return out;
}
function czDay_(ms) { return Utilities.formatDate(new Date(ms), CZ_TZ, 'yyyy-MM-dd'); }
/** London wall-clock time -> a moment, whatever timezone the script itself runs in. */
function czLondon_(ymd, hm) {
  var d = String(ymd).split('-'), t = String(hm || '00:00').split(':');
  var guess = Date.UTC(Number(d[0]), Number(d[1]) - 1, Number(d[2]), Number(t[0]) || 0, Number(t[1]) || 0);
  var off = Utilities.formatDate(new Date(guess), CZ_TZ, 'Z'), sign = off.charAt(0) === '-' ? -1 : 1;
  return guess - sign * (Number(off.substr(1, 2)) * 60 + Number(off.substr(3, 2))) * 60000;
}
/** A date cell (or yyyy-MM-dd / dd/MM/yyyy text) -> 'yyyy-MM-dd', or ''. */
function czYmd_(v, sstz) {
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, sstz, 'yyyy-MM-dd');
  var s = String(v || '').trim(), m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}
/** A time cell (a time value, or "17:30" / "5.30" text) -> 'HH:mm', or ''. */
function czHm_(v, sstz) {
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, sstz, 'HH:mm');
  var m = String(v || '').match(/(\d{1,2})[:.](\d{2})/); if (!m || Number(m[1]) > 23 || Number(m[2]) > 59) return '';
  return ('0' + m[1]).slice(-2) + ':' + m[2];
}
/** A date-time cell -> ms, or 0. Also reads the old "Yes – 24/11 11:13" first-view note. */
function czMs_(v, nowMs) {
  if (v instanceof Date) return isNaN(v.getTime()) ? 0 : v.getTime();
  var s = String(v || '').trim(); if (!s) return 0;
  var m = s.match(/(\d{1,2})\/(\d{1,2})(?:\/(\d{4}))?\s+(\d{1,2}):(\d{2})/);
  if (m) {
    var y = m[3] ? Number(m[3]) : Number(czDay_(nowMs).slice(0, 4));
    var ms = czLondon_(y + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2), m[4] + ':' + m[5]);
    if (!m[3] && ms > nowMs + 86400000) ms = czLondon_((y - 1) + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2), m[4] + ':' + m[5]);
    return ms;
  }
  var d = new Date(s); return isNaN(d.getTime()) ? 0 : d.getTime();
}
function czAddDays_(ymd, n) { return czDay_(czLondon_(ymd, '12:00') + n * 86400000); }

/** Every closer lead across every rep tab, with its section worked out. */
function czRows_(ss, nowMs) {
  nowMs = nowMs || Date.now();
  var sstz = ss.getSpreadsheetTimeZone() || CZ_TZ, today = czDay_(nowMs), tomorrow = czAddDays_(today, 1), out = [];
  czRepTabs_(ss).forEach(function (sh) {
    var lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2) return;
    var v = sh.getRange(1, 1, lr, lc).getValues(), head = v[0].map(function (h) { return String(h).trim(); });
    var col = function (h) { return head.indexOf(h); }, get = function (r, h) { var c = col(h); return c === -1 ? '' : r[c]; };
    for (var i = 1; i < v.length; i++) {
      var r = v[i], id = String(r[0]).trim(); if (!id) continue;
      var outcome = String(get(r, 'Outcome') || '').trim();
      if (!/^yes/i.test(outcome)) continue;
      if (get(r, 'Do not call') === true) continue;
      var o = {
        tab: sh.getName(), row: i + 1, id: id,
        business: String(get(r, 'Business') || ''), trade: String(get(r, 'Trade') || ''), town: String(get(r, 'Town') || ''),
        phone: String(get(r, 'Phone') || ''), email: String(get(r, 'Client email') || '').trim(), notes: String(get(r, 'Call notes') || ''),
        bookedBy: String(get(r, 'Booked by') || '') || sh.getName(), closer: String(get(r, 'Closer') || ''),
        preview: String(get(r, 'Preview link') || '').trim(), sent: !!String(get(r, 'Sent to client') || '').trim(),
        views: Number(get(r, 'Views')) || 0, pkg: String(get(r, 'Package') || ''),
        paid: /^paid/i.test(String(get(r, 'Paid') || '').trim()),
        agreed: /^y/i.test(String(get(r, 'CB agreed') || '')), tries: Number(get(r, 'CB tries')) || 0,
        lastTried: czMs_(get(r, 'CB last tried'), nowMs), lastResult: String(get(r, 'CB last result') || ''),
        status: String(get(r, 'CB status') || '').trim()
      };
      var sentRaw = get(r, 'Sent to client'), sm = String(sentRaw).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      o.sentAt = sentRaw instanceof Date ? sentRaw.getTime() : sm ? czLondon_(sm[3] + '-' + ('0' + sm[2]).slice(-2) + '-' + ('0' + sm[1]).slice(-2), '12:00') : 0;
      o.lastViewed = czMs_(get(r, 'Last viewed'), nowMs) || (o.views ? czMs_(get(r, 'Hot?'), nowMs) : 0);
      o.hot = !!o.lastViewed && nowMs - o.lastViewed <= CZ_HOT_HOURS * 3600000;
      var ymd = czYmd_(get(r, 'Next contact'), sstz), hm = czHm_(get(r, 'Callback time'), sstz);
      var at = ymd ? czLondon_(ymd, hm || '09:00') : 0;
      // An old time nobody confirmed and that has already gone is treated as "no time agreed".
      if (at && !o.agreed && ymd < today) at = 0;
      // A time that a closer has already tried after is used up.
      if (at && o.lastTried && o.lastTried >= at) at = 0;
      o.at = at; o.day = at ? ymd : ''; o.time = at && hm ? hm : '';
      var noAnswer = /^(no answer|voicemail|texted)/i.test(o.lastResult);
      o.triedToday = !!o.lastTried && czDay_(o.lastTried) === today;
      o.openedSinceTry = !!o.lastViewed && o.lastViewed > (o.lastTried || 0);
      var gap = CZ_GAPS_DAYS[Math.min(Math.max(o.tries, 1), CZ_GAPS_DAYS.length) - 1];
      o.nextTry = o.tries && o.lastTried && noAnswer ? czLondon_(czAddDays_(czDay_(o.lastTried), gap), '08:00') : 0;
      if (o.paid || /^not interested/i.test(o.status)) o.section = 'done';
      else if (!o.sent) o.section = 'notsent';
      else if (/^link sent/i.test(o.status)) o.section = o.lastTried && nowMs - o.lastTried >= CZ_CHASE_HOURS * 3600000 ? 'chase' : 'awaiting';
      else if (at && o.day > tomorrow) o.section = 'later';
      else if (at && o.day === tomorrow) o.section = 'tomorrow';
      else if (at && (o.day < today || (o.time && at < nowMs - 3600000))) o.section = 'overdue';
      else if (o.tries && o.hot && o.openedSinceTry) o.section = 'today';        // they opened it again after we tried: call now
      else if (o.tries >= CZ_QUIET_TRIES) o.section = 'quiet';
      else if (!at && o.nextTry && o.nextTry > nowMs) o.section = 'tried';
      else if (!at && !o.views && o.sentAt && nowMs - o.sentAt > CZ_COLD_DAYS * 86400000) o.section = 'cold';
      else o.section = 'today';
      czWhy_(o, nowMs);
      o.links = { preview: o.preview ? o.preview + (o.preview.indexOf('?') === -1 ? '?me' : '&me') : '', client: o.preview,
        monthly: CZ_PAY.monthly + '?client_reference_id=' + encodeURIComponent(id), oneoff: CZ_PAY.oneoff + '?client_reference_id=' + encodeURIComponent(id) };
      out.push(o);
    }
  });
  out.sort(function (a, b) {
    if (a.due !== b.due) return a.due ? -1 : 1;
    if (a.due) return a.prio - b.prio || (a.at || 0) - (b.at || 0) || (b.lastViewed || 0) - (a.lastViewed || 0) || a.tries - b.tries;
    var s = CZ_ORDER.indexOf(a.section) - CZ_ORDER.indexOf(b.section); if (s) return s;
    if (a.section === 'tomorrow' || a.section === 'later') return (a.at || 0) - (b.at || 0);   // other days: plain time order
    if (a.hot !== b.hot) return a.hot ? -1 : 1;
    if (!!a.time !== !!b.time) return a.time ? -1 : 1;                 // agreed times before "any time"
    if (a.at !== b.at) return (a.at || 0) - (b.at || 0);
    if (a.tries !== b.tries) return a.tries - b.tries;                  // fresh ones before ones already tried
    return (b.lastViewed || 0) - (a.lastViewed || 0);
  });
  return out;
}

function czCount_(rows) {
  var c = {}; CZ_ORDER.forEach(function (s) { c[s] = 0; }); c.hot = 0; c.due = 0; c.laterToday = 0; c.doneToday = 0;
  rows.forEach(function (o) {
    c[o.section]++; if (o.hot && o.section !== 'done') c.hot++;
    if (o.due) c.due++; else if (o.section === 'today' && o.time) c.laterToday++;
    if (o.triedToday) c.doneToday++;
  });
  return c;
}
/** Only the leads to call now, best first. */
function czQueue_(rows) { return rows.filter(function (o) { return o.due; }); }

/** Sets o.due, o.prio and o.why (one plain line: why call this one now). */
function czWhy_(o, nowMs) {
  var hm = function (ms) { return Utilities.formatDate(new Date(ms), CZ_TZ, 'HH:mm'); }, day = function (ms) { return Utilities.formatDate(new Date(ms), CZ_TZ, 'EEE d MMM'); };
  var ago = function (ms) { var m = Math.round((nowMs - ms) / 60000); return m < 2 ? 'just now' : m < 60 ? m + ' min ago' : czAgo_(ms, nowMs); };
  o.due = false; o.prio = 9;
  var s = o.section;
  if (s === 'overdue') { o.due = true; o.prio = 0; o.why = 'Asked for a call ' + (o.time ? 'at ' + o.time : 'on ' + day(o.at)) + (o.day < czDay_(nowMs) ? ' (' + day(o.at) + ')' : '') + ' – overdue'; return; }
  if (s === 'chase') { o.due = true; o.prio = 3; o.why = 'Said yes ' + czAgo_(o.lastTried, nowMs) + ', not paid yet'; return; }
  if (s === 'cold') { o.due = true; o.prio = 6; o.why = 'Preview sent ' + Math.round((nowMs - o.sentAt) / 86400000) + ' days ago, never opened'; return; }
  if (s === 'today') {
    if (o.time) { o.due = o.at <= nowMs + 15 * 60000; o.prio = 1; o.why = 'Asked for a call at ' + o.time; return; }
    o.due = true;
    if (o.tries && o.openedSinceTry) { o.prio = 2; o.why = 'Opened the preview again ' + ago(o.lastViewed) + ' – ring now'; return; }
    if (o.hot) { o.prio = o.views && nowMs - o.lastViewed < 3 * 3600000 ? 2 : 4; o.why = 'Opened the preview ' + ago(o.lastViewed) + (o.views > 1 ? ' (' + o.views + ' views)' : ''); return; }
    o.prio = 5;
    if (o.tries) { o.why = 'Try ' + (o.tries + 1) + ' of ' + CZ_QUIET_TRIES + ' · last: ' + String(o.lastResult).toLowerCase() + ' ' + czAgo_(o.lastTried, nowMs); return; }
    if (o.views) { o.why = 'Opened the preview ' + o.views + '×, last ' + czAgo_(o.lastViewed, nowMs); return; }
    o.why = 'Preview sent' + (o.sentAt ? ' ' + czAgo_(o.sentAt, nowMs).replace(/ \d\d:\d\d$/, '') : '') + ', not opened yet';
    return;
  }
  if (s === 'tried') { o.why = 'Tried ' + o.tries + '× · next try ' + (czDay_(o.nextTry) === czAddDays_(czDay_(nowMs), 1) ? 'tomorrow' : day(o.nextTry)); return; }
  if (s === 'awaiting') { o.why = 'Pay link sent ' + czAgo_(o.lastTried, nowMs); return; }
  if (s === 'tomorrow' || s === 'later') { o.why = 'Asked for a call ' + day(o.at) + (o.time ? ' at ' + o.time : ''); return; }
  if (s === 'quiet') { o.why = 'Tried ' + o.tries + ' times, no answer'; return; }
  if (s === 'notsent') { o.why = 'Preview not sent yet'; return; }
  o.why = o.paid ? 'Paid' : 'Not interested';
}

/** Short plain-English line for a lead: "Wed 25 Nov 17:30" / "any time". */
function czWhenText_(o) {
  if (!o.at) return o.section === 'tried' ? 'next try ' + Utilities.formatDate(new Date(o.nextTry), CZ_TZ, 'EEE d') : 'any time';
  return Utilities.formatDate(new Date(o.at), CZ_TZ, 'EEE d MMM') + (o.time ? ' ' + o.time : '');
}
function czAgo_(ms, nowMs) {
  if (!ms) return '';
  var day = czDay_(ms), today = czDay_(nowMs), hm = Utilities.formatDate(new Date(ms), CZ_TZ, 'HH:mm');
  if (day === today) return 'today ' + hm;
  if (day === czAddDays_(today, -1)) return 'yesterday ' + hm;
  return Utilities.formatDate(new Date(ms), CZ_TZ, 'EEE d MMM');
}

/**
 * Logs what happened on a closer's call, on the lead's own row.
 * act: na | vm | text | cb (needs o.date yyyy-MM-dd and o.time HH:mm) | sold | ni | reopen
 * Returns { ok, error? }.
 */
function czLog_(ss, tab, id, act, o, who) {
  o = o || {}; id = String(id || '').trim();
  if (!CZ_ACTS[act]) return { ok: false, error: 'Unknown result.' };
  var sh = ss.getSheetByName(String(tab)); if (!sh || !id) return { ok: false, error: 'That lead was not found. Refresh and try again.' };
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  if (head[0] !== 'Lead ID' || CZ_NOT_REP.indexOf(sh.getName()) !== -1) return { ok: false, error: 'That is not a call list.' };
  var ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues(), row = 0;
  for (var i = 1; i < ids.length; i++) if (String(ids[i][0]).trim() === id) { row = i + 1; break; }
  if (!row) return { ok: false, error: 'That lead has moved. Refresh and try again.' };
  if (act === 'cb') {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(o.date || '')) || !/^\d{2}:\d{2}$/.test(String(o.time || ''))) return { ok: false, error: 'Pick the day and time they asked for.' };
    if (czLondon_(o.date, o.time) < Date.now() - 600000) return { ok: false, error: 'That time has already gone.' };
  }
  var need = CZ_COLS.concat(['Closer', 'Next contact', 'Callback time', 'Call notes']);
  need.forEach(function (h) { if (head.indexOf(h) === -1) { sh.getRange(1, head.length + 1).setValue(h).setFontWeight('bold'); head.push(h); } });
  var cur = sh.getRange(row, 1, 1, head.length).getValues()[0];
  var set = function (h, val) { sh.getRange(row, head.indexOf(h) + 1).setValue(val); };
  var now = new Date(), tries = Number(cur[head.indexOf('CB tries')]) || 0;
  if (act === 'reopen') { set('CB status', ''); set('CB tries', 0); set('CB last result', CZ_ACTS.reopen); return { ok: true }; }
  var result = CZ_ACTS[act];
  if (act === 'na' && (o.vm || o.text)) result += ' (' + [o.vm ? 'left voicemail' : '', o.text ? 'texted' : ''].filter(String).join(', ') + ')';
  set('CB last tried', now); set('CB last result', result);
  if (act === 'na' || act === 'vm' || act === 'text') set('CB tries', tries + 1);
  else set('CB tries', 0);
  if (act === 'cb') {
    var d = o.date.split('-');
    set('Next contact', new Date(Number(d[0]), Number(d[1]) - 1, Number(d[2])));
    sh.getRange(row, head.indexOf('Callback time') + 1).setNumberFormat('@').setValue(o.time);
    set('CB agreed', 'Yes'); set('CB status', '');
  }
  if (act === 'sold') set('CB status', 'Link sent');
  if (act === 'ni') { set('CB status', 'Not interested'); var dnc = head.indexOf('Do not call'); if (dnc !== -1) sh.getRange(row, dnc + 1).setValue(true); }
  if (who && !String(cur[head.indexOf('Closer')] || '').trim()) set('Closer', who);
  var note = String(o.note || '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, 500);
  if (note) {
    var old = String(cur[head.indexOf('Call notes')] || '');
    set('Call notes', (old ? old + '\n' : '') + Utilities.formatDate(now, CZ_TZ, 'dd/MM') + ' ' + (who || 'Closer') + ': ' + note);
  }
  return { ok: true };
}
