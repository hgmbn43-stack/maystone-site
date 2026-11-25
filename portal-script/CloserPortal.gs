/**
 * Maystone – Closer page (portal script). 24 Nov 2026. Needs CloserCore.gs in the same project.
 * Admin and Managers only. Wired into doPost:
 *   if (a === 'closerList') { if (!isStaff_(me)) return json_({ ok: false, error: 'Closers only.' }); return json_(closerList_(me)); }
 *   if (a === 'closerLead') { if (!isStaff_(me)) return json_({ ok: false, error: 'Closers only.' }); return json_(closerLead_(me, req.email)); }
 *   if (a === 'closerLock') { if (!isStaff_(me)) return json_({ ok: false, error: 'Closers only.' }); return json_(closerLock_(me, req.tab, req.id)); }
 *   if (a === 'closerLog')  { if (!isStaff_(me)) return json_({ ok: false, error: 'Closers only.' }); lock.waitLock(20000); return json_(closerLog_(me, req.tab, req.id, req.act, req.opts || {})); }
 */

function closerName_(me) { return String(me.Name || me.Email || '').trim().split(/\s+/)[0]; }

/** A closer has a lead open: the other closer's queue skips it for 10 minutes. */
function closerLock_(me, tab, id) {
  try { CacheService.getScriptCache().put('lk:' + tab + '|' + id, closerName_(me), 600); } catch (e) {}
  return { ok: true };
}
/** Marks leads another closer has open (fresh every time, not part of the 30-second copy). */
function closerApplyLocks_(me, r) {
  if (!r || !r.ok || !r.rows) return r;
  var name = closerName_(me), keys = r.rows.map(function (o) { return 'lk:' + o.tab + '|' + o.id; }), got = {};
  try { got = CacheService.getScriptCache().getAll(keys); } catch (e) {}
  r.rows.forEach(function (o, i) { var who = got[keys[i]]; o.lockedBy = who && who !== name ? who : ''; });
  return r;
}
function closerList_(me) {
  var now = Date.now(), rows = czRows_(SpreadsheetApp.openById(CALL_LISTS_ID), now);
  rows.forEach(function (o) { o.when = czWhenText_(o); o.viewedAgo = czAgo_(o.lastViewed, now); o.triedAgo = czAgo_(o.lastTried, now); });
  return { ok: true, now: now, me: closerName_(me), counts: czCount_(rows), labels: CZ_LABEL, rows: rows.filter(function (o) { return o.section !== 'done'; }) };
}

/** The last few emails this customer sent us (read from the mailbox the portal runs as). */
function closerLead_(me, email) {
  email = cleanEmail_(email); if (!email) return { ok: true, emails: [] };
  try {
    var out = [];
    GmailApp.search('from:' + email, 0, 3).forEach(function (th) {
      var ms = th.getMessages(), m = null;
      for (var i = ms.length - 1; i >= 0; i--) if (ms[i].getFrom().toLowerCase().indexOf(email) !== -1) { m = ms[i]; break; }
      if (m) out.push({ when: Utilities.formatDate(m.getDate(), CZ_TZ, 'EEE d MMM HH:mm'), subject: str_(m.getSubject(), 120), text: str_(m.getPlainBody().split(/\n\s*On .{10,80} wrote:/)[0].replace(/\s+/g, ' '), 400) });
    });
    return { ok: true, emails: out };
  } catch (e) { return { ok: true, emails: [], note: 'Could not read emails.' }; }
}

function closerLog_(me, tab, id, act, opts) {
  var r = czLog_(SpreadsheetApp.openById(CALL_LISTS_ID), tab, id, act, opts, closerName_(me));
  if (!r.ok) return r;
  try { CacheService.getScriptCache().remove('lk:' + tab + '|' + id); } catch (e) {}
  log_(me.Email, 'closer call', String(id) + ' · ' + CZ_ACTS[act] + (act === 'cb' ? ' ' + opts.date + ' ' + opts.time : ''));
  return closerApplyLocks_(me, closerList_(me));
}
