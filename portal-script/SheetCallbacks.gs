/**
 * Maystone – Callbacks tab and morning call sheet (sheet script). 24 Nov 2026.
 * Replaces the old "Callbacks for the closers" block. Needs CloserCore.gs in the same project.
 *
 * - "Callbacks" tab: every preview that was booked, in sections (see CloserCore.gs), rebuilt every
 *   5 minutes from processAll. A rep's own "Call back" leads are no longer mixed in.
 * - Closer column: pick Harvey or Fin on the tab; it is saved back to the lead's own row.
 * - Weekdays 9am: one email with today's calls. Run installMorning() once to move it from 8am daily.
 */

var CB_TAB = 'Callbacks';
var CB_HEAD = ['Section', 'Why now', 'When', 'Hot', 'Business', 'Trade', 'Town', 'Phone', 'Viewed', 'Tried', 'Last result', 'Closer', 'Booked by', 'Notes', 'Preview', 'Monthly link', 'One-off link', 'Rep tab', 'Lead ID'];
var CBI = {}; CB_HEAD.forEach(function (h, i) { CBI[h] = i; });
var CB_BG = { overdue: '#FDECEC', today: '#FFF4E0', chase: '#FDECEC', tried: '#F3F4F6', cold: '#F3F4F6', awaiting: '#E6F1FB', tomorrow: '#FFFFFF', later: '#FFFFFF', notsent: '#F7F2FA', quiet: '#EEEEEE', done: '#E8F5E9' };
var CLOSERS = ['Harvey', 'Fin'];
var MORNING_TO = ['harvey@maystone.uk', 'finsj101@gmail.com'];

function buildCallbacks() {
  var ss = SpreadsheetApp.getActive(), now = Date.now(), rows = czRows_(ss, now);
  var out = rows.map(function (o) {
    return [o.due ? 'Call now' : CZ_LABEL[o.section], o.why, czWhenText_(o), o.hot ? 'HOT' : '', o.business, o.trade, o.town, o.phone,
      o.views ? o.views + 'x' + (o.lastViewed ? ', last ' + czAgo_(o.lastViewed, now) : '') : 'Not yet',
      o.tries ? o.tries + 'x' + (o.lastTried ? ', last ' + czAgo_(o.lastTried, now) : '') : (o.lastTried ? czAgo_(o.lastTried, now) : ''),
      o.lastResult, o.closer, o.bookedBy, o.notes, o.links.preview, o.links.monthly, o.links.oneoff, o.tab, o.id];
  });
  var cb = ss.getSheetByName(CB_TAB) || ss.insertSheet(CB_TAB);
  cb.clear(); try { cb.getRange(1, 1, cb.getMaxRows(), cb.getMaxColumns()).clearDataValidations(); } catch (e) {}
  cb.getRange(1, 1, 1, CB_HEAD.length).setValues([CB_HEAD]).setFontWeight('bold').setBackground('#142D4C').setFontColor('#ffffff');
  cb.setFrozenRows(1);
  if (out.length) {
    cb.getRange(2, 1, out.length, CB_HEAD.length).setValues(out);
    cb.getRange(2, CBI.Closer + 1, out.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(CLOSERS, true).setAllowInvalid(true).build());
    cb.getRange(2, 1, out.length, CB_HEAD.length).setBackgrounds(rows.map(function (o) { return CB_HEAD.map(function () { return CB_BG[o.section]; }); }));
    cb.getRange(2, 1, out.length, 1).setFontWeight('bold');
    cb.getRange(2, CBI.Hot + 1, out.length, 1).setFontWeight('bold').setFontColor('#D94A12');
    cb.getRange(2, CBI.Notes + 1, out.length, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  } else {
    cb.getRange(2, 1).setValue('No previews booked yet.');
  }
  var c = czCount_(rows);
  cb.getRange(1, CB_HEAD.length + 2).setValue('Rebuilt ' + Utilities.formatDate(new Date(), CZ_TZ, 'dd/MM HH:mm') +
    '. Call now ' + c.due + ' (in this order) · Hot ' + c.hot + '. "Call now" rows are in the order to ring them; "Why now" says why. ' +
    'No answer: next try is spaced out (1, 1, 2, 4, 6 days); after 6 tries it goes to Gone quiet. Log calls on the Closer page in the portal; pick your name in Closer here.');
  return rows;
}

/** Called from onEditRun: a Closer picked on the Callbacks tab is saved back to the lead's row. */
function callbackEdit_(e) {
  var sh = e.range.getSheet(); if (sh.getName() !== CB_TAB) return false;
  if (e.range.getColumn() !== CBI.Closer + 1 || e.range.getNumRows() !== 1 || e.range.getRow() < 2) return true;
  var row = sh.getRange(e.range.getRow(), 1, 1, CB_HEAD.length).getValues()[0], tab = String(row[CBI['Rep tab']]), id = String(row[CBI['Lead ID']]).trim();
  var rep = SpreadsheetApp.getActive().getSheetByName(tab); if (!rep || !id) return true;
  var head = rep.getRange(1, 1, 1, rep.getLastColumn()).getValues()[0].map(String), c = head.indexOf('Closer');
  if (c === -1) { c = head.length; rep.getRange(1, c + 1).setValue('Closer').setFontWeight('bold'); }
  var ids = rep.getRange(1, 1, rep.getLastRow(), 1).getValues();
  for (var i = 1; i < ids.length; i++) if (String(ids[i][0]).trim() === id) { rep.getRange(i + 1, c + 1).setValue(e.range.getValue()); break; }
  return true;
}

/** Called from recordView_ after a real view is counted: stamps "Last viewed" on the lead's row. */
function cbStampView_(sh, r, id) {
  var now = new Date(), head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); }), c = head.indexOf('Last viewed');
  if (c === -1) { c = head.length; sh.getRange(1, c + 1).setValue('Last viewed').setFontWeight('bold'); }
  sh.getRange(r, c + 1).setValue(now).setNumberFormat('dd/mm/yyyy hh:mm');
}

function cbEsc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

/** The morning call sheet: the call queue in order, with why to ring each one. Returns the HTML. */
function morningHtml_(rows, now) {
  var queue = czQueue_(rows), later = rows.filter(function (o) { return o.section === 'today' && !o.due && o.time; });
  var c = czCount_(rows), td = 'padding:10px 12px;border-bottom:1px solid #E4DED3;vertical-align:top';
  var line = function (o, i) {
    return '<tr><td style="' + td + ';white-space:nowrap;color:#667385"><b>' + (i + 1) + '</b></td>' +
      '<td style="' + td + '"><b>' + cbEsc_(o.business) + '</b> <small style="color:#667385">' + cbEsc_([o.trade, o.town].filter(String).join(' · ')) + '</small><br>' +
      '<b style="color:' + (o.section === 'overdue' || o.section === 'chase' ? '#B42318' : o.hot ? '#D94A12' : '#142D4C') + '">' + cbEsc_(o.why) + '</b><br>' +
      '<a href="tel:' + String(o.phone).replace(/[^\d+]/g, '') + '" style="font-size:16px;font-weight:bold;color:#142D4C">' + cbEsc_(o.phone) + '</a> <small style="color:#667385">· booked by ' + cbEsc_(o.bookedBy) + '</small>' +
      (o.notes ? '<br><small style="color:#4A5A70">' + cbEsc_(o.notes.slice(0, 200)) + '</small>' : '') + '</td></tr>';
  };
  var top = queue.slice(0, 15);
  return '<div style="font:15px Arial,sans-serif;color:#142D4C;max-width:640px">' +
    '<h2 style="margin:0 0 4px">' + queue.length + ' second calls to make</h2><p style="margin:0 0 12px;color:#4A5A70">In order: ring from the top. ' + c.hot + ' opened their preview in the last 2 days.</p>' +
    '<p style="margin:0 0 14px"><a href="https://maystone.uk/login/#closer" style="display:inline-block;background:#D94A12;color:#fff;padding:12px 18px;border-radius:10px;font-weight:bold;text-decoration:none">Open the call list</a></p>' +
    (top.length ? '<table style="border-collapse:collapse;width:100%;font:14px Arial,sans-serif">' + top.map(line).join('') + '</table>' : '<p>Nothing to call right now.</p>') +
    (queue.length > top.length ? '<p><small>…and ' + (queue.length - top.length) + ' more on the call list.</small></p>' : '') +
    (later.length ? '<p style="margin-top:16px"><b>Later today at a set time:</b> ' + later.map(function (o) { return cbEsc_(o.time + ' ' + o.business); }).join(' · ') + '</p>' : '') +
    '<p style="margin-top:16px;color:#4A5A70"><small>' + c.tomorrow + ' booked for tomorrow · ' + c.later + ' later · ' + c.awaiting + ' waiting on payment · ' + c.tried + ' waiting for their next try · ' + c.quiet + ' gone quiet.</small></p></div>';
}

function morningCallbacks() {
  var now = Date.now(), dow = Utilities.formatDate(new Date(now), CZ_TZ, 'u');   // 1 = Monday … 7 = Sunday
  if (dow === '6' || dow === '7') return;
  var rows = buildCallbacks(), c = czCount_(rows);
  if (!c.due && !c.laterToday) return;
  GmailApp.sendEmail(MORNING_TO.join(','), 'Second calls today: ' + c.due + (c.hot ? ' (' + c.hot + ' hot)' : ''), 'Open in an email app that shows HTML.',
    { htmlBody: morningHtml_(rows, now), from: SENDER, name: SENDER_NAME });
}

/** Run once from the editor: the 9am email, weekdays. */
function installMorning() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'morningCallbacks') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('morningCallbacks').timeBased().everyDays(1).atHour(9).nearMinute(0).inTimezone('Europe/London').create();
  buildCallbacks();
  return 'Morning email set for 9am on weekdays. Callbacks tab built.';
}
