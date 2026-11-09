/**
 * Maystone Team Portal – backend.
 * A Google Apps Script web app, bound to the "Maystone – Team Portal" sheet in harvey@maystone.uk's Drive.
 *
 * - Reps log in with a one-time link sent to their email (no passwords).
 * - The admin (ADMIN_EMAIL) sees everyone: roles, partners, previews, sales, earnings; creates logins.
 * - Previews booked and paid sales are pulled from the Call Lists sheet every hour (syncActivity).
 *
 * Deploy: Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
 * The site (maystone.uk/login) calls this with POST {action, ...} as text/plain JSON.
 */

var ADMIN_EMAIL = 'harvey@maystone.uk';
var SITE_LOGIN = 'https://maystone.uk/login/';
var FROM_ALIAS = 'hello@maystone.uk';
var FROM_NAME = 'Maystone';
var CALL_LISTS_ID = '1QWrKee1Li45r1Zw2YcIV4XMSUyC-oFueAx4wWglWoEQ';
var CL = { id: 1, biz: 2, outcome: 10, link: 13, pkg: 18, paid: 19 };   // Call Lists columns (A, B, J, M, R, S)
var LINK_MINUTES = 30;
var SESSION_DAYS = 30;
var CLEAR_DAYS = 14;                                                   // refund window before a sale clears

var SHEETS = {
  Users: ['Email', 'Name', 'Role', 'Partner', 'Status', 'Call list tab', 'Commission per sale', 'Phone', 'Joined', 'Last login', 'Notes'],
  Partners: ['Name', 'What reps sell', 'Commission per sale', 'Status', 'Started', 'Contact', 'Notes'],
  Sales: ['ID', 'Date', 'Rep email', 'Partner', 'Customer / lead', 'Sale amount', 'Commission', 'Status', 'Source', 'Notes'],
  Activity: ['Rep email', 'Previews booked', 'Sales', 'Updated'],
  Training: ['Role', 'Title', 'Link', 'Type', 'Order'],
  Links: ['Token hash', 'Email', 'Expires', 'Used'],
  Sessions: ['Session hash', 'Email', 'Expires', 'Created'],
  Applications: ['Date', 'First name', 'Last name', 'Email', 'Mobile', 'Based', 'Experience', 'Best time', 'Note', 'Status'],
  Enquiries: ['Date', 'First name', 'Last name', 'Business', 'Email', 'Phone', 'Website', 'What they sell', 'Who buys it', 'Sale price', 'Commission', 'Reps wanted', 'Note', 'Status'],
  Progress: ['Email', 'Title', 'Done'],
  Passwords: ['Email', 'Hash', 'Salt', 'Set'],
  Events: ['ID', 'Title', 'When', 'Minutes', 'Link', 'For', 'Details', 'Created'],
  Callbacks: ['ID', 'Email', 'When', 'Name', 'Business', 'Phone', 'Note', 'Status', 'Created'],
  Log: ['Time', 'Who', 'Action', 'Details']
};
var ROLES = ['Rep', 'Top Rep', 'Coach', 'Team Lead', 'Manager', 'Admin'];
var MANAGER_CAN_SET = ['Rep', 'Top Rep', 'Coach', 'Team Lead'];
var ADMIN_ONLY = ['adminSavePartner', 'adminAddSale', 'adminSetSaleStatus'];
var LEVEL = { 'All': 0, 'Rep': 1, 'Top Rep': 2, 'Coach': 3, 'Team Lead': 4, 'Manager': 6, 'Admin': 9 };
var USER_STATUS = ['Active', 'Paused', 'Left'];
var SALE_STATUS = ['Pending', 'Cleared', 'Paid out', 'Refunded'];

/* ---------------- web entry points ---------------- */

function doGet() {
  return json_({ ok: true, service: 'Maystone team portal' });
}

function doPost(e) {
  var req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, error: 'Bad request.' }); }
  var lock = LockService.getScriptLock();
  try {
    var a = String(req.action || '');
    if (a === 'requestLink') return json_(requestLink_(req.email));
    if (a === 'verify') return json_(verify_(req.token));
    if (a === 'loginPassword') return json_(loginPassword_(req.email, req.password));
    if (a === 'apply' || a === 'enquiry') { lock.waitLock(20000); return json_(publicForm_(a, req.form || {})); }
    var me = auth_(req.session);
    if (!me) return json_({ ok: false, error: 'Please log in again.', auth: false });
    if (a === 'me') return json_(mePayload_(me));
    if (a === 'logout') { endSession_(req.session); return json_({ ok: true }); }
    if (a === 'setPassword') { lock.waitLock(20000); return json_(setPassword_(me, req.password)); }
    if (a === 'module') return json_(module_(me, req.title));
    if (a === 'slide') return json_(slide_(me, req.title, req.n));
    if (a === 'leads') return json_(leads_(me, req.tab));
    if (a === 'adminPreviews') { if (!isStaff_(me)) return json_({ ok: false, error: 'Admins only.' }); return json_(adminPreviews_()); }
    if (a === 'saveLead') { lock.waitLock(20000); return json_(saveLead_(me, req.tab, req.row, req.id, req.fields || {})); }
    if (a === 'callbacks') return json_(callbacks_(me));
    if (a === 'saveCallback') { lock.waitLock(20000); return json_(saveCallback_(me, req.cb || {})); }
    if (a === 'deleteCallback') { lock.waitLock(20000); return json_(deleteCallback_(me, req.id)); }
    if (a === 'markDone') { lock.waitLock(20000); return json_(markDone_(me, req.title, !!req.done)); }
    if (a.indexOf('academy') === 0) return json_(academy_(me, a, req, lock));   // Academy.gs
    if (!isStaff_(me)) return json_({ ok: false, error: 'Only the admin can do that.' });
    if (!isAdmin_(me) && ADMIN_ONLY.indexOf(a) !== -1) return json_({ ok: false, error: 'Only Harvey can change partners and sales.' });
    lock.waitLock(20000);
    if (a === 'adminOverview') return json_(adminOverview_());
    if (a === 'adminSaveUser') return json_(adminSaveUser_(me, req.user || {}, !!req.invite));
    if (a === 'adminSendLink') return json_(adminSendLink_(me, req.email));
    if (a === 'adminSavePartner') return json_(adminSavePartner_(me, req.partner || {}));
    if (a === 'adminAddSale') return json_(adminAddSale_(me, req.sale || {}));
    if (a === 'adminSetSaleStatus') return json_(adminSetSaleStatus_(me, req.id, req.status));
    if (a === 'adminSaveEvent') { lock.waitLock(20000); return json_(adminSaveEvent_(me, req.event || {})); }
    if (a === 'adminDeleteEvent') { lock.waitLock(20000); return json_(adminDeleteEvent_(me, req.id)); }
    if (a === 'adminSaveTraining') return json_(adminSaveTraining_(me, req.item || {}));
    if (a === 'adminSync') { syncActivity(); return json_(adminOverview_()); }
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'Something went wrong. Try again in a minute.' });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/* ---------------- passwords (optional, after a first email-link login) ---------------- */

function startSession_(u, how) {
  var now = new Date(), session = randomToken_();
  appendRow_('Sessions', [sha_(session), u.Email, new Date(Date.now() + SESSION_DAYS * 86400000), now]);
  setUserField_(u.Email, 'Last login', now);
  log_(u.Email, how, '');
  return { ok: true, session: session, user: publicUser_(u) };
}
function pwHash_(password, salt) { var h = salt + '|' + password; for (var i = 0; i < 300; i++) h = sha_(salt + h); return h; }
function pwRow_(email) {
  if (!sheet_('Passwords')) return null;
  var rows = sheet_('Passwords').getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if (cleanEmail_(rows[i][0]) === email) return { row: i + 1, hash: rows[i][1], salt: rows[i][2] };
  return null;
}
function hasPassword_(email) { return !!pwRow_(email); }
function setPassword_(me, password) {
  password = String(password || '');
  if (password.length < 8 || password.length > 100) return { ok: false, error: 'Use at least 8 characters.' };
  ensureSheet_('Passwords'); var sh = sheet_('Passwords');
  try { sh.hideSheet(); } catch (e) {}
  var salt = randomToken_().slice(0, 32), hash = pwHash_(password, salt), r = pwRow_(me.Email);
  if (r) sh.getRange(r.row, 2, 1, 3).setValues([[hash, salt, new Date()]]); else sh.appendRow([me.Email, hash, salt, new Date()]);
  log_(me.Email, 'password set', '');
  return { ok: true };
}
function loginPassword_(email, password) {
  email = cleanEmail_(email); var bad = { ok: false, error: 'That email and password don\'t match. Try again, or use "Email me a login link".' };
  if (!email || !password) return { ok: false, error: 'Enter your email and password.' };
  var cache = CacheService.getScriptCache(), key = 'pw:' + email, n = Number(cache.get(key) || 0);
  if (n >= 6) return { ok: false, error: 'Too many tries. Wait 15 minutes, or use "Email me a login link".' };
  var u = findUser_(email), r = pwRow_(email);
  if (!u || !r || pwHash_(String(password), r.salt) !== r.hash) { cache.put(key, String(n + 1), 900); log_(email, 'password login refused', ''); return bad; }
  if (u.Status !== 'Active') return { ok: false, error: 'This login is switched off. Contact Harvey.' };
  cache.remove(key);
  return startSession_(u, 'logged in (password)');
}

/* ---------------- Academy: training as members-only pages ---------------- */
// Docs are read live and turned into page blocks. Decks are rendered to slide images once by refreshAcademy()
// (run it from the editor after changing a deck) and kept in the private "Maystone Academy (site)" Drive folder.

var ACADEMY_FOLDER = 'Maystone Academy (site)';
function academyFolder_() { var it = DriveApp.getFoldersByName(ACADEMY_FOLDER); return it.hasNext() ? it.next() : DriveApp.createFolder(ACADEMY_FOLDER); }
function fileIdFrom_(link) { var m = String(link || '').match(/\/d\/([\w-]{20,})/); return m ? m[1] : ''; }
function trainingFor_(me, title) {
  var lvl = LEVEL[me.Role] || 1, t = table_('Training').filter(function (x) { return x.Title === title; })[0];
  if (!t) return null;
  if ((LEVEL[t.Role] || 0) > lvl) return { locked: true };
  return t;
}
function module_(me, title) {
  var t = trainingFor_(me, String(title || ''));
  if (!t) return { ok: false, error: 'Module not found.' };
  if (t.locked) return { ok: false, error: 'This module unlocks at a higher level.' };
  var id = fileIdFrom_(t.Link);
  if (!id) return { ok: true, kind: 'link', link: t.Link };
  var f = DriveApp.getFileById(id), mt = f.getMimeType();
  if (mt === MimeType.GOOGLE_DOCS) return { ok: true, kind: 'doc', blocks: docBlocks_(id) };
  var meta = deckMeta_(id);
  if (meta) return { ok: true, kind: 'deck', count: meta.slides.length, notes: meta.slides.map(function (x) { return x.notes; }), texts: meta.slides.map(function (x) { return x.text; }) };
  return { ok: true, kind: 'link', link: t.Link };
}
function deckMeta_(id) {
  var it = academyFolder_().getFilesByName(id + '.json');
  return it.hasNext() ? JSON.parse(it.next().getBlob().getDataAsString()) : null;
}
function slide_(me, title, n) {
  var t = trainingFor_(me, String(title || ''));
  if (!t || t.locked) return { ok: false, error: 'Module not found.' };
  var meta = deckMeta_(fileIdFrom_(t.Link)); n = Number(n) || 0;
  if (!meta || !meta.slides[n]) return { ok: false, error: 'Slide not found.' };
  var cache = CacheService.getScriptCache(), key = 'sl:' + meta.slides[n].img, hit = cache.get(key);
  if (hit) return { ok: true, img: hit };
  var b = DriveApp.getFileById(meta.slides[n].img).getBlob(), data = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
  if (data.length < 95000) cache.put(key, data, 21600);
  return { ok: true, img: data };
}

/** Run from the editor whenever a deck changes: renders every deck in the Training tab to slide images. */
function pngFetch_(presId, pageId, token) {
  for (var k = 0; k < 7; k++) {
    var r = UrlFetchApp.fetch('https://docs.google.com/presentation/d/' + presId + '/export/png?pageid=' + pageId, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (r.getResponseCode() === 200) { Utilities.sleep(1200); return r.getBlob(); }
    Utilities.sleep(2000 * Math.pow(2, k));
  }
  throw new Error('Slide export kept failing - run refreshAcademy again in a few minutes.');
}
/** Rebuild every deck, even ones already built (run after editing a deck). */
function rebuildAcademy() { return refreshAcademy(true); }
function refreshAcademy(force) {
  force = force === true;
  var folder = academyFolder_(), done = [];
  table_('Training').forEach(function (t) {
    var id = fileIdFrom_(t.Link); if (!id) return;
    var f; try { f = DriveApp.getFileById(id); } catch (e) { done.push('missing: ' + t.Title); return; }
    var mt = f.getMimeType(); if (mt === MimeType.GOOGLE_DOCS || mt.indexOf('presentation') === -1) return;
    var presId = id, temp = null;
    if (mt !== MimeType.GOOGLE_SLIDES) { temp = Drive.Files.copy({ name: '(temp) ' + f.getName(), mimeType: MimeType.GOOGLE_SLIDES }, id, { supportsAllDrives: true }); presId = temp.id; }
    if (!force && folder.getFilesByName(id + '.json').hasNext()) { done.push(t.Title + ': already built'); return; }
    var old = folder.getFilesByName(id + '.json'); while (old.hasNext()) { var o = old.next(); try { JSON.parse(o.getBlob().getDataAsString()).slides.forEach(function (x) { DriveApp.getFileById(x.img).setTrashed(true); }); } catch (e) {} o.setTrashed(true); }
    var pres = SlidesApp.openById(presId), token = ScriptApp.getOAuthToken(), slides = [];
    pres.getSlides().forEach(function (sl, i) {
      var png = pngFetch_(presId, sl.getObjectId(), token).setName(id + '-' + (i + 1) + '.png');
      var img = folder.createFile(png);
      var notes = ''; try { notes = sl.getNotesPage().getSpeakerNotesShape().getText().asString().trim(); } catch (e) {}
      var text = sl.getShapes().map(function (sh) { try { return sh.getText().asString().trim(); } catch (e) { return ''; } }).filter(String).join('\n').slice(0, 2000);
      slides.push({ img: img.getId(), notes: notes, text: text });
    });
    folder.createFile(id + '.json', JSON.stringify({ title: t.Title, built: new Date(), slides: slides }), 'application/json');
    if (temp) DriveApp.getFileById(temp.id).setTrashed(true);
    done.push(t.Title + ': ' + slides.length + ' slides');
  });
  console.log(done.join('\n'));
  return done.join('\n');
}

function docBlocks_(id) {
  var body = DocumentApp.openById(id).getBody(), out = [], E = DocumentApp.ElementType, H = DocumentApp.ParagraphHeading;
  for (var i = 0; i < body.getNumChildren(); i++) {
    var el = body.getChild(i), type = el.getType();
    if (type === E.PARAGRAPH) {
      var p = el.asParagraph(); if (!p.getText().trim()) continue;
      var h = p.getHeading(), tag = (h === H.TITLE) ? 'h1' : (h === H.HEADING1 || h === H.SUBTITLE) ? 'h2' : (h === H.HEADING2) ? 'h3' : (h === H.HEADING3 || h === H.HEADING4 || h === H.HEADING5 || h === H.HEADING6) ? 'h4' : 'p';
      out.push({ t: tag, h: rich_(p) });
    } else if (type === E.LIST_ITEM) {
      var li = el.asListItem(); if (!li.getText().trim()) continue;
      var g = li.getGlyphType(), ordered = g === DocumentApp.GlyphType.NUMBER || g === DocumentApp.GlyphType.LATIN_LOWER || g === DocumentApp.GlyphType.LATIN_UPPER || g === DocumentApp.GlyphType.ROMAN_LOWER || g === DocumentApp.GlyphType.ROMAN_UPPER;
      out.push({ t: 'li', h: rich_(li), n: li.getNestingLevel(), o: ordered, id: li.getListId() });
    } else if (type === E.TABLE) {
      var tb = el.asTable(), rows = [];
      for (var r = 0; r < tb.getNumRows(); r++) { var row = tb.getRow(r), cells = []; for (var c = 0; c < row.getNumCells(); c++) cells.push(esc_(row.getCell(c).getText())); rows.push(cells); }
      out.push({ t: 'table', rows: rows });
    } else if (type === E.HORIZONTAL_RULE) out.push({ t: 'hr' });
  }
  return out;
}
function esc_(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function rich_(el) {
  var t = el.editAsText(), s = t.getText(), idx = t.getTextAttributeIndices(), html = '';
  for (var k = 0; k < idx.length; k++) {
    var a = idx[k], b = k + 1 < idx.length ? idx[k + 1] : s.length; if (b <= a) continue;
    var seg = esc_(s.substring(a, b)), url = t.getLinkUrl(a);
    if (t.isBold(a)) seg = '<b>' + seg + '</b>';
    if (t.isItalic(a)) seg = '<i>' + seg + '</i>';
    if (url && /^https?:\/\//.test(url)) seg = '<a href="' + esc_(url) + '" target="_blank" rel="noopener">' + seg + '</a>';
    html += seg;
  }
  return html;
}

/* ---------------- task progress ---------------- */

function progressFor_(email) {
  var m = {}; if (!sheet_('Progress')) return m;
  table_('Progress').forEach(function (r) { if (cleanEmail_(r.Email) === email) m[r.Title] = r.Done; });
  return m;
}

function markDone_(me, title, done) {
  title = str_(title, 200); if (!title) return { ok: false, error: 'Missing task.' };
  ensureSheet_('Progress');
  var sh = sheet_('Progress'), rows = sh.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) if (cleanEmail_(rows[i][0]) === me.Email && rows[i][1] === title) sh.deleteRow(i + 1);
  if (done) sh.appendRow([me.Email, title, new Date()]);
  return { ok: true };
}

/* ---------------- public forms: Apply to join, Find reps ---------------- */

function ensureSheet_(name) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(SHEETS[name]); sh.setFrozenRows(1); sh.getRange(1, 1, 1, SHEETS[name].length).setFontWeight('bold'); }
  return sh;
}

function publicForm_(kind, f) {
  ensureSheet_('Applications'); ensureSheet_('Enquiries');
  if (str_(f.website)) return { ok: true };                            // honeypot filled: quietly ignore bots
  var cache = CacheService.getScriptCache(), key = 'form_' + sha_(kind + cleanEmail_(f.email));
  var n = Number(cache.get(key) || 0); if (n >= 3) return { ok: false, error: 'You have already sent this. We will be in touch soon.' };
  var email = cleanEmail_(f.email), phone = str_(f.phone, 30);
  if (!str_(f.first) || !email || phone.replace(/\D/g, '').length < 10 || !f.consent) return { ok: false, error: 'Please fill in all the required fields.' };
  var lines;
  if (kind === 'apply') {
    var row = { 'Date': new Date(), 'First name': str_(f.first, 60), 'Last name': str_(f.last, 60), 'Email': email, 'Mobile': phone, 'Based': str_(f.based, 80),
      'Experience': str_(f.experience, 30), 'Best time': str_(f.time, 30), 'Note': str_(f.note, 1000), 'Status': 'New' };
    appendObj_('Applications', row);
    lines = ['New rep application', '', 'Name: ' + row['First name'] + ' ' + row['Last name'], 'Mobile: ' + phone, 'Email: ' + email, 'Based: ' + row.Based,
      'Experience: ' + row.Experience, 'Best time to call: ' + row['Best time'], 'Note: ' + (row.Note || '-')];
  } else {
    var e = { 'Date': new Date(), 'First name': str_(f.first, 60), 'Last name': str_(f.last, 60), 'Business': str_(f.business, 120), 'Email': email, 'Phone': phone,
      'Website': str_(f.web, 200), 'What they sell': str_(f.sells, 300), 'Who buys it': str_(f.buyers, 300), 'Sale price': str_(f.price, 30),
      'Commission': str_(f.commission, 30), 'Reps wanted': str_(f.reps, 30), 'Note': str_(f.note, 1000), 'Status': 'New' };
    if (!e.Business) return { ok: false, error: 'Please fill in all the required fields.' };
    appendObj_('Enquiries', e);
    lines = ['New business enquiry', '', 'Business: ' + e.Business, 'Contact: ' + e['First name'] + ' ' + e['Last name'], 'Phone: ' + phone, 'Email: ' + email,
      'Website: ' + (e.Website || '-'), 'Sells: ' + e['What they sell'], 'Buyers: ' + e['Who buys it'], 'Sale price: ' + e['Sale price'],
      'Commission: ' + (e.Commission || '-'), 'Reps wanted: ' + e['Reps wanted'], 'Note: ' + (e.Note || '-')];
  }
  cache.put(key, String(n + 1), 3600);
  try { MailApp.sendEmail({ to: ADMIN_EMAIL, subject: (kind === 'apply' ? 'New application: ' : 'New enquiry: ') + str_(f.first, 60) + ' ' + str_(f.last, 60), body: lines.join('\n') + '\n\nOpen the Team Portal sheet to update the status.', name: FROM_NAME }); } catch (err) { log_('system', 'form email failed', err.message); }
  log_(email, kind, 'received');
  return { ok: true };
}

/* ---------------- login: one-time email links ---------------- */

function requestLink_(email) {
  email = cleanEmail_(email);
  var generic = { ok: true, message: 'If that email is on the team, a login link is on its way. It works for ' + LINK_MINUTES + ' minutes.' };
  if (!email) return { ok: false, error: 'Enter your email address.' };
  var cache = CacheService.getScriptCache(), key = 'rl:' + email, n = Number(cache.get(key) || 0);
  if (n >= 5) return generic;                                   // max 5 links an hour per email, quietly
  cache.put(key, String(n + 1), 3600);
  var u = findUser_(email);
  if (!u || u.Status !== 'Active') { log_(email, 'login link refused', u ? 'status ' + u.Status : 'not on the team'); return generic; }
  sendLoginLink_(u, false);
  return generic;
}

function sendLoginLink_(u, welcome) {
  var token = randomToken_();
  appendRow_('Links', [sha_(token), u.Email, new Date(Date.now() + LINK_MINUTES * 60000), '']);
  var link = SITE_LOGIN + '#t=' + token;
  var first = String(u.Name || '').split(' ')[0] || 'there';
  var subject = welcome ? 'Welcome to Maystone – your login' : 'Your Maystone login link';
  var intro = welcome
    ? 'Hi ' + first + ',\n\nWelcome to Maystone Sales Group. Your rep login is ready. Tap the link below to log in and see your training, your numbers and the leaderboard.'
    : 'Hi ' + first + ',\n\nHere is your Maystone login link.';
  var body = intro + '\n\n' + link + '\n\nThe link works once, for ' + LINK_MINUTES + ' minutes. Next time, go to ' + SITE_LOGIN + ' and enter your email to get a new one.\n\nIf you didn\'t ask for this, you can ignore this email.\n\nMaystone Sales Group';
  var opts = { name: FROM_NAME };
  try { if (GmailApp.getAliases().indexOf(FROM_ALIAS) !== -1) opts.from = FROM_ALIAS; } catch (e) {}
  GmailApp.sendEmail(u.Email, subject, body, opts);
  log_(u.Email, welcome ? 'welcome link sent' : 'login link sent', '');
}

function verify_(token) {
  token = String(token || '');
  if (!/^[a-f0-9]{48,}$/.test(token)) return { ok: false, error: 'That login link is not valid. Ask for a new one.' };
  var sh = sheet_('Links'), rows = sh.getDataRange().getValues(), h = sha_(token), now = new Date();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] !== h) continue;
    if (rows[i][3]) return { ok: false, error: 'That login link has already been used. Ask for a new one.' };
    if (new Date(rows[i][2]) < now) return { ok: false, error: 'That login link has expired. Ask for a new one.' };
    sh.getRange(i + 1, 4).setValue(now);
    var u = findUser_(rows[i][1]);
    if (!u || u.Status !== 'Active') return { ok: false, error: 'This login is switched off. Contact Harvey.' };
    return startSession_(u, 'logged in (link)');
  }
  return { ok: false, error: 'That login link is not valid. Ask for a new one.' };
}

function auth_(session) {
  session = String(session || '');
  if (!/^[a-f0-9]{48,}$/.test(session)) return null;
  var h = sha_(session), cache = CacheService.getScriptCache(), email = cache.get('s:' + h);
  if (!email) {
    var rows = sheet_('Sessions').getDataRange().getValues(), now = new Date();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === h && new Date(rows[i][2]) > now) { email = rows[i][1]; break; }
    }
    if (!email) return null;
    cache.put('s:' + h, email, 600);
  }
  var u = findUser_(email);
  return u && u.Status === 'Active' ? u : null;
}

function endSession_(session) {
  var h = sha_(String(session || '')), sh = sheet_('Sessions'), rows = sh.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) if (rows[i][0] === h) sh.getRange(i + 1, 3).setValue(new Date(0));
  CacheService.getScriptCache().remove('s:' + h);
}

function isAdmin_(u) { return u && u.Role === 'Admin' && cleanEmail_(u.Email) === ADMIN_EMAIL; }
/** Manager: runs the reps (team, logins, training, group calls, previews list, every rep's leads) but not money, partners or the admin's own list. */
function isManager_(u) { return !!u && u.Role === 'Manager' && u.Status === 'Active'; }
function isStaff_(u) { return isAdmin_(u) || isManager_(u); }

/* ---------------- what a rep sees ---------------- */

function mePayload_(u) {
  var act = activityMap_(), sales = table_('Sales'), mine = { previews: 0, sales: 0, pending: 0, cleared: 0, paidOut: 0 };
  var a = act[u.Email] || {};
  mine.previews = Number(a['Previews booked'] || 0);
  sales.forEach(function (s) {
    if (cleanEmail_(s['Rep email']) !== u.Email || s.Status === 'Refunded') return;
    mine.sales++;
    var c = Number(s.Commission || 0);
    if (s.Status === 'Paid out') mine.paidOut += c; else if (s.Status === 'Cleared') mine.cleared += c; else mine.pending += c;
  });
  var board = leaderboard_(act, sales), rank = 0;
  board.forEach(function (b, i) { if (b.email === u.Email) rank = i + 1; });
  var lvl = LEVEL[u.Role] || 1, done = progressFor_(u.Email);
  var training = table_('Training')
    .sort(function (x, y) { return (LEVEL[x.Role] || 0) - (LEVEL[y.Role] || 0) || Number(x.Order || 0) - Number(y.Order || 0); })
    .map(function (t) { var locked = (LEVEL[t.Role] || 0) > lvl; return { role: t.Role, title: t.Title, link: locked ? '' : t.Link, type: t.Type, about: t.About || '', locked: locked, done: !!done[t.Title] }; });
  var users = table_('Users'), memberNo = 0;
  users.forEach(function (x, i) { if (cleanEmail_(x.Email) === u.Email) memberNo = i + 1; });
  return {
    ok: true, user: publicUser_(u), admin: isStaff_(u), manager: isManager_(u), hasPassword: hasPassword_(u.Email), mine: mine, rank: rank, training: training, events: eventsFor_(lvl), dash: dash_(u, board), joined: fmt_(u.Joined), memberNo: memberNo, boardSize: board.length,
    leaderboard: board.slice(0, 15).map(function (b, i) { return { rank: i + 1, name: b.name, role: b.role, previews: b.previews, sales: b.sales, me: b.email === u.Email }; })
  };
}

function leaderboard_(act, sales) {
  var users = table_('Users').filter(function (x) { return x.Status === 'Active' && x.Role !== 'Admin'; });
  var salesBy = {};
  sales.forEach(function (s) { if (s.Status !== 'Refunded') { var e = cleanEmail_(s['Rep email']); salesBy[e] = (salesBy[e] || 0) + 1; } });
  return users.map(function (x) {
    var e = cleanEmail_(x.Email), parts = String(x.Name || e).trim().split(/\s+/);
    return { email: e, name: parts[0] + (parts.length > 1 ? ' ' + parts[parts.length - 1].charAt(0) + '.' : ''), role: x.Role,
             previews: Number((act[e] || {})['Previews booked'] || 0), sales: salesBy[e] || 0 };
  }).sort(function (p, q) { return q.sales - p.sales || q.previews - p.previews; });
}

/* ---------------- home dashboard: today, streak, team feed (from the Log, so it counts calls logged on the site) ---------------- */

function dash_(u, board) {
  var tz = 'Europe/London', day = function (d) { return Utilities.formatDate(new Date(d), tz, 'yyyy-MM-dd'); }, today = day(new Date());
  var names = {}; table_('Users').forEach(function (x) { names[cleanEmail_(x.Email)] = String(x.Name || '').trim().split(/\s+/)[0] || 'A rep'; });
  var sh = sheet_('Log'), rows = [];
  if (sh && sh.getLastRow() > 1) { var n = Math.min(4000, sh.getLastRow() - 1); rows = sh.getRange(sh.getLastRow() - n + 1, 1, n, 4).getValues(); }
  var calls = {}, prevWeek = 0, weekAgo = Date.now() - 7 * 86400000, feed = [], teamToday = 0;
  rows.forEach(function (r) {
    var who = cleanEmail_(r[1]), act = String(r[2]), t = new Date(r[0]).getTime(); if (!t) return;
    if (act === 'call logged') { if (who === u.Email) { var k = day(t); calls[k] = (calls[k] || 0) + 1; } if (day(t) === today) teamToday++; }
    if (act === 'preview booked') { if (who === u.Email && t >= weekAgo) prevWeek++; feed.push({ t: t, who: names[who] || 'A rep', me: who === u.Email, what: 'booked a preview' }); }
    if (act === 'added login') feed.push({ t: t, who: names[cleanEmail_(String(r[3]).split(' ')[0])] || 'A new rep', me: false, what: 'joined the team' });
  });
  table_('Sales').forEach(function (s) { var t = new Date(s.Date).getTime(); if (t && s.Status !== 'Refunded' && t > Date.now() - 30 * 86400000) { var e = cleanEmail_(s['Rep email']); feed.push({ t: t, who: names[e] || 'A rep', me: e === u.Email, what: 'closed a sale', big: true }); } });
  feed.sort(function (a, b) { return b.t - a.t; });
  var streak = 0, d = new Date(); if (!calls[today]) d.setDate(d.getDate() - 1);
  while (calls[day(d)]) { streak++; d.setDate(d.getDate() - 1); }
  var best = 0; Object.keys(calls).forEach(function (k) { best = Math.max(best, calls[k]); });
  var due = 0, endToday = new Date(); endToday.setHours(23, 59, 59, 999);
  if (sheet_('Callbacks')) table_('Callbacks').forEach(function (c) { if (cleanEmail_(c.Email) === u.Email && c.Status === 'Open' && Number(c.When) <= endToday.getTime()) due++; });
  var me = -1; board.forEach(function (b, i) { if (b.email === u.Email) me = i; });
  var above = me > 0 ? board[me - 1] : null, mine = me >= 0 ? board[me] : null;
  var week = []; for (var i = 6; i >= 0; i--) { var w = new Date(); w.setDate(w.getDate() - i); week.push({ d: Utilities.formatDate(w, tz, 'EEE'), n: calls[day(w)] || 0 }); }
  return { callsToday: calls[today] || 0, streak: streak, best: best, previewsWeek: prevWeek, callbacksDue: due, teamCallsToday: teamToday, week: week,
           chase: above && mine ? { name: above.name, sales: above.sales - mine.sales, previews: above.previews - mine.previews } : null,
           feed: feed.slice(0, 12).map(function (f) { return { ago: Math.round((Date.now() - f.t) / 60000), who: f.who, me: f.me, what: f.what, big: !!f.big }; }) };
}

/* ---------------- admin ---------------- */

function adminOverview_() {
  var users = table_('Users'), partners = table_('Partners'), sales = table_('Sales'), act = activityMap_();
  var by = {}, tot = { reps: 0, active: 0, partners: 0, sales: 0, saleValue: 0, owed: 0, paidOut: 0, pending: 0 };
  users.forEach(function (x) { by[cleanEmail_(x.Email)] = { sales: 0, earned: 0, owed: 0 }; if (x.Role !== 'Admin') { tot.reps++; if (x.Status === 'Active') tot.active++; } });
  partners.forEach(function (p) { if (p.Status === 'Active') tot.partners++; });
  sales.forEach(function (s) {
    if (s.Status === 'Refunded') return;
    var e = cleanEmail_(s['Rep email']), c = Number(s.Commission || 0), v = Number(s['Sale amount'] || 0);
    tot.sales++; tot.saleValue += v;
    if (s.Status === 'Paid out') tot.paidOut += c; else { tot.owed += c; if (s.Status === 'Pending') tot.pending += c; }
    if (!by[e]) by[e] = { sales: 0, earned: 0, owed: 0 };
    by[e].sales++; by[e].earned += c; if (s.Status !== 'Paid out') by[e].owed += c;
  });
  return {
    ok: true, totals: tot, roles: ROLES, userStatus: USER_STATUS, saleStatus: SALE_STATUS,
    callListTabs: callListTabs_(),
    users: users.map(function (x) {
      var e = cleanEmail_(x.Email), s = by[e] || { sales: 0, earned: 0, owed: 0 };
      return { email: e, name: x.Name, role: x.Role, partner: x.Partner, status: x.Status, callListTab: x['Call list tab'],
               commission: Number(x['Commission per sale'] || 0), phone: x.Phone, joined: fmt_(x.Joined), lastLogin: fmt_(x['Last login'], true), notes: x.Notes,
               previews: Number((act[e] || {})['Previews booked'] || 0), sales: s.sales, earned: s.earned, owed: s.owed };
    }),
    partners: partners.map(function (p) {
      var reps = users.filter(function (x) { return x.Partner === p.Name && x.Status === 'Active'; }).length;
      return { name: p.Name, what: p['What reps sell'], commission: Number(p['Commission per sale'] || 0), status: p.Status, started: fmt_(p.Started), contact: p.Contact, notes: p.Notes, reps: reps };
    }),
    sales: sales.slice(-60).reverse().map(function (s) {
      return { id: s.ID, date: fmt_(s.Date), rep: cleanEmail_(s['Rep email']), partner: s.Partner, customer: s['Customer / lead'], amount: Number(s['Sale amount'] || 0),
               commission: Number(s.Commission || 0), status: s.Status, source: s.Source, notes: s.Notes };
    }),
    training: table_('Training').map(function (t) { return { role: t.Role, title: t.Title, link: t.Link, type: t.Type, order: t.Order }; }),
    events: eventsFor_(99),
    activityUpdated: (function () { var v = ''; Object.keys(act).forEach(function (k) { if (act[k].Updated) v = fmt_(act[k].Updated, true); }); return v; })()
  };
}

function adminSaveUser_(me, d, invite) {
  var email = cleanEmail_(d.email);
  if (!email) return { ok: false, error: 'Add a valid email address.' };
  var name = str_(d.name, 80);
  if (!name) return { ok: false, error: 'Add their name.' };
  var role = ROLES.indexOf(d.role) !== -1 ? d.role : 'Rep';
  if (!isAdmin_(me)) { var tgt = findUser_(email); if (tgt && (tgt.Role === 'Admin' || tgt.Role === 'Manager')) return { ok: false, error: 'Only Harvey can change that login.' }; if (MANAGER_CAN_SET.indexOf(role) === -1) return { ok: false, error: 'You can set reps up to Team Lead.' }; }
  if (role === 'Admin' && email !== ADMIN_EMAIL) return { ok: false, error: 'Only ' + ADMIN_EMAIL + ' can be the admin.' };
  var status = USER_STATUS.indexOf(d.status) !== -1 ? d.status : 'Active';
  if (email === ADMIN_EMAIL) { role = 'Admin'; status = 'Active'; }          // the admin can't lock himself out
  var commission = Math.max(0, Math.min(100000, Number(d.commission) || 0));
  if (!isAdmin_(me)) { var ex = findUser_(email), pt = table_('Partners').filter(function (x) { return x.Name === d.partner; })[0]; commission = ex ? Number(ex['Commission per sale'] || 0) : Number((pt || {})['Commission per sale'] || 0); }   // pay rates are Harvey's call
  var rowData = { 'Email': email, 'Name': name, 'Role': role, 'Partner': str_(d.partner, 80), 'Status': status,
                  'Call list tab': str_(d.callListTab, 60), 'Commission per sale': commission, 'Phone': str_(d.phone, 30), 'Notes': str_(d.notes, 500) };
  var existing = findUser_(email), isNew = !existing, madeTab = '';
  if (!rowData['Call list tab'] && role !== 'Admin' && status === 'Active') { try { madeTab = createRepTab_(name); rowData['Call list tab'] = madeTab; } catch (e) { log_('system', 'could not make call list tab', name + ': ' + e.message); } }
  if (isNew) {
    rowData.Joined = new Date();
    appendObj_('Users', rowData);
  } else {
    Object.keys(rowData).forEach(function (k) { setUserField_(email, k, rowData[k]); });
  }
  log_(me.Email, isNew ? 'added login' : 'updated login', email + ' · ' + role + ' · ' + status);
  var shared = [];
  if (isNew && status === 'Active') shared = shareTraining_(email, role);
  if (invite && status === 'Active') sendLoginLink_(findUser_(email), isNew);
  var out = adminOverview_();
  out.message = (isNew ? 'Login created for ' + name + '.' : 'Saved.') + (madeTab ? ' Call list tab "' + madeTab + '" made.' : '') + (invite && status === 'Active' ? ' Login email sent to ' + email + '.' : '') +
                (shared.length ? ' Shared ' + shared.length + ' training file' + (shared.length === 1 ? '' : 's') + '.' : '');
  return out;
}

function adminSendLink_(me, email) {
  var u = findUser_(email);
  if (!u) return { ok: false, error: 'No login with that email.' };
  if (u.Status !== 'Active') return { ok: false, error: 'That login is ' + u.Status + '. Set it to Active first.' };
  if (!isAdmin_(me) && (u.Role === 'Admin' || u.Role === 'Manager') && cleanEmail_(u.Email) !== me.Email) return { ok: false, error: 'Only Harvey can do that.' };
  sendLoginLink_(u, false);
  return { ok: true, message: 'Login link sent to ' + u.Email + '.' };
}

function adminSavePartner_(me, d) {
  var name = str_(d.name, 80);
  if (!name) return { ok: false, error: 'Add the partner\'s name.' };
  var row = { 'Name': name, 'What reps sell': str_(d.what, 200), 'Commission per sale': Math.max(0, Number(d.commission) || 0),
              'Status': ['Active', 'Paused', 'Ended'].indexOf(d.status) !== -1 ? d.status : 'Active', 'Contact': str_(d.contact, 120), 'Notes': str_(d.notes, 500) };
  var sh = sheet_('Partners'), rows = sh.getDataRange().getValues(), found = 0;
  for (var i = 1; i < rows.length; i++) if (String(rows[i][0]).trim().toLowerCase() === name.toLowerCase()) found = i + 1;
  if (found) {
    var head = SHEETS.Partners;
    Object.keys(row).forEach(function (k) { sh.getRange(found, head.indexOf(k) + 1).setValue(row[k]); });
  } else {
    row.Started = new Date();
    appendObj_('Partners', row);
  }
  log_(me.Email, found ? 'updated partner' : 'added partner', name);
  var out = adminOverview_(); out.message = found ? 'Partner saved.' : 'Partner added.'; return out;
}

function adminAddSale_(me, d) {
  var rep = findUser_(d.rep);
  if (!rep) return { ok: false, error: 'Pick the rep who made the sale.' };
  var amount = Math.max(0, Number(d.amount) || 0);
  var commission = d.commission === '' || d.commission == null ? Number(rep['Commission per sale'] || 0) : Math.max(0, Number(d.commission) || 0);
  var date = d.date ? new Date(d.date) : new Date();
  if (isNaN(date.getTime())) date = new Date();
  var id = 'S' + Utilities.formatDate(new Date(), 'Europe/London', 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
  appendObj_('Sales', { 'ID': id, 'Date': date, 'Rep email': rep.Email, 'Partner': str_(d.partner, 80) || rep.Partner, 'Customer / lead': str_(d.customer, 120),
                        'Sale amount': amount, 'Commission': commission, 'Status': SALE_STATUS.indexOf(d.status) !== -1 ? d.status : 'Pending', 'Source': 'Added by admin', 'Notes': str_(d.notes, 300) });
  log_(me.Email, 'added sale', id + ' · ' + rep.Email + ' · £' + commission);
  var out = adminOverview_(); out.message = 'Sale logged for ' + rep.Name + '.'; return out;
}

function adminSetSaleStatus_(me, id, status) {
  if (SALE_STATUS.indexOf(status) === -1) return { ok: false, error: 'Unknown status.' };
  var sh = sheet_('Sales'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sh.getRange(i + 1, SHEETS.Sales.indexOf('Status') + 1).setValue(status);
      log_(me.Email, 'sale status', id + ' → ' + status);
      var out = adminOverview_(); out.message = 'Sale marked ' + status + '.'; return out;
    }
  }
  return { ok: false, error: 'Sale not found.' };
}

function adminSaveTraining_(me, d) {
  var role = d.role === 'All' || ROLES.indexOf(d.role) !== -1 ? d.role : 'Rep';
  var title = str_(d.title, 120), link = str_(d.link, 400);
  if (!title || !/^https:\/\//.test(link)) return { ok: false, error: 'Add a title and a link starting with https://' };
  appendObj_('Training', { 'Role': role, 'Title': title, 'Link': link, 'Type': str_(d.type, 20) || 'Link', 'Order': Number(d.order) || 50 });
  log_(me.Email, 'added training', role + ' · ' + title);
  var out = adminOverview_(); out.message = 'Training added.'; return out;
}

/* Share the Drive files in the training list with a new rep (view only). Files Harvey can't share are skipped. */
function shareTraining_(email, role) {
  var lvl = LEVEL[role] || 1, done = [];
  table_('Training').forEach(function (t) {
    if ((LEVEL[t.Role] || 0) > lvl) return;
    var m = String(t.Link || '').match(/\/d\/([a-zA-Z0-9_-]{20,})/);
    if (!m) return;
    try { DriveApp.getFileById(m[1]).addViewer(email); done.push(t.Title); }
    catch (e) { log_('system', 'could not share training', t.Title + ' → ' + email + ': ' + e.message); }
  });
  return done;
}

/* ---------------- hourly sync from the Call Lists sheet ---------------- */

function syncActivity() {
  var users = table_('Users').filter(function (u) { return u['Call list tab']; });
  var cl;
  try { cl = SpreadsheetApp.openById(CALL_LISTS_ID); } catch (e) { log_('system', 'sync failed', 'cannot open Call Lists: ' + e.message); return; }
  var sales = table_('Sales'), known = {};
  sales.forEach(function (s) { if (s.Source === 'Call Lists') known[String(s['Customer / lead']).split(' ')[0]] = true; });
  var now = new Date(), actRows = [];
  users.forEach(function (u) {
    var sh = cl.getSheetByName(String(u['Call list tab']).trim());
    if (!sh || sh.getLastRow() < 2) { actRows.push([cleanEmail_(u.Email), 0, 0, now]); return; }
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(CL.paid, CL.link)).getValues(), previews = 0, paid = 0;
    vals.forEach(function (r) {
      var id = String(r[CL.id - 1]).trim();
      if (!id) return;
      if (/^yes/i.test(String(r[CL.outcome - 1]).trim()) || String(r[CL.link - 1]).trim()) previews++;
      var p = String(r[CL.paid - 1]).trim();
      if (!/^paid/i.test(p)) return;
      paid++;
      if (known[id]) return;
      var amt = Number((p.match(/£\s*([\d,.]+)/) || [])[1] ? p.match(/£\s*([\d,.]+)/)[1].replace(/,/g, '') : 0);
      var dm = p.match(/(\d{2})\/(\d{2})\/(\d{4})/), when = dm ? new Date(Number(dm[3]), Number(dm[2]) - 1, Number(dm[1])) : now;
      appendObj_('Sales', { 'ID': 'CL-' + id, 'Date': when, 'Rep email': cleanEmail_(u.Email), 'Partner': u.Partner || 'HGM Leads',
                            'Customer / lead': id + ' ' + String(r[CL.biz - 1]).trim(), 'Sale amount': amt, 'Commission': Number(u['Commission per sale'] || 0),
                            'Status': 'Pending', 'Source': 'Call Lists', 'Notes': String(r[CL.pkg - 1]).trim() });
      known[id] = true;
    });
    actRows.push([cleanEmail_(u.Email), previews, paid, now]);
  });
  var ash = sheet_('Activity');
  if (ash.getLastRow() > 1) ash.getRange(2, 1, ash.getLastRow() - 1, 4).clearContent();
  if (actRows.length) ash.getRange(2, 1, actRows.length, 4).setValues(actRows);
  // Pending sales clear after the refund window
  var ssh = sheet_('Sales'), srows = ssh.getDataRange().getValues(), cStatus = SHEETS.Sales.indexOf('Status'), cDate = SHEETS.Sales.indexOf('Date');
  for (var i = 1; i < srows.length; i++) {
    if (srows[i][cStatus] === 'Pending' && srows[i][cDate] && (now - new Date(srows[i][cDate])) > CLEAR_DAYS * 86400000) ssh.getRange(i + 1, cStatus + 1).setValue('Cleared');
  }
  // Old one-time links and sessions are cleared out
  tidy_('Links', 3, 7); tidy_('Sessions', 3, 1);
}

function callListTabs_() {
  try { return SpreadsheetApp.openById(CALL_LISTS_ID).getSheets().map(function (s) { return s.getName(); }).filter(function (n) { return ['Summary'].indexOf(n) === -1; }); }
  catch (e) { return []; }
}

/* ---------------- one-time setup (run from the editor) ---------------- */

function setup() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(SHEETS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var head = SHEETS[name];
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#142D4C').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  });
  var first = ss.getSheetByName('Sheet1'); if (first && ss.getSheets().length > 1) ss.deleteSheet(first);
  ['Links', 'Sessions'].forEach(function (n) { try { ss.getSheetByName(n).hideSheet(); } catch (e) {} });
  if (!findUser_(ADMIN_EMAIL)) appendObj_('Users', { 'Email': ADMIN_EMAIL, 'Name': 'Harvey Matthews', 'Role': 'Admin', 'Partner': 'HGM Leads', 'Status': 'Active',
                                                      'Call list tab': 'Harvey', 'Commission per sale': 0, 'Joined': new Date(), 'Notes': 'Admin login' });
  if (table_('Partners').length === 0) appendObj_('Partners', { 'Name': 'HGM Leads', 'What reps sell': 'Websites for local trades (free preview first; £500 one-off or £298 today + £49 a month)',
                                                                'Commission per sale': 260, 'Status': 'Active', 'Started': new Date(), 'Contact': 'hello@maystone.uk', 'Notes': 'Founding partner: our sister web studio' });
  if (table_('Training').length === 0) {
    var d = function (id) { return 'https://drive.google.com/file/d/' + id + '/view'; }, g = function (id) { return 'https://docs.google.com/document/d/' + id + '/edit'; };
    [['All', 'Start here – how Maystone works (deck 00)', d('1JQ0pSlURxQKRSiQ9O-hV27KL91Xip_pr'), 'Deck', 1],
     ['Rep', 'Rep Start Pack – day 1', g('10ACY4FIFCEHj6DwTEp9DuHSnVGav1gjZyCgXa_yyv-w'), 'Doc', 2],
     ['Rep', 'The call script (03f)', g('1uhWBJG3UJ36SxWvOezGCw9kUfh1WPgHelQG2u7HKZT8'), 'Doc', 3],
     ['Rep', 'Deck 01 – The call script', d('15k0cqTfWfpSbsOMpYLixyMOf7WBfAQJO'), 'Deck', 4],
     ['Rep', 'Deck 02 – Objections and emotions', d('1qOfBKbAA-Sj1yefhcecAQ9D6OiiIg-O1'), 'Deck', 5],
     ['Rep', 'Deck 04 – After the yes', d('1bFkG7FpFnkxjrREPMjiHo4Q1fIGC2Ck1'), 'Deck', 6],
     ['Rep', 'Rep Training Pack v4', g('1jevyDcwmLZs31rYQVL-Doi0CYQjUEC-d89Qy0brtLoA'), 'Doc', 7],
     ['Coach', 'Deck 03 – Training Module 1 (for coaches)', d('1V6Xox6yBp2LImkkZlJQr8BNsfwgp-cOh'), 'Deck', 10],
     ['Coach', 'Deck 06 – Reps: start, pay and keep', d('18pdD4DryxRc9v34uZv1fZUTTMOxtUoG_'), 'Deck', 11],
     ['Team Lead', 'Manager Playbook', g('1Of_mMNVzy-OQFU80m3xJpdqJJ8VOyXfA7iOsoIeO0Bk'), 'Doc', 20],
     ['Admin', 'Deck 05 – Leads and compliance', d('18FQhxzDgn8PBnw3Q3XDU11PGsRBlEHiT'), 'Deck', 30]]
      .forEach(function (t) { appendObj_('Training', { 'Role': t[0], 'Title': t[1], 'Link': t[2], 'Type': t[3], 'Order': t[4] }); });
  }
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'syncActivity') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncActivity').timeBased().everyHours(1).create();
  GmailApp.getAliases(); DriveApp.getRootFolder();         // asks for every permission the portal needs, once
  syncActivity();
  return 'Setup done';
}

/* A check that runs the login flow without sending email (run from the editor). */
function selfTest() {
  var out = [];
  var u = findUser_(ADMIN_EMAIL); out.push('admin user: ' + (u ? u.Role + ' / ' + u.Status : 'MISSING'));
  var token = randomToken_();
  appendRow_('Links', [sha_(token), ADMIN_EMAIL, new Date(Date.now() + 60000), '']);
  var v = verify_(token); out.push('verify: ' + (v.ok ? 'ok' : v.error));
  var again = verify_(token); out.push('reuse blocked: ' + (!again.ok));
  if (v.ok) {
    var me = auth_(v.session); out.push('session: ' + (me ? me.Email : 'FAILED'));
    var p = mePayload_(me); out.push('me: ' + p.ok + ', training ' + p.training.length + ', admin ' + p.admin);
    var o = adminOverview_(); out.push('overview: users ' + o.users.length + ', partners ' + o.totals.partners + ', sales ' + o.totals.sales);
    endSession_(v.session); out.push('logout: ' + (auth_(v.session) ? 'FAILED' : 'ok'));
  }
  out.push('bad token refused: ' + (!verify_('abc').ok));
  console.log(out.join('\n'));
  return out.join('\n');
}

/* ---------------- leads: the rep's Call Lists tab, worked from the website ---------------- */
// The Google Sheet stays the database, so syncActivity, previews and everything else that reads it keep working.
// Reps can only change the columns below; everything else is read-only on the site.

var LEAD_EDIT = { 'Called (tick)': 60, 'Outcome': 120, 'Client email': 120, 'Call notes': 3000, 'TPS': 40, 'Next contact': 'date', 'Do not call': 'bool', 'Interested in': 80, 'OK to refer': 'bool', 'Last offer': 60, 'Details': 1000 };
function adminTab_() { var a = findUser_(ADMIN_EMAIL); return a ? String(a['Call list tab'] || '').trim() : ''; }
function leadTab_(me, tab) {
  var name = isStaff_(me) && tab ? String(tab) : String(me['Call list tab'] || '').trim();
  if (isManager_(me) && tab && name === adminTab_()) return { error: 'That call list is private.' };
  if (!name) return { error: 'You do not have a call list yet. Ask Harvey to set one up.' };
  var sh; try { sh = SpreadsheetApp.openById(CALL_LISTS_ID).getSheetByName(name); } catch (e) { return { error: 'Could not open the call lists.' }; }
  if (!sh) return { error: 'Your call list tab "' + name + '" was not found.' };
  return { sh: sh, name: name };
}
function leadVal_(x) { if (x instanceof Date) return isNaN(x.getTime()) ? '' : Utilities.formatDate(x, 'Europe/London', 'yyyy-MM-dd'); return x; }
function leads_(me, tab) {
  var t = leadTab_(me, tab); if (t.error) return { ok: false, error: t.error };
  var sh = t.sh, lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return { ok: true, tab: t.name, head: [], leads: [] };
  var v = sh.getRange(1, 1, lr, lc).getValues(), head = v[0].map(String), out = [];
  for (var i = 1; i < v.length; i++) {
    if (!String(v[i][0]).trim()) continue;
    var o = { _row: i + 1 }; head.forEach(function (h, c) { if (h) o[h] = leadVal_(v[i][c]); }); out.push(o);
  }
  var tabs = isStaff_(me) ? SpreadsheetApp.openById(CALL_LISTS_ID).getSheets().map(function (s) { return s.getName(); }).filter(function (n) { return n !== 'Summary' && !(isManager_(me) && n === adminTab_()); }) : [];
  return { ok: true, tab: t.name, tabs: tabs, head: head, editable: Object.keys(LEAD_EDIT), leads: out };
}
function repTabs_() {
  return SpreadsheetApp.openById(CALL_LISTS_ID).getSheets().filter(function (s) { return String(s.getRange(1, 1).getValue()).trim() === 'Lead ID'; });
}
/** Every lead that has a preview (or a yes), across all rep tabs, for the admin Previews list. */
function adminPreviews_() {
  var out = [];
  repTabs_().forEach(function (sh) {
    var lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2) return;
    var v = sh.getRange(1, 1, lr, lc).getValues(), head = v[0].map(String), col = function (h) { return head.indexOf(h); };
    for (var i = 1; i < v.length; i++) {
      var r = v[i], id = String(r[0]).trim(); if (!id) continue;
      var prev = String(r[col('Preview link')] || '').trim(), out1 = String(r[col('Outcome')] || '');
      if (!prev && !/^\s*yes/i.test(out1)) continue;
      var o = { tab: sh.getName(), _row: i + 1 }; head.forEach(function (h, c) { if (h) o[h] = leadVal_(r[c]); }); out.push(o);
    }
  });
  return { ok: true, previews: out };
}
/** Makes a new rep's Call Lists tab: a copy of the first rep tab's headings, checkboxes and dropdowns, with no leads. */
function createRepTab_(name) {
  var ss = SpreadsheetApp.openById(CALL_LISTS_ID), tpl = repTabs_()[0]; if (!tpl) return '';
  var base = str_(String(name).split(/\s+/)[0], 30).replace(/[^\w '-]/g, '') || 'Rep', tab = base.charAt(0).toUpperCase() + base.slice(1), n = 2;
  while (ss.getSheetByName(tab)) tab = base + ' ' + (n++);
  var sh = tpl.copyTo(ss).setName(tab);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent().clearNote();
  try { sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).concat(sh.getProtections(SpreadsheetApp.ProtectionType.SHEET)).forEach(function (p) { p.remove(); }); } catch (e) {}
  return tab;
}
function saveLead_(me, tab, row, id, f) {
  var t = leadTab_(me, tab); if (t.error) return { ok: false, error: t.error };
  var sh = t.sh, lc = sh.getLastColumn(), head = sh.getRange(1, 1, 1, lc).getValues()[0].map(String);
  row = Number(row); id = String(id || '').trim(); if (!id) return { ok: false, error: 'Missing lead.' };
  if (!(row >= 2 && row <= sh.getLastRow() && String(sh.getRange(row, 1).getValue()).trim() === id)) {
    var ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues(); row = 0;
    for (var i = 1; i < ids.length; i++) if (String(ids[i][0]).trim() === id) { row = i + 1; break; }
    if (!row) return { ok: false, error: 'That lead has moved or been removed. Refresh and try again.' };
  }
  var cur = sh.getRange(row, 1, 1, lc).getValues()[0], changed = [];
  Object.keys(f).forEach(function (h) {
    var rule = LEAD_EDIT[h] || (h === 'Approved (Harvey)' && isAdmin_(me) ? 'bool' : null), c = head.indexOf(h); if (!rule || c === -1) return;
    var val = f[h];
    if (rule === 'bool') val = val === true || val === 'true';
    else if (rule === 'date') { var m = String(val || '').match(/^(\d{4})-(\d{2})-(\d{2})$/); val = m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : ''; }
    else val = String(val == null ? '' : val).replace(/[\u0000-\u0009\u000b-\u001f]/g, ' ').trim().slice(0, rule);
    if (String(leadVal_(cur[c])) === String(leadVal_(val))) return;
    cur[c] = val; changed.push(h); sh.getRange(row, c + 1).setValue(val);
  });
  if (changed.length) {
    log_(me.Email, 'updated lead', t.name + ' ' + id + ': ' + changed.join(', '));
    var cv = cur[head.indexOf('Called (tick)')], ov = cur[head.indexOf('Outcome')];
    if (changed.indexOf('Called (tick)') !== -1 && String(cv).trim()) log_(me.Email, 'call logged', id + ' · ' + cv);
    if (changed.indexOf('Outcome') !== -1 && /^\s*yes/i.test(String(ov))) log_(me.Email, 'preview booked', id);
  }
  var o = { _row: row }; head.forEach(function (h, c) { if (h) o[h] = leadVal_(cur[c]); });
  return { ok: true, lead: o, changed: changed };
}

/* ---------------- team events: group calls, trainings (set by the admin, seen by everyone at that level) ---------------- */

function evOut_(r) { return { id: String(r.ID), title: String(r.Title || ''), when: Number(r.When) || 0, minutes: Number(r.Minutes) || 60, link: String(r.Link || ''), role: String(r.For || 'All'), details: String(r.Details || '') }; }
function eventsFor_(lvl) {
  if (!sheet_('Events')) return [];
  var from = Date.now() - 3 * 3600000;
  return table_('Events').filter(function (r) { return Number(r.When) >= from && (LEVEL[r.For] || 0) <= lvl; }).map(evOut_)
    .sort(function (x, y) { return x.when - y.when; }).slice(0, 60);
}
function adminSaveEvent_(me, d) {
  var when = Number(d.when), title = str_(d.title, 120), link = str_(d.link, 400);
  if (!title) return { ok: false, error: 'Give the event a name.' };
  if (!when || when < 1.6e12 || when > 2.2e12) return { ok: false, error: 'Pick a date and time.' };
  if (link && !/^https:\/\//.test(link)) return { ok: false, error: 'The join link must start with https://' };
  var role = d.role === 'All' || ROLES.indexOf(d.role) !== -1 ? d.role : 'All';
  var row = { Title: title, When: when, Minutes: Math.max(5, Math.min(600, Number(d.minutes) || 60)), Link: link, For: role, Details: str_(d.details, 1000) };
  var sh = ensureSheet_('Events'), head = SHEETS.Events;
  if (d.id) {
    var v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) if (String(v[i][0]) === String(d.id)) { head.forEach(function (h, c) { if (row[h] !== undefined) v[i][c] = row[h]; }); sh.getRange(i + 1, 1, 1, head.length).setValues([v[i]]); break; }
  } else { row.ID = 'ev' + Utilities.getUuid().slice(0, 8); row.Created = new Date(); appendObj_('Events', row); }
  log_(me.Email, d.id ? 'edited event' : 'added event', title);
  var out = adminOverview_(); out.message = d.id ? 'Event updated.' : 'Event added. Reps will see it on their calendar.'; return out;
}
function adminDeleteEvent_(me, id) {
  var sh = sheet_('Events'); if (sh) { var v = sh.getDataRange().getValues(); for (var i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === String(id)) { sh.deleteRow(i + 1); log_(me.Email, 'deleted event', v[i][1]); } }
  var out = adminOverview_(); out.message = 'Event deleted.'; return out;
}

/* ---------------- callbacks calendar (each rep sees only their own) ---------------- */
// 'When' is stored as epoch milliseconds so time zones never shift it.

function cbOut_(r) { return { id: String(r.ID), when: Number(r.When) || 0, name: String(r.Name || ''), business: String(r.Business || ''), phone: String(r.Phone || ''), note: String(r.Note || ''), status: String(r.Status || 'Open') }; }
function callbacks_(me) {
  if (!sheet_('Callbacks')) return { ok: true, callbacks: [] };
  var list = table_('Callbacks').filter(function (r) { return cleanEmail_(r.Email) === me.Email && r.Status !== 'Deleted'; }).map(cbOut_);
  list.sort(function (x, y) { return x.when - y.when; });
  return { ok: true, callbacks: list };
}
function saveCallback_(me, cb) {
  var when = Number(cb.when), name = str_(cb.name, 120);
  if (!when || when < 1.6e12 || when > 2.2e12) return { ok: false, error: 'Pick a date and time.' };
  if (!name) return { ok: false, error: 'Add who you are calling back.' };
  var status = ['Open', 'Done', 'No answer'].indexOf(cb.status) > -1 ? cb.status : 'Open';
  var row = { When: when, Name: name, Business: str_(cb.business, 120), Phone: str_(cb.phone, 40), Note: str_(cb.note, 1000), Status: status };
  var sh = ensureSheet_('Callbacks'), head = SHEETS.Callbacks;
  if (cb.id) {
    var v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) if (String(v[i][0]) === String(cb.id) && cleanEmail_(v[i][1]) === me.Email) {
      head.forEach(function (h, c) { if (row[h] !== undefined) v[i][c] = row[h]; });
      sh.getRange(i + 1, 1, 1, head.length).setValues([v[i]]);
      return callbacks_(me);
    }
    return { ok: false, error: 'Call-back not found.' };
  }
  row.ID = 'cb' + Utilities.getUuid().slice(0, 8); row.Email = me.Email; row.Created = new Date();
  appendObj_('Callbacks', row);
  return callbacks_(me);
}
function deleteCallback_(me, id) {
  var sh = sheet_('Callbacks'); if (!sh) return { ok: false, error: 'Call-back not found.' };
  var v = sh.getDataRange().getValues(), col = SHEETS.Callbacks.indexOf('Status') + 1;
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === String(id) && cleanEmail_(v[i][1]) === me.Email) { sh.getRange(i + 1, col).setValue('Deleted'); return callbacks_(me); }
  return { ok: false, error: 'Call-back not found.' };
}

/* ---------------- helpers ---------------- */

function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }
function table_(name) {
  var sh = sheet_(name); if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues(), head = v[0];
  return v.slice(1).filter(function (r) { return r.join('') !== ''; }).map(function (r) { var o = {}; head.forEach(function (h, i) { o[h] = r[i]; }); return o; });
}
function appendRow_(name, row) { sheet_(name).appendRow(row); }
function appendObj_(name, obj) { sheet_(name).appendRow(SHEETS[name].map(function (h) { return obj[h] === undefined ? '' : obj[h]; })); }
function findUser_(email) {
  email = cleanEmail_(email); if (!email) return null;
  var t = table_('Users');
  for (var i = 0; i < t.length; i++) if (cleanEmail_(t[i].Email) === email) { t[i].Email = email; return t[i]; }
  return null;
}
function setUserField_(email, field, value) {
  var sh = sheet_('Users'), rows = sh.getDataRange().getValues(), col = SHEETS.Users.indexOf(field) + 1;
  if (!col) return;
  for (var i = 1; i < rows.length; i++) if (cleanEmail_(rows[i][0]) === cleanEmail_(email)) { sh.getRange(i + 1, col).setValue(value); return; }
}
function activityMap_() { var m = {}; table_('Activity').forEach(function (a) { m[cleanEmail_(a['Rep email'])] = a; }); return m; }
function publicUser_(u) { return { email: cleanEmail_(u.Email), name: u.Name, role: u.Role, partner: u.Partner, status: u.Status }; }
function cleanEmail_(e) { e = String(e || '').trim().toLowerCase(); return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e) && e.length <= 120 ? e : ''; }
function str_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max || 200); }
function fmt_(d, time) { if (!d) return ''; var x = new Date(d); if (isNaN(x.getTime())) return String(d); return Utilities.formatDate(x, 'Europe/London', time ? 'd MMM yyyy, HH:mm' : 'd MMM yyyy'); }
function randomToken_() { return (Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').toLowerCase().slice(0, 64); }
function sha_(s) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s)).map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join(''); }
function log_(who, action, details) { try { appendRow_('Log', [new Date(), who, action, String(details || '').slice(0, 300)]); } catch (e) {} }
function tidy_(name, expiresCol, keepDays) {
  var sh = sheet_(name), rows = sh.getDataRange().getValues(), cutoff = Date.now() - keepDays * 86400000;
  for (var i = rows.length - 1; i >= 1; i--) if (rows[i][expiresCol - 1] && new Date(rows[i][expiresCol - 1]).getTime() < cutoff) sh.deleteRow(i + 1);
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
