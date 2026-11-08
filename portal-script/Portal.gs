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
  Log: ['Time', 'Who', 'Action', 'Details']
};
var ROLES = ['Rep', 'Top Rep', 'Coach', 'Team Lead', 'Admin'];
var LEVEL = { 'All': 0, 'Rep': 1, 'Top Rep': 2, 'Coach': 3, 'Team Lead': 4, 'Admin': 9 };
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
    if (a === 'markDone') { lock.waitLock(20000); return json_(markDone_(me, req.title, !!req.done)); }
    if (!isAdmin_(me)) return json_({ ok: false, error: 'Only the admin can do that.' });
    lock.waitLock(20000);
    if (a === 'adminOverview') return json_(adminOverview_());
    if (a === 'adminSaveUser') return json_(adminSaveUser_(me, req.user || {}, !!req.invite));
    if (a === 'adminSendLink') return json_(adminSendLink_(me, req.email));
    if (a === 'adminSavePartner') return json_(adminSavePartner_(me, req.partner || {}));
    if (a === 'adminAddSale') return json_(adminAddSale_(me, req.sale || {}));
    if (a === 'adminSetSaleStatus') return json_(adminSetSaleStatus_(me, req.id, req.status));
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
    ok: true, user: publicUser_(u), admin: isAdmin_(u), hasPassword: hasPassword_(u.Email), mine: mine, rank: rank, training: training, joined: fmt_(u.Joined), memberNo: memberNo, boardSize: board.length,
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
    activityUpdated: (function () { var v = ''; Object.keys(act).forEach(function (k) { if (act[k].Updated) v = fmt_(act[k].Updated, true); }); return v; })()
  };
}

function adminSaveUser_(me, d, invite) {
  var email = cleanEmail_(d.email);
  if (!email) return { ok: false, error: 'Add a valid email address.' };
  var name = str_(d.name, 80);
  if (!name) return { ok: false, error: 'Add their name.' };
  var role = ROLES.indexOf(d.role) !== -1 ? d.role : 'Rep';
  if (role === 'Admin' && email !== ADMIN_EMAIL) return { ok: false, error: 'Only ' + ADMIN_EMAIL + ' can be the admin.' };
  var status = USER_STATUS.indexOf(d.status) !== -1 ? d.status : 'Active';
  if (email === ADMIN_EMAIL) { role = 'Admin'; status = 'Active'; }          // the admin can't lock himself out
  var commission = Math.max(0, Math.min(100000, Number(d.commission) || 0));
  var rowData = { 'Email': email, 'Name': name, 'Role': role, 'Partner': str_(d.partner, 80), 'Status': status,
                  'Call list tab': str_(d.callListTab, 60), 'Commission per sale': commission, 'Phone': str_(d.phone, 30), 'Notes': str_(d.notes, 500) };
  var existing = findUser_(email), isNew = !existing;
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
  out.message = (isNew ? 'Login created for ' + name + '.' : 'Saved.') + (invite && status === 'Active' ? ' Login email sent to ' + email + '.' : '') +
                (shared.length ? ' Shared ' + shared.length + ' training file' + (shared.length === 1 ? '' : 's') + '.' : '');
  return out;
}

function adminSendLink_(me, email) {
  var u = findUser_(email);
  if (!u) return { ok: false, error: 'No login with that email.' };
  if (u.Status !== 'Active') return { ok: false, error: 'That login is ' + u.Status + '. Set it to Active first.' };
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
     ['Rep', 'Deck 02 – Objections and emotions', d('1qDOfBKbAA-Sj1yefhcecAQ9D6OiiIg-O1'), 'Deck', 5],
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
