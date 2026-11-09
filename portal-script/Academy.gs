/**
 * Maystone Academy – lessons, randomised tests (graded here, never in the browser) and progress for the admin.
 * Content lives in AcademyContent.gs. Every attempt, lesson and video is logged in the "Academy" tab.
 * Wired in from doPost with one line:  if (a.indexOf('academy') === 0) return json_(academy_(me, a, req, lock));
 */
var ACAD_TAB = 'Academy';
var ACAD_HEAD = ['Date', 'Email', 'Name', 'Type', 'Module', 'Score', 'Passed', 'Detail'];

function academy_(me, a, req, lock) {
  if (a === 'academy') return acadOverview_(me);
  if (a === 'academyModule') return acadModule_(me, String(req.id || ''));
  if (a === 'academyMark') { lock.waitLock(20000); return acadMark_(me, String(req.id || ''), String(req.what || '')); }
  if (a === 'academyStart') return acadStart_(me, String(req.id || ''));
  if (a === 'academySubmit') { lock.waitLock(20000); return acadSubmit_(me, String(req.token || ''), req.answers || []); }
  if (a === 'academyAdmin') { if (!isStaff_(me)) return { ok: false, error: 'Only the admin can do that.' }; return acadAdmin_(); }
  return { ok: false, error: 'Unknown request.' };
}

/* ---------- storage ---------- */
function acadSheet_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(ACAD_TAB);
  if (!sh) { sh = ss.insertSheet(ACAD_TAB); sh.appendRow(ACAD_HEAD); sh.setFrozenRows(1); sh.getRange(1, 1, 1, ACAD_HEAD.length).setFontWeight('bold'); }
  return sh;
}
function acadRows_(email) {
  var sh = acadSheet_(), n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, ACAD_HEAD.length).getValues()
    .map(function (r) { return { date: r[0], email: cleanEmail_(r[1]), name: r[2], type: String(r[3]), module: String(r[4]), score: Number(r[5]) || 0, passed: String(r[6]) === 'Yes', detail: String(r[7] || '') }; })
    .filter(function (r) { return !email || r.email === email; });
}
function acadLog_(me, type, mod, score, passed, detail) {
  acadSheet_().appendRow([new Date(), cleanEmail_(me.Email), str_(me.Name, 80), type, mod, score === '' ? '' : score, passed === '' ? '' : (passed ? 'Yes' : 'No'), str_(detail, 400)]);
}

/* ---------- progress ---------- */
function acadStatus_(rows) {
  var st = {}, vids = {};
  ACADEMY.modules.forEach(function (m) { st[m.id] = { attempts: 0, best: 0, passed: false, lesson: false, anim: false, last: [], lastDate: '' }; });
  rows.forEach(function (r) {
    if (r.type === 'Video') { vids[r.module] = true; return; }
    var s = st[r.module]; if (!s) return;
    if (r.type === 'Lesson') s.lesson = true;
    if (r.type === 'Animation') s.anim = true;
    if (r.type === 'Test') {
      s.attempts++; s.best = Math.max(s.best, r.score); if (r.passed) s.passed = true;
      var m = r.detail.match(/Q:\s*([\d,]+)/); s.last = m ? m[1].split(',') : [];
      s.lastDate = fmt_(r.date, true);
    }
  });
  return { mods: st, vids: vids };
}
function acadUnlocked_(me, st) {
  var open = {}, prev = {}, staff = isStaff_(me);
  ACADEMY.modules.forEach(function (m) { var t = acadTrack_(m); if (!(t in prev)) prev[t] = true; open[m.id] = acadCanSee_(me, m) && (staff || prev[t]); prev[t] = st.mods[m.id].passed; });
  return open;
}
function acadFind_(id) { return ACADEMY.modules.filter(function (m) { return m.id === id; })[0]; }
function acadTrack_(m) { return m.track || 'rep'; }
function acadCanSee_(me, m) { return acadTrack_(m) === 'rep' || isStaff_(me); }   // the leaders track is for Harvey and Fin only

function acadOverview_(me) {
  var st = acadStatus_(acadRows_(cleanEmail_(me.Email))), open = acadUnlocked_(me, st);
  return {
    ok: true, passMark: ACADEMY.passMark, testSize: ACADEMY.testSize,
    modules: ACADEMY.modules.filter(function (m) { return acadCanSee_(me, m); }).map(function (m) {
      var s = st.mods[m.id], n = ACADEMY.modules.filter(function (x) { return acadTrack_(x) === acadTrack_(m); }).indexOf(m) + 1;
      return { id: m.id, n: n, track: acadTrack_(m), title: m.title, tagline: m.tagline, mins: m.mins, bank: m.q.length, locked: !open[m.id], passed: s.passed, best: s.best, attempts: s.attempts, lesson: s.lesson, anim: s.anim };
    }),
    videos: ACADEMY.videos.map(function (v) { return { id: v.id, title: v.title, who: v.who, why: v.why, watched: !!st.vids[v.id] }; })
  };
}
function acadModule_(me, id) {
  var m = acadFind_(id); if (!m || !acadCanSee_(me, m)) return { ok: false, error: 'Module not found.' };
  var st = acadStatus_(acadRows_(cleanEmail_(me.Email)));
  if (!acadUnlocked_(me, st)[id]) return { ok: false, error: 'Pass the module before this one to unlock it.' };
  return { ok: true, module: { id: m.id, title: m.title, tagline: m.tagline, mins: m.mins, anim: m.anim, lesson: m.lesson } };
}
function acadMark_(me, id, what) {
  var type = { lesson: 'Lesson', anim: 'Animation', video: 'Video' }[what]; if (!type) return { ok: false, error: 'Bad request.' };
  var valid = type === 'Video' ? ACADEMY.videos.some(function (v) { return v.id === id; }) : !!acadFind_(id);
  if (!valid) return { ok: false, error: 'Not found.' };
  var done = acadRows_(cleanEmail_(me.Email)).some(function (r) { return r.type === type && r.module === id; });
  if (!done) acadLog_(me, type, id, '', '', type === 'Video' ? (ACADEMY.videos.filter(function (v) { return v.id === id; })[0].title) : '');
  return { ok: true };
}

/* ---------- tests ---------- */
function acadShuffle_(arr) { for (var i = arr.length - 1; i > 0; i--) { var j = Math.floor(Math.random() * (i + 1)), t = arr[i]; arr[i] = arr[j]; arr[j] = t; } return arr; }
function acadStart_(me, id) {
  var m = acadFind_(id); if (!m || !acadCanSee_(me, m)) return { ok: false, error: 'Module not found.' };
  var st = acadStatus_(acadRows_(cleanEmail_(me.Email)));
  if (!acadUnlocked_(me, st)[id]) return { ok: false, error: 'Pass the module before this one to unlock it.' };
  var last = {}; st.mods[id].last.forEach(function (x) { last[x] = true; });
  var fresh = acadShuffle_(m.q.filter(function (q) { return !last[q[0]]; })), seen = acadShuffle_(m.q.filter(function (q) { return last[q[0]]; }));
  var pick = fresh.concat(seen).slice(0, ACADEMY.testSize);              // never repeats the last attempt's questions
  var paper = pick.map(function (q) { return { id: q[0], order: acadShuffle_([0, 1, 2, 3]) }; });
  var token = randomToken_();
  CacheService.getScriptCache().put('acad:' + token, JSON.stringify({ email: cleanEmail_(me.Email), mod: id, paper: paper, t: Date.now() }), 7200);
  var byId = {}; m.q.forEach(function (q) { byId[q[0]] = q; });
  return { ok: true, token: token, title: m.title, passMark: ACADEMY.passMark,
    questions: paper.map(function (p, i) { var q = byId[p.id]; return { n: i + 1, q: q[2], opts: p.order.map(function (k) { return q[3][k]; }) }; }) };
}
function acadSubmit_(me, token, answers) {
  var cache = CacheService.getScriptCache(), raw = cache.get('acad:' + token);
  if (!raw) return { ok: false, error: 'This test has expired. Start a new one.' };
  var t = JSON.parse(raw); if (t.email !== cleanEmail_(me.Email)) return { ok: false, error: 'This test belongs to someone else.' };
  cache.remove('acad:' + token);
  var m = acadFind_(t.mod), byId = {}; m.q.forEach(function (q) { byId[q[0]] = q; });
  var topics = {}; m.lesson.forEach(function (b) { if (b.id) topics[b.id] = b.x; });
  var score = 0, results = t.paper.map(function (p, i) {
    var q = byId[p.id], pickIdx = Number(answers[i]), got = isNaN(pickIdx) ? -1 : p.order[pickIdx], right = got === 0;
    if (right) score++;
    return { q: q[2], ok: right, your: got >= 0 ? q[3][got] : '(no answer)', right: q[3][0], why: q[4], topic: topics[q[1]] || '' };
  });
  var passed = score >= ACADEMY.passMark;
  acadLog_(me, 'Test', m.id, score, passed, 'Q: ' + t.paper.map(function (p) { return p.id; }).join(',') + ' | wrong: ' + t.paper.filter(function (p, i) { return !results[i].ok; }).map(function (p) { return p.id; }).join(','));
  var review = passed ? results : results.map(function (r) { return { q: r.q, ok: r.ok, your: r.ok ? r.your : r.your, topic: r.topic }; });   // a fail never reveals the right answers
  var revise = {}; results.forEach(function (r) { if (!r.ok && r.topic) revise[r.topic] = true; });
  return { ok: true, score: score, total: t.paper.length, passed: passed, passMark: ACADEMY.passMark, review: review, revise: Object.keys(revise), overview: acadOverview_(me) };
}

/* ---------- admin ---------- */
function acadAdmin_() {
  var rows = acadRows_(''), all = table_('Users'), repMods = ACADEMY.modules.filter(function (m) { return acadTrack_(m) === 'rep'; }), leadMods = ACADEMY.modules.filter(function (m) { return acadTrack_(m) === 'lead'; });
  var staffRole = function (u) { return u.Role === 'Admin' || u.Role === 'Manager'; };
  var person = function (u, mods) {
    var email = cleanEmail_(u.Email), mine = rows.filter(function (r) { return r.email === email; }), st = acadStatus_(mine);
    var last = mine.length ? mine[mine.length - 1].date : '';
    return {
      name: u.Name, email: email, role: u.Role, status: u.Status,
      passed: mods.filter(function (m) { return st.mods[m.id].passed; }).length,
      attempts: mine.filter(function (r) { return r.type === 'Test' && mods.some(function (m) { return m.id === r.module; }); }).length,
      videos: Object.keys(st.vids).length, lastActive: last ? fmt_(last, true) : '', lastTs: last ? new Date(last).getTime() : 0,
      mods: mods.map(function (m) { var s = st.mods[m.id]; return { id: m.id, passed: s.passed, best: s.best, attempts: s.attempts, lesson: s.lesson, anim: s.anim, last: s.lastDate }; })
    };
  };
  var reps = all.filter(function (u) { return u.Role !== 'Admin'; }).map(function (u) { return person(u, repMods); });
  var leaders = all.filter(staffRole).map(function (u) { return person(u, leadMods); });
  var recent = rows.slice(-40).reverse().map(function (r) { return { when: fmt_(r.date, true), name: r.name || r.email, type: r.type, module: (acadFind_(r.module) || {}).title || r.detail || r.module, score: r.score, passed: r.passed }; });
  var num = function (list) { return list.map(function (m, i) { return { id: m.id, n: i + 1, title: m.title }; }); };
  return { ok: true, modules: num(repMods), leadModules: num(leadMods), videos: ACADEMY.videos.length, passMark: ACADEMY.passMark, reps: reps, leaders: leaders, risk: acadRisk_(all, rows, repMods.length), recent: recent };
}

/* Reps who've gone quiet: no login, no calls, no preview yes, or a stalled Academy. */
function acadRisk_(users, acadRows, repModCount) {
  var now = Date.now(), DAY = 86400000, days = function (t) { return t ? Math.floor((now - t) / DAY) : null; };
  var lastCall = {}, previews = {};
  try {
    var sh = sheet_('Log');
    if (sh && sh.getLastRow() > 1) {
      var n = Math.min(6000, sh.getLastRow() - 1);
      sh.getRange(sh.getLastRow() - n + 1, 1, n, 3).getValues().forEach(function (r) {
        var who = cleanEmail_(r[1]), act = String(r[2]), t = new Date(r[0]).getTime(); if (!t) return;
        if (act === 'call logged') lastCall[who] = Math.max(lastCall[who] || 0, t);
        if (act === 'preview booked') previews[who] = (previews[who] || 0) + 1;
      });
    }
  } catch (e) {}
  var out = [];
  users.forEach(function (u) {
    if (u.Role === 'Admin' || u.Role === 'Manager' || u.Status !== 'Active') return;
    var email = cleanEmail_(u.Email), joined = u.Joined ? new Date(u.Joined).getTime() : 0, jd = joined ? days(joined) : null;
    var login = u['Last login'] ? new Date(u['Last login']).getTime() : 0, ld = days(login), cd = days(lastCall[email] || 0);
    var mine = acadRows.filter(function (r) { return r.email === email; }), ad = mine.length ? days(new Date(mine[mine.length - 1].date).getTime()) : null;
    var passed = Object.keys(acadStatus_(mine).mods).filter(function (k) { var m = acadFind_(k); return m && acadTrack_(m) === 'rep' && acadStatus_(mine).mods[k].passed; }).length;
    var why = [];
    if (!login) { if (jd === null || jd >= 1) why.push('Never logged in'); } else if (ld >= 3) why.push('Not logged in for ' + ld + ' days');
    if (!lastCall[email]) { if (jd !== null && jd >= 3) why.push('No calls logged since joining'); } else if (cd >= 3) why.push('No calls logged for ' + cd + ' days');
    if (!previews[email] && jd !== null && jd >= 3) why.push('No preview yes yet (day ' + jd + ')');
    if (passed < repModCount) { if (ad === null) { if (jd === null || jd >= 2) why.push('Hasn\'t started the Academy'); } else if (ad >= 3) why.push('Academy stalled for ' + ad + ' days (' + passed + '/' + repModCount + ' passed)'); }
    if (why.length) out.push({ name: u.Name, email: email, phone: u.Phone || '', reasons: why, joined: joined ? fmt_(joined) : '' });
  });
  return out.sort(function (a, b) { return b.reasons.length - a.reasons.length; });
}
