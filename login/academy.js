/* Maystone Academy – course, animated explainers, randomised tests and admin progress.
   Talks to the portal backend through window.MSBridge.api (set up in index.html). Content and grading live on the server. */
(function () {
  'use strict';
  var B = function () { return window.MSBridge || {}; };
  var api = function (a, d) { return B().api ? B().api(a, d) : Promise.resolve({ ok: false, error: 'Not ready.' }); };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); };
  var DEMOS = {
    DEMO1: { img: 'academy/img/demo1.jpg', name: 'Harbour Heating & Plumbing', trade: 'Plumber', url: 'https://hgmbn43-stack.github.io/maystone-previews/demo/DEMO1-harbour-heating-plumbing/' },
    DEMO2: { img: 'academy/img/demo2.jpg', name: 'Downland Roofing', trade: 'Roofer', url: 'https://hgmbn43-stack.github.io/maystone-previews/demo/DEMO2-downland-roofing/' },
    DEMO5: { img: 'academy/img/demo5.jpg', name: 'Brightwire Electrical', trade: 'Electrician', url: 'https://hgmbn43-stack.github.io/maystone-previews/demo/DEMO5-brightwire-electrical/' }
  };
  var FACTS = [
    ['Testing yourself beats re-reading: 56% vs 42% remembered a week later.', 'Roediger & Karpicke, 2006'],
    ['The average UK commute hit 59 minutes a day in 2018 – about 221 hours a year. Yours is zero.', 'TUC, 2019'],
    ['Saying the reason for your call made cold calls 2.1× more successful.', 'Gong, 90,380 cold calls'],
    ['In won deals, reps talked less: 57% of the call vs 62% in lost ones.', 'Gong, 326,000 sales calls'],
    ['"Did I catch you at a bad time?" made booking a meeting 40% less likely.', 'Gong, 90,380 cold calls']
  ];
  var S = { root: null, opts: {}, ov: null, view: 'course', mod: null, cache: {}, test: null, result: null, player: null };

  /* ================= rep: mount ================= */
  function mount(root, opts) {
    S.root = root; S.opts = opts || {};
    if (S.player) S.player.stop();
    if (S.ov) return render();
    root.innerHTML = loading('Loading the Academy…');
    api('academy').then(function (r) {
      if (!r.ok) { root.innerHTML = '<div class="ms-panel"><p class="ms-p">' + esc(r.error || 'Could not load the Academy.') + '</p></div>'; return; }
      S.ov = r; render();
    });
  }
  function loading(t) { return '<div class="ms-panel ac-loadp"><div class="ms-load"><i></i><i></i><i></i><span>' + esc(t) + '</span></div></div>'; }
  function go(view, mod) { if (S.player) S.player.stop(); S.view = view; if (mod !== undefined) S.mod = mod; render(); try { S.root.scrollIntoView({ block: 'start' }); window.scrollTo(0, 0); } catch (e) {} }
  function render() {
    var v = S.view, h = '';
    if (v === 'course') h = courseHTML();
    else if (v === 'module') return renderModule();
    else if (v === 'test') h = testHTML();
    else if (v === 'result') h = resultHTML();
    else if (v === 'videos') h = videosHTML();
    else if (v === 'library') h = libraryHTML();
    S.root.innerHTML = h; bind();
  }

  /* ================= course overview ================= */
  function nextMod(track) { return S.ov.modules.filter(function (m) { return (m.track || 'rep') === (track || 'rep') && !m.passed && !m.locked; })[0]; }
  function trackMods(t) { return S.ov.modules.filter(function (m) { return (m.track || 'rep') === t; }); }
  function pathHTML(list, nx) {
    return '<div class="ac-path">' + list.map(function (m, i) {
      var lead = m.track === 'lead', stt = m.passed ? 'passed' : m.locked ? 'locked' : (nx && nx.id === m.id) ? 'now' : 'open';
      var badge = m.passed ? '<span class="ac-badge ok">✓ Passed · best ' + m.best + '/10</span>' : m.locked ? '<span class="ac-badge">🔒 Pass module ' + i + ' to unlock</span>' : m.attempts ? '<span class="ac-badge warn">Best ' + m.best + '/10 · ' + m.attempts + ' attempt' + (m.attempts === 1 ? '' : 's') + '</span>' : '<span class="ac-badge new">' + m.mins + ' min</span>';
      var steps = '<span class="ac-st' + (m.anim ? ' on' : '') + '">▶ Watch</span><span class="ac-st' + (m.lesson ? ' on' : '') + '">✎ Learn</span><span class="ac-st' + (m.passed ? ' on' : '') + '">★ Test</span>';
      return '<button class="ac-node ' + stt + (lead ? ' lead' : '') + '" ' + (m.locked ? 'disabled' : 'data-mod="' + m.id + '"') + '><span class="ac-num">' + (m.passed ? '✓' : (lead ? 'L' : '') + m.n) + '</span><span class="ac-nb"><small>' + (lead ? 'Leaders · module ' : 'Module ') + m.n + '</small><b>' + esc(m.title) + '</b><em>' + esc(m.tagline) + '</em><span class="ac-steps">' + steps + '</span></span><span class="ac-nr">' + badge + '</span></button>';
    }).join('') + '</div>';
  }
  function courseHTML() {
    var o = S.ov, repM = trackMods('rep'), leadM = trackMods('lead'), done = repM.filter(function (m) { return m.passed; }).length, tot = repM.length, pct = tot ? done / tot : 0;
    var nx = nextMod(), fact = FACTS[Math.floor(Date.now() / 864e5) % FACTS.length], vw = o.videos.filter(function (v) { return v.watched; }).length;
    var r = 52, c = 2 * Math.PI * r;
    return '<section class="ac-hero"><div class="ac-hero-l"><p class="ms-kick">Maystone Academy</p><h2>' + (done === tot ? 'You\'ve passed every module.' : done ? 'Keep going. You\'re ' + done + ' of ' + tot + ' in.' : 'Learn it. Prove it. Sell it.') + '</h2>' +
      '<p class="ac-sub">Watch the video, learn the lesson, then pass the test: 10 questions from a bank of 30, and you need ' + o.passMark + '. Fail and you get different questions next time – so the only way through is to actually learn it.</p>' +
      (nx ? '<button class="ms-btn acc" data-mod="' + nx.id + '">' + (nx.attempts || nx.lesson ? 'Continue' : 'Start') + ': ' + esc(nx.title) + ' →</button>' : '<button class="ms-btn acc" data-go2="videos">Learn from the best →</button>') +
      '</div><div class="ac-hero-r"><svg width="132" height="132" viewBox="0 0 132 132" aria-hidden="true"><circle cx="66" cy="66" r="' + r + '" fill="none" stroke="rgba(255,255,255,.12)" stroke-width="11"/><circle class="ac-arc" cx="66" cy="66" r="' + r + '" fill="none" stroke="#FF6A2E" stroke-width="11" stroke-linecap="round" stroke-dasharray="' + c + '" stroke-dashoffset="' + c * (1 - pct) + '" transform="rotate(-90 66 66)"/></svg><div class="ac-ringl"><b>' + done + '/' + tot + '</b><small>passed</small></div></div></section>' +
      '<div class="ac-fact"><span class="ac-fact-i">⚡</span><p><b>Did you know?</b> ' + esc(fact[0]) + ' <small>' + esc(fact[1]) + '</small></p></div>' +
      pathHTML(repM, nx) +
      (leadM.length ? '<div class="ac-lhead"><span>Leaders only</span><h3>The leaders track</h3><p>For Harvey and Fin: teaching new reps, coaching, motivation and keeping reps from quitting. Reps can\'t see this.</p></div>' + pathHTML(leadM, nextMod('lead')) : '') +
      '<div class="ac-two"><button class="ac-tile" data-go2="videos"><span class="ac-ti">▶</span><b>Learn from the best</b><p>Hand-picked videos from top cold-calling trainers.</p><em>' + vw + ' of ' + o.videos.length + ' watched</em></button>' +
      '<button class="ac-tile" data-go2="library"><span class="ac-ti">📄</span><b>Reference library</b><p>The full call script and decks, word for word.</p><em>Open any time</em></button></div>';
  }

  /* ================= module page ================= */
  function renderModule() {
    var id = S.mod, meta = S.ov.modules.filter(function (m) { return m.id === id; })[0];
    if (!meta) return go('course');
    var paint = function (r) {
      if (!r.ok) { S.root.innerHTML = '<div class="ms-panel"><p class="ms-p">' + esc(r.error) + '</p><button class="ms-btn ghost" data-go2="course">← Back</button></div>'; bind(); return; }
      S.cache[id] = r; var m = r.module;
      S.root.innerHTML = '<div class="ac-mhead ms-phead"><div><button class="ms-link back" data-go2="course">← All modules</button><p class="ms-kick dark">' + (meta.track === 'lead' ? 'Leaders · module ' : 'Module ') + meta.n + ' · about ' + m.mins + ' minutes</p><h2>' + esc(m.title) + '</h2><p class="ms-lead">' + esc(m.tagline) + '</p></div>' +
        '<div class="ac-flow"><a href="#ac-watch" class="ac-fl' + (meta.anim ? ' on' : '') + '"><b>1</b>Watch</a><a href="#ac-learn" class="ac-fl' + (meta.lesson ? ' on' : '') + '"><b>2</b>Learn</a><a href="#ac-test" class="ac-fl' + (meta.passed ? ' on' : '') + '"><b>3</b>Test</a></div></div>' +
        '<section id="ac-watch" class="ac-sec"><div class="ac-sech"><h3>1. Watch: ' + esc(m.anim.title) + '</h3><span>About a minute</span></div><div class="ac-player" id="ac-player"></div></section>' +
        '<section id="ac-learn" class="ac-sec ms-panel"><div class="ac-sech"><h3>2. Learn</h3><span>Grab a notebook – write down every orange box</span></div><article class="ac-lesson">' + lessonHTML(m.lesson) + '</article>' +
        '<div class="ac-done"><button class="ms-btn ' + (meta.lesson ? 'ghost' : 'acc') + '" data-read="' + id + '">' + (meta.lesson ? '✓ Lesson read' : 'I\'ve read it and made my notes ✓') + '</button></div></section>' +
        '<section id="ac-test" class="ac-sec ac-testcta"><div><h3>3. Test yourself</h3><p>10 questions from a bank of 30. You need ' + S.ov.passMark + ' to pass. Get it wrong and you\'ll get different questions next time.</p>' + (meta.attempts ? '<p class="ac-small">Your best so far: <b>' + meta.best + '/10</b> from ' + meta.attempts + ' attempt' + (meta.attempts === 1 ? '' : 's') + '.' + (meta.passed ? ' Passed ✓' : '') + '</p>' : '') + '</div><button class="ms-btn acc" data-start="' + id + '">' + (meta.passed ? 'Take it again' : meta.attempts ? 'Try again – new questions' : 'Start the test') + ' →</button></section>';
      bind();
      S.player = player(document.getElementById('ac-player'), m.anim, function () { if (!meta.anim) { meta.anim = true; api('academyMark', { id: id, what: 'anim' }); var f = S.root.querySelector('.ac-fl'); if (f) f.classList.add('on'); } });
    };
    if (S.cache[id]) return paint(S.cache[id]);
    S.root.innerHTML = loading('Loading module…');
    api('academyModule', { id: id }).then(paint);
  }
  function lessonHTML(blocks) {
    return blocks.map(function (b) {
      switch (b.t) {
        case 'h': return '<h3 id="ac-t-' + esc(b.id) + '">' + esc(b.x) + '</h3>';
        case 'p': return '<p>' + esc(b.x) + '</p>';
        case 'mem': return '<div class="ac-mem"><span>✎ Remember this – write it down</span><p>' + esc(b.x) + '</p></div>';
        case 'say': return '<blockquote class="ac-say"><span>Say it like this</span>' + esc(b.x) + '</blockquote>';
        case 'tip': return '<div class="ac-tip"><b>Tip</b> ' + esc(b.x) + '</div>';
        case 'stat': return '<div class="ac-stat"><span>📊</span><p>' + esc(b.x) + '<small>Source: ' + esc(b.src) + '</small></p></div>';
        case 'list': return '<ul>' + b.x.map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') + '</ul>';
        case 'never': return '<div class="ac-never"><b>Never</b><ul>' + b.x.map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') + '</ul></div>';
        case 'table': return '<div class="ac-tw"><table>' + b.rows.map(function (r, i) { return '<tr>' + r.map(function (c) { return i ? '<td>' + esc(c) + '</td>' : '<th>' + esc(c) + '</th>'; }).join('') + '</tr>'; }).join('') + '</table></div>';
        case 'demo': return '<p>' + esc(b.x) + '</p><div class="ac-demos">' + b.imgs.map(function (k) { var d = DEMOS[k]; return d ? '<figure><div class="ac-phone"><img src="' + d.img + '" alt="Example preview for ' + esc(d.name) + ', a made-up ' + esc(d.trade.toLowerCase()) + '" loading="lazy"></div><figcaption><b>' + esc(d.name) + '</b><span>Example ' + esc(d.trade.toLowerCase()) + ' – not a real business</span><a href="' + d.url + '" target="_blank" rel="noopener">Open the full example →</a></figcaption></figure>' : ''; }).join('') + '</div>';
        default: return '';
      }
    }).join('');
  }

  /* ================= animated explainer player ================= */
  function player(el, anim, onDone) {
    var scenes = anim.scenes, i = 0, timer = null, playing = false, voice = false, finished = false;
    try { voice = localStorage.getItem('maystone.acvoice') === '1'; } catch (e) {}
    var canSpeak = 'speechSynthesis' in window;
    el.innerHTML = '<div class="ac-stage" tabindex="0"><div class="ac-bars">' + scenes.map(function () { return '<i><b></b></i>'; }).join('') + '</div><div class="ac-scene" aria-live="polite"></div>' +
      '<div class="ac-cap"></div><button class="ac-big" aria-label="Play">▶</button></div>' +
      '<div class="ac-ctrl"><button data-c="prev" aria-label="Previous scene">⏮</button><button data-c="play" aria-label="Play or pause">▶</button><button data-c="next" aria-label="Next scene">⏭</button><button data-c="restart" aria-label="Restart">↺</button>' +
      (canSpeak ? '<button data-c="voice" class="ac-voice' + (voice ? ' on' : '') + '">' + (voice ? '🔊 Voice on' : '🔈 Voice off') + '</button>' : '') + '<span class="ac-sc"></span></div>';
    var sceneEl = el.querySelector('.ac-scene'), capEl = el.querySelector('.ac-cap'), bars = el.querySelectorAll('.ac-bars i'), big = el.querySelector('.ac-big'), pb = el.querySelector('[data-c="play"]'), sc = el.querySelector('.ac-sc');
    function dur(s) { return Math.max(4200, s.cap.split(/\s+/).length * 330 + (s.k === 'chat' ? s.lines.length * 900 : 900)); }
    function show(n) {
      i = Math.max(0, Math.min(scenes.length - 1, n)); var s = scenes[i];
      sceneEl.innerHTML = sceneHTML(s); sceneEl.className = 'ac-scene k-' + s.k; capEl.textContent = s.cap; sc.textContent = (i + 1) + ' / ' + scenes.length;
      Array.prototype.forEach.call(bars, function (b, k) { b.className = k < i ? 'done' : ''; b.firstChild.style.animation = 'none'; });
      if (s.k === 'money') countUp(sceneEl.querySelector('[data-to]'));
      if (playing) run();
    }
    function run() {
      clearTimeout(timer); if (canSpeak) speechSynthesis.cancel();
      var s = scenes[i], d = dur(s), b = bars[i].firstChild;
      bars[i].className = 'on'; b.style.animation = 'none'; void b.offsetWidth; b.style.animation = 'acbar ' + d + 'ms linear forwards';
      var next = function () { if (!playing) return; if (i < scenes.length - 1) show(i + 1); else end(); };
      if (voice && canSpeak) {
        var u = new SpeechSynthesisUtterance(s.cap); u.lang = 'en-GB'; u.rate = 1.02;
        var gb = speechSynthesis.getVoices().filter(function (v) { return /en-GB/i.test(v.lang); })[0]; if (gb) u.voice = gb;
        var t0 = Date.now(); u.onend = function () { timer = setTimeout(next, Math.max(600, d - (Date.now() - t0))); };
        speechSynthesis.speak(u);
        timer = setTimeout(next, d + 9000);                   // safety net if the voice never ends
      } else timer = setTimeout(next, d);
    }
    function play() { if (finished) { finished = false; show(0); } playing = true; big.hidden = true; pb.textContent = '❚❚'; run(); }
    function pause() { playing = false; clearTimeout(timer); if (canSpeak) speechSynthesis.cancel(); pb.textContent = '▶'; var b = bars[i].firstChild; var w = getComputedStyle(b).width; b.style.animation = 'none'; b.style.width = w; }
    function end() { playing = false; finished = true; bars[i].className = 'done'; pb.textContent = '↺'; big.hidden = false; big.textContent = '↺'; big.setAttribute('aria-label', 'Watch again'); if (onDone) onDone(); }
    big.onclick = play;
    el.querySelector('.ac-ctrl').onclick = function (e) {
      var c = e.target.closest('[data-c]'); if (!c) return; c = c.dataset.c;
      if (c === 'play') { playing ? pause() : play(); }
      if (c === 'next') { if (i < scenes.length - 1) show(i + 1); else if (playing) end(); }
      if (c === 'prev') show(i - 1);
      if (c === 'restart') { finished = false; show(0); play(); }
      if (c === 'voice') { voice = !voice; try { localStorage.setItem('maystone.acvoice', voice ? '1' : '0'); } catch (x) {} var vb = el.querySelector('.ac-voice'); vb.classList.toggle('on', voice); vb.textContent = voice ? '🔊 Voice on' : '🔈 Voice off'; if (playing) run(); else if (!voice && canSpeak) speechSynthesis.cancel(); }
    };
    show(0);
    return { stop: function () { playing = false; clearTimeout(timer); if (canSpeak) speechSynthesis.cancel(); } };
  }
  function countUp(n) {
    if (!n) return; var to = +n.dataset.to, t0 = performance.now();
    (function tick(t) { var k = Math.min(1, (t - t0) / 1400); n.textContent = Math.round(to * (1 - Math.pow(1 - k, 3))); if (k < 1 && n.isConnected) requestAnimationFrame(tick); })(t0);
  }
  function sceneHTML(s) {
    var d = function (k) { return ' style="--d:' + k + 'ms"'; };
    switch (s.k) {
      case 'title': return '<div class="sc-title"><span class="sc-logo"><i></i><i></i><i></i><i></i><i></i><i class="a"></i></span><h4' + d(250) + '>' + esc(s.big) + '</h4><p' + d(700) + '>' + esc(s.sub) + '</p></div>';
      case 'steps': return '<div class="sc-steps"><h5>' + esc(s.title) + '</h5>' + s.items.map(function (x, k) { return '<div class="sc-step"' + d(400 + k * 700) + '><b>' + (k + 1) + '</b>' + esc(x) + '</div>'; }).join('') + '</div>';
      case 'money': return '<div class="sc-money"><div class="sc-coin"' + d(0) + '>£</div><b>£<span data-to="' + s.amount + '">0</span></b><p' + d(900) + '>' + esc(s.label) + '</p></div>';
      case 'stat': return '<div class="sc-stat"><b' + d(100) + '>' + esc(s.num) + '</b><p' + d(700) + '>' + esc(s.label) + '</p><small' + d(1300) + '>' + esc(s.src) + '</small></div>';
      case 'ladder': return '<div class="sc-ladder">' + s.items.map(function (x, k) { return '<div class="sc-rung"' + d(300 + k * 550) + ' data-h="' + (k + 1) + '"><b>' + (k + 1) + '</b><span>' + esc(x) + '</span></div>'; }).join('') + '</div>';
      case 'rules': return '<div class="sc-rules">' + s.items.map(function (x, k) { return '<div class="sc-rule ' + x[0] + '"' + d(300 + k * 650) + '><b>' + (x[0] === 'ok' ? '✓' : '✕') + '</b>' + esc(x[1]) + '</div>'; }).join('') + '</div>';
      case 'price': return '<div class="sc-price"><div class="sc-pc"' + d(200) + '><small>One-off</small><b>£500</b><p>Theirs outright · domain in their name</p><em>+ £10 a month hosting after year 1</em></div><div class="sc-or"' + d(800) + '>or</div><div class="sc-pc hl"' + d(1100) + '><small>Monthly</small><b>£298 <span>today</span></b><p>then £49 a month · 30 days\' notice</p><em>Up to 3 small changes a month</em></div></div>';
      case 'chat': return '<div class="sc-call"><div class="sc-callh"><span class="sc-dot"></span>On a call · Maystone</div>' + s.lines.map(function (l, k) { return '<div class="sc-b ' + l[0] + '"' + d(300 + k * 1600) + '><small>' + (l[0] === 'rep' ? 'You' : 'Owner') + '</small>' + esc(l[1]) + '</div>'; }).join('') + '</div>';
      case 'site': var ini = s.name.split(/\s+/).map(function (w) { return w.charAt(0); }).join('').slice(0, 2);
        return '<div class="sc-site"><div class="sc-ph"><div class="sc-rib"' + d(100) + '>Website preview, made by Maystone</div><div class="sc-hd"' + d(400) + '><span>' + esc(ini) + '</span>' + esc(s.name) + '</div><div class="sc-chip"' + d(800) + '>' + esc(s.trade) + '</div><h6' + d(1100) + '>' + esc(s.name) + '</h6><div class="sc-stars"' + d(1600) + '>★★★★★ <small>4.9 · 31 Google reviews</small></div><div class="sc-tel"' + d(2100) + '>07700 900417</div><div class="sc-btns"' + d(2600) + '><i>📞 Call now</i><i class="g">WhatsApp a photo</i></div><div class="sc-q"' + d(3100) + '>Get a quote</div></div><p class="sc-ex">Example – made-up business</p></div>';
      default: return '';
    }
  }

  /* ================= test ================= */
  function startTest(id) {
    if (S.player) S.player.stop();
    S.root.innerHTML = loading('Picking your 10 questions…');
    api('academyStart', { id: id }).then(function (r) {
      if (!r.ok) { S.root.innerHTML = '<div class="ms-panel"><p class="ms-p">' + esc(r.error) + '</p><button class="ms-btn ghost" data-mod="' + id + '">← Back</button></div>'; bind(); return; }
      S.test = { id: id, token: r.token, title: r.title, qs: r.questions, ans: r.questions.map(function () { return null; }), at: 0 };
      go('test');
    });
  }
  function testHTML() {
    var t = S.test, q = t.qs[t.at], L = 'ABCD', answered = t.ans.filter(function (a) { return a !== null; }).length;
    return '<div class="ac-test"><div class="ac-th"><button class="ms-link back" data-quit="1">← Leave test</button><span>' + esc(t.title) + ' · test</span></div>' +
      '<div class="ac-qdots">' + t.qs.map(function (x, k) { return '<button class="' + (k === t.at ? 'now ' : '') + (t.ans[k] !== null ? 'done' : '') + '" data-jump="' + k + '" aria-label="Question ' + (k + 1) + '">' + (k + 1) + '</button>'; }).join('') + '</div>' +
      '<div class="ac-qcard"><small>Question ' + (t.at + 1) + ' of ' + t.qs.length + '</small><h3>' + esc(q.q) + '</h3><div class="ac-opts">' + q.opts.map(function (o, k) { return '<button class="ac-opt' + (t.ans[t.at] === k ? ' sel' : '') + '" data-opt="' + k + '"><b>' + L[k] + '</b><span>' + esc(o) + '</span></button>'; }).join('') + '</div></div>' +
      '<div class="ac-tnav"><button class="ms-btn ghost" data-step="-1"' + (t.at ? '' : ' disabled') + '>← Back</button>' +
      (t.at < t.qs.length - 1 ? '<button class="ms-btn acc" data-step="1"' + (t.ans[t.at] === null ? ' disabled' : '') + '>Next →</button>' : '<button class="ms-btn acc" data-submit="1"' + (answered < t.qs.length ? ' disabled' : '') + '>Submit my answers ✓</button>') + '</div>' +
      '<p class="ac-small ac-center">' + answered + ' of ' + t.qs.length + ' answered · you need ' + S.ov.passMark + ' right to pass</p></div>';
  }
  function submit() {
    var t = S.test; S.root.innerHTML = loading('Marking your test…');
    api('academySubmit', { token: t.token, answers: t.ans }).then(function (r) {
      if (!r.ok) { S.root.innerHTML = '<div class="ms-panel"><p class="ms-p">' + esc(r.error) + '</p><button class="ms-btn acc" data-start="' + t.id + '">Start a new test</button></div>'; bind(); return; }
      S.result = r; S.result.id = t.id; if (r.overview) S.ov = r.overview; S.test = null; go('result');
      if (r.passed) confetti();
    });
  }
  function resultHTML() {
    var r = S.result, meta = S.ov.modules.filter(function (m) { return m.id === r.id; })[0], nx = nextMod(meta && meta.track), pct = r.score / r.total, R = 58, c = 2 * Math.PI * R;
    var head = '<section class="ac-res ' + (r.passed ? 'pass' : 'fail') + '"><svg width="150" height="150" viewBox="0 0 150 150" aria-hidden="true"><circle cx="75" cy="75" r="' + R + '" fill="none" stroke="rgba(255,255,255,.14)" stroke-width="12"/><circle class="ac-arc" cx="75" cy="75" r="' + R + '" fill="none" stroke="' + (r.passed ? '#3DDC97' : '#FF6A2E') + '" stroke-width="12" stroke-linecap="round" stroke-dasharray="' + c + '" stroke-dashoffset="' + c * (1 - pct) + '" transform="rotate(-90 75 75)"/></svg><div class="ac-resl"><b>' + r.score + '/' + r.total + '</b></div>' +
      '<div><p class="ms-kick">' + esc(meta ? meta.title : '') + '</p><h2>' + (r.passed ? (r.score === r.total ? 'Perfect score. Proper.' : 'Passed. Nice work.') : 'Not yet – you need ' + r.passMark + '.') + '</h2><p class="ac-sub">' +
      (r.passed ? (nx ? 'Next up: ' + esc(nx.title) + '. It\'s unlocked now.' : 'That\'s every module done. Harvey can see your results.') : 'No stress. Go back over the sections below, check your notes, then try again – you\'ll get different questions.') + '</p>' +
      '<div class="ac-resb">' + (r.passed ? (nx ? '<button class="ms-btn acc" data-mod="' + nx.id + '">Start ' + esc(nx.title) + ' →</button>' : '<button class="ms-btn acc" data-go2="videos">Learn from the best →</button>') + '<button class="ms-btn ghost" data-go2="course">All modules</button>'
        : '<button class="ms-btn acc" data-mod="' + r.id + '">Back to the lesson</button><button class="ms-btn ghost" data-start="' + r.id + '">Try again – new questions</button>') + '</div></div></section>';
    var body = r.passed
      ? '<div class="ms-panel"><h3>Your answers</h3>' + r.review.map(function (x, k) { return '<div class="ac-rv ' + (x.ok ? 'ok' : 'no') + '"><b>' + (x.ok ? '✓' : '✕') + '</b><div><p>' + (k + 1) + '. ' + esc(x.q) + '</p>' + (x.ok ? '' : '<small>You said: ' + esc(x.your) + '</small>') + '<small class="r">Answer: ' + esc(x.right) + '</small><em>' + esc(x.why) + '</em></div></div>'; }).join('') + '</div>'
      : '<div class="ms-panel"><h3>Go back over these sections</h3><div class="ac-revise">' + r.revise.map(function (t) { return '<button class="ac-chip" data-revise="' + esc(t) + '">' + esc(t) + ' →</button>'; }).join('') + '</div>' +
        '<h3 style="margin-top:22px">The ones you missed</h3>' + r.review.filter(function (x) { return !x.ok; }).map(function (x) { return '<div class="ac-rv no"><b>✕</b><div><p>' + esc(x.q) + '</p><small>You said: ' + esc(x.your) + '</small><small class="r">Find the answer in: ' + esc(x.topic) + '</small></div></div>'; }).join('') + '</div>';
    return head + body;
  }
  function confetti() {
    var box = document.createElement('div'); box.className = 'ac-confetti'; var cols = ['#FF6A2E', '#FFB18F', '#3DDC97', '#FFFFFF', '#6FA8FF'];
    for (var k = 0; k < 70; k++) { var p = document.createElement('i'); p.style.left = Math.random() * 100 + '%'; p.style.background = cols[k % cols.length]; p.style.animationDelay = Math.random() * .6 + 's'; p.style.animationDuration = 1.8 + Math.random() * 1.6 + 's'; p.style.transform = 'rotate(' + Math.random() * 360 + 'deg)'; box.appendChild(p); }
    document.body.appendChild(box); setTimeout(function () { box.remove(); }, 4200);
  }

  /* ================= videos and library ================= */
  function videosHTML() {
    var v = S.ov.videos;
    return '<div class="ms-phead"><div><button class="ms-link back" data-go2="course">← All modules</button><p class="ms-kick dark">Learn from the best</p><h2>Watch how the pros do it.</h2><p class="ms-lead">Short videos from some of the best cold-calling trainers around. Most of them sell software in the US, but the skills are the same: a calm opener, real questions, listening, and handling objections without a fight. Tap one to watch, then tick it off.</p></div><div class="ac-vcount"><b>' + v.filter(function (x) { return x.watched; }).length + '/' + v.length + '</b><small>watched</small></div></div>' +
      '<div class="ac-vgrid">' + v.map(function (x) {
        return '<div class="ac-vid ms-panel"><div class="ac-vth" data-yt="' + esc(x.id) + '"><img src="https://i.ytimg.com/vi/' + esc(x.id) + '/hqdefault.jpg" alt="" loading="lazy"><span>▶</span></div><div class="ac-vb"><small>' + esc(x.who) + '</small><b>' + esc(x.title) + '</b><p>' + esc(x.why) + '</p>' +
          '<button class="ms-check' + (x.watched ? ' on' : '') + '" data-watch="' + esc(x.id) + '"' + (x.watched ? ' disabled' : '') + '><span>' + (x.watched ? '✓ Watched' : 'Mark as watched') + '</span></button></div></div>';
      }).join('') + '</div>';
  }
  function libraryHTML() {
    var m = S.opts.m || { training: [] }, items = m.training.filter(function (t) { return !t.locked; });
    return '<div class="ms-phead"><div><button class="ms-link back" data-go2="course">← All modules</button><p class="ms-kick dark">Reference library</p><h2>The source documents.</h2><p class="ms-lead">The full call script and decks, word for word. The Academy teaches you; these are for checking an exact line.</p></div></div>' +
      '<div class="ms-panel">' + (items.length ? items.map(function (t) { return '<div class="ac-lib"><div><b>' + esc(t.title) + '</b><small>' + esc(t.type || 'Doc') + '</small></div><button class="ms-btn ghost" data-doc="' + esc(t.title) + '">Open</button></div>'; }).join('') : '<p class="ms-p">Nothing here yet.</p>') + '</div>';
  }

  /* ================= events ================= */
  function bind() {
    var R = S.root;
    R.querySelectorAll('[data-mod]').forEach(function (b) { b.onclick = function () { go('module', b.dataset.mod); }; });
    R.querySelectorAll('[data-go2]').forEach(function (b) { b.onclick = function () { go(b.dataset.go2); }; });
    R.querySelectorAll('[data-start]').forEach(function (b) { b.onclick = function () { startTest(b.dataset.start); }; });
    R.querySelectorAll('[data-read]').forEach(function (b) { b.onclick = function () {
      var meta = S.ov.modules.filter(function (m) { return m.id === b.dataset.read; })[0];
      if (meta && !meta.lesson) { meta.lesson = true; api('academyMark', { id: meta.id, what: 'lesson' }); }
      b.className = 'ms-btn ghost'; b.textContent = '✓ Lesson read';
      var fl = R.querySelectorAll('.ac-fl')[1]; if (fl) fl.classList.add('on');
      var t = document.getElementById('ac-test'); if (t) t.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }; });
    R.querySelectorAll('.ac-flow a').forEach(function (a) { a.onclick = function (e) { e.preventDefault(); var t = document.querySelector(a.getAttribute('href')); if (t) t.scrollIntoView({ behavior: 'smooth', block: 'start' }); }; });
    R.querySelectorAll('[data-opt]').forEach(function (b) { b.onclick = function () {
      var t = S.test; t.ans[t.at] = +b.dataset.opt; render();
      if (t.at < t.qs.length - 1) setTimeout(function () { if (S.view === 'test' && S.test === t) { t.at++; render(); } }, 260);
    }; });
    R.querySelectorAll('[data-step]').forEach(function (b) { b.onclick = function () { S.test.at = Math.max(0, Math.min(S.test.qs.length - 1, S.test.at + +b.dataset.step)); render(); }; });
    R.querySelectorAll('[data-jump]').forEach(function (b) { b.onclick = function () { S.test.at = +b.dataset.jump; render(); }; });
    var sb = R.querySelector('[data-submit]'); if (sb) sb.onclick = submit;
    var qt = R.querySelector('[data-quit]'); if (qt) qt.onclick = function () { if (confirmLeave()) go('module', S.test.id); };
    R.querySelectorAll('[data-revise]').forEach(function (b) { b.onclick = function () {
      var title = b.dataset.revise, id = S.result.id; go('module', id);
      var tries = 0; (function find() { var hs = S.root.querySelectorAll('.ac-lesson h3'); for (var k = 0; k < hs.length; k++) if (hs[k].textContent === title) { hs[k].scrollIntoView({ behavior: 'smooth', block: 'start' }); hs[k].classList.add('ac-flash'); return; } if (tries++ < 30) setTimeout(find, 150); })();
    }; });
    R.querySelectorAll('[data-yt]').forEach(function (d) { d.onclick = function () { d.innerHTML = '<iframe src="https://www.youtube-nocookie.com/embed/' + encodeURIComponent(d.dataset.yt) + '?autoplay=1&rel=0" title="YouTube video" allow="autoplay; encrypted-media; picture-in-picture; fullscreen" allowfullscreen></iframe>'; d.onclick = null; d.classList.add('playing'); }; });
    R.querySelectorAll('[data-watch]').forEach(function (b) { b.onclick = function () { var v = S.ov.videos.filter(function (x) { return x.id === b.dataset.watch; })[0]; if (!v || v.watched) return; v.watched = true; api('academyMark', { id: v.id, what: 'video' }); render(); }; });
    R.querySelectorAll('[data-doc]').forEach(function (b) { b.onclick = function () { if (S.opts.openDoc) S.opts.openDoc(b.dataset.doc); }; });
  }
  function confirmLeave() { var n = S.test ? S.test.ans.filter(function (a) { return a !== null; }).length : 0; return !n || window.confirm('Leave this test? Your answers won\'t be saved, and next time you\'ll get new questions.'); }

  /* ================= home card ================= */
  function mountHome(el) {
    if (!el) return;
    var paint = function () {
      var mods = trackMods('rep'), done = mods.filter(function (m) { return m.passed; }).length, nx = nextMod();
      el.innerHTML = '<div class="ac-home"><div class="ac-hbar"><i style="width:' + Math.round(done / mods.length * 100) + '%"></i></div><p><b>' + done + ' of ' + mods.length + ' modules passed</b></p>' + (nx ? '<button class="ms-next" data-acnext>' + esc('Up next: ' + nx.title) + '</button>' : '<p class="ms-p">Every module passed. Nice work.</p>') + '</div>';
      var b = el.querySelector('[data-acnext]'); if (b && S.opts.openAcademy) b.onclick = function () { S.view = 'module'; S.mod = nx.id; S.opts.openAcademy(); };
    };
    if (S.ov) return paint();
    api('academy').then(function (r) { if (r.ok) { S.ov = r; paint(); } });
  }
  function setOpts(o) { S.opts = Object.assign(S.opts || {}, o || {}); }

  /* ================= admin ================= */
  function mountAdmin(el) {
    el.innerHTML = '<p class="note">Loading Academy progress…</p>';
    api('academyAdmin').then(function (r) {
      if (!r.ok) { el.innerHTML = '<p class="note">' + esc(r.error || 'Could not load.') + '</p>'; return; }
      var reps = r.reps.slice().sort(function (a, b) { return b.passed - a.passed || b.lastTs - a.lastTs; }), tot = r.modules.length;
      var week = Date.now() - 7 * 864e5, activeWk = reps.filter(function (u) { return u.lastTs > week; }).length, all = reps.filter(function (u) { return u.passed === tot; }).length;
      var riskHTML = '<div class="ac-risk"><div class="ac-riskh"><h3>Reps at risk</h3><span>' + (r.risk.length ? r.risk.length + ' to ring today' : 'Nobody right now') + '</span></div>' + (r.risk.length ? r.risk.map(function (x) { return '<div class="ac-riskr"><div><b>' + esc(x.name) + '</b><small>' + esc(x.email) + (x.phone ? ' · ' + esc(x.phone) : '') + (x.joined ? ' · joined ' + esc(x.joined) : '') + '</small></div><div class="ac-why">' + x.reasons.map(function (w) { return '<span>' + esc(w) + '</span>'; }).join('') + '</div></div>'; }).join('') + '<p class="note" style="margin:10px 0 0">Ring, don\'t text – and open with "How are you finding it?". See Leaders module 5.</p>' : '<p class="note" style="margin:0">Everyone active has logged in, logged calls and kept moving in the last 3 days.</p>') + '</div>';
      var leadHTML = (r.leaders && r.leaders.length && r.leadModules.length) ? '<h3 style="margin:26px 0 10px">Leaders track</h3><div class="ac-atw"><table class="ac-at" style="min-width:640px"><thead><tr><th>Leader</th><th>Passed</th>' + r.leadModules.map(function (m) { return '<th title="' + esc(m.title) + '">L' + m.n + '<small>' + esc(m.title) + '</small></th>'; }).join('') + '<th>Last active</th></tr></thead><tbody>' + r.leaders.map(function (u) { return '<tr><td><b>' + esc(u.name) + '</b><small>' + esc(u.role) + '</small></td><td><div class="ac-pbar"><i style="width:' + Math.round(u.passed / r.leadModules.length * 100) + '%"></i></div><small>' + u.passed + '/' + r.leadModules.length + '</small></td>' + u.mods.map(function (m) { return '<td class="ac-c ' + (m.passed ? 'ok' : m.attempts ? 'try' : '') + '">' + (m.attempts ? (m.passed ? '✓ ' : '') + m.best + '/10<small>' + m.attempts + ' tr' + (m.attempts === 1 ? 'y' : 'ies') + '</small>' : '<span class="ac-dash">–</span>') + '</td>'; }).join('') + '<td><small>' + esc(u.lastActive || 'Never') + '</small></td></tr>'; }).join('') + '</tbody></table></div>' : '';
      el.innerHTML = riskHTML + '<div class="tiles">' + [['Reps in the Academy', reps.filter(function (u) { return u.attempts || u.lastTs; }).length, 'of ' + reps.length + ' on the team'], ['Active this week', activeWk, 'did something in the last 7 days'], ['Finished every module', all, 'passed all ' + tot], ['Tests taken', reps.reduce(function (s, u) { return s + u.attempts; }, 0), 'pass mark ' + r.passMark + '/10']].map(function (t) { return '<div class="tile"><span>' + t[0] + '</span><b>' + t[1] + '</b><small>' + t[2] + '</small></div>'; }).join('') + '</div>' +
        '<p class="note" style="margin:16px 0 10px">Each cell shows the best score and number of attempts. ✓ = passed (' + r.passMark + '+). ▶ watched the video, ✎ read the lesson. Sorted by modules passed.</p>' +
        '<div class="ac-atw"><table class="ac-at"><thead><tr><th>Rep</th><th>Passed</th>' + r.modules.map(function (m) { return '<th title="' + esc(m.title) + '">M' + m.n + '<small>' + esc(m.title) + '</small></th>'; }).join('') + '<th>Videos</th><th>Last active</th></tr></thead><tbody>' +
        (reps.map(function (u) {
          return '<tr><td><b>' + esc(u.name) + '</b><small>' + esc(u.role) + (u.status && u.status !== 'Active' ? ' · ' + esc(u.status) : '') + '</small></td><td><div class="ac-pbar"><i style="width:' + Math.round(u.passed / tot * 100) + '%"></i></div><small>' + u.passed + '/' + tot + '</small></td>' +
            u.mods.map(function (m) { return '<td class="ac-c ' + (m.passed ? 'ok' : m.attempts ? 'try' : '') + '">' + (m.attempts ? (m.passed ? '✓ ' : '') + m.best + '/10<small>' + m.attempts + ' tr' + (m.attempts === 1 ? 'y' : 'ies') + '</small>' : '<span class="ac-dash">–</span>') + '<small class="ac-ic">' + (m.anim ? '▶' : '') + (m.lesson ? '✎' : '') + '</small></td>'; }).join('') +
            '<td>' + u.videos + '/' + r.videos + '</td><td><small>' + esc(u.lastActive || 'Never') + '</small></td></tr>';
        }).join('') || '<tr><td colspan="' + (tot + 4) + '">No reps yet.</td></tr>') + '</tbody></table></div>' +
        leadHTML + '<h3 style="margin:22px 0 10px">Latest activity</h3><div class="list">' + (r.recent.length ? r.recent.map(function (a) { return '<div class="item"><span><b>' + esc(a.name) + '</b><small>' + esc(a.type === 'Test' ? 'Took the ' + a.module + ' test: ' + a.score + '/10' + (a.passed ? ' – passed' : '') : a.type === 'Lesson' ? 'Read the ' + a.module + ' lesson' : a.type === 'Animation' ? 'Watched the ' + a.module + ' video' : 'Watched "' + a.module + '"') + '</small></span><small>' + esc(a.when) + '</small></div>'; }).join('') : '<p class="note">Nothing yet.</p>') + '</div>';
    });
  }

  window.MSAcademy = { mount: mount, mountHome: mountHome, mountAdmin: mountAdmin, setOpts: setOpts, reset: function () { S.ov = null; S.cache = {}; S.view = 'course'; } };
})();
