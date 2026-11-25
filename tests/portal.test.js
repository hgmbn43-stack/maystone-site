// Portal page checks (25 Nov 2026). Runs the real login/index.html in a headless browser with a fake backend,
// so nothing touches the live sheet. Run: node tests/portal.test.js  (needs: npm i playwright && npx playwright install chromium)
// Fails (exit 1) if booking, the ready link, "not interested" or the admin pages break, or the page throws an error.
const { chromium, devices } = require('playwright');
const http = require('http'), fs = require('fs'), path = require('path');
const ROOT = path.join(__dirname, '..'), PORT = 8771;
const LINK = 'https://hgmbn43-stack.github.io/maystone-previews/p/T0001-test-plumbing/';
const srv = http.createServer((q, s) => {
  let f = path.join(ROOT, decodeURIComponent(q.url.split('?')[0]));
  if (fs.existsSync(f) && fs.statSync(f).isDirectory()) f = path.join(f, 'index.html');
  if (!fs.existsSync(f)) { s.writeHead(404); return s.end(); }
  const ext = path.extname(f); s.writeHead(200, { 'Content-Type': { '.js': 'text/javascript', '.css': 'text/css', '.svg': 'image/svg+xml', '.png': 'image/png' }[ext] || 'text/html' });
  s.end(fs.readFileSync(f));
}).listen(PORT);
const fails = [];
const check = (ok, what) => { console.log((ok ? 'PASS ' : 'FAIL ') + what); if (!ok) fails.push(what); };
const lead = () => ({ _row: 2, 'Lead ID': 'T0001', Business: 'Test Plumbing', Trade: 'Plumber', Town: 'Hove', Phone: '07700 900000', Priority: 'A', 'Google reviews': 5, 'Google rating': 5 });

async function session(b, admin, fn) {
  const ctx = await b.newContext(devices['iPhone 13']), errs = [], calls = [];
  await ctx.addInitScript(() => { try { localStorage.setItem('maystone.session', 'test-session-not-real'); } catch (e) {} });
  let L = lead(), seen = 0;
  await ctx.route('**/*', async r => {
    const u = r.request().url();
    if (u.startsWith('http://localhost:' + PORT)) return r.continue();
    if (!u.includes('script.google.com')) return r.abort();
    const q = JSON.parse(r.request().postData() || '{}'); calls.push(q); let out = { ok: false, error: 'stub' };
    if (q.action === 'me') out = { ok: true, admin, manager: false, hasPassword: true, user: { name: admin ? 'Harvey Matthews' : 'Test Rep', email: 'test@example.com', role: admin ? 'Admin' : 'Rep' },
      mine: { previews: 0, sales: 0, pending: 0, cleared: 0, paidOut: 0 }, training: [], events: [], dash: { week: [{ d: 'Mon', n: 0 }], feed: [], previewsToday: 0, callsToday: 0 }, leaderboard: [], rank: 1 };
    if (q.action === 'leads') out = { ok: true, tab: 'Test', tabs: [], head: Object.keys(L), editable: [], leads: [L] };
    if (q.action === 'bookNow') { L = Object.assign({}, L, { Outcome: 'Yes – send preview', 'Called (tick)': 'Answered', 'Preview link': LINK }); out = { ok: true, lead: L }; }
    if (q.action === 'saveLead') { L = Object.assign({}, L, q.fields); out = { ok: true, lead: L }; }
    if (q.action === 'leadSeen') { seen++; out = seen < 2 ? { ok: true, views: 0, lastViewed: 0 } : seen < 3 ? { ok: true, views: 1, lastViewed: Date.now() - 3000 } : { ok: true, views: 1, lastViewed: Date.now() - 3000, liked: Date.now() - 1000 }; }
    if (q.action === 'adminPreviews') out = { ok: true, previews: [] };
    if (q.action === 'closerList') out = { ok: true, now: Date.now(), me: 'Harvey', counts: {}, labels: {}, rows: [] };
    r.fulfill({ contentType: 'application/json', body: JSON.stringify(out) });
  });
  const p = await ctx.newPage(); p.on('pageerror', e => errs.push(e.message));
  try { await fn(p, calls); } catch (e) { check(false, 'flow ran without getting stuck: ' + e.message.split('\n')[0]); }
  check(!errs.length, (admin ? 'admin' : 'rep') + ' page has no script errors' + (errs.length ? ': ' + errs[0] : ''));
  await ctx.close();
}
async function openLead(p) {
  await p.goto('http://localhost:' + PORT + '/login/'); await p.waitForTimeout(2000);
  await p.evaluate(() => { const b = [...document.querySelectorAll('button,a')].find(x => /^\s*(My )?leads\s*$/i.test(x.textContent)); if (b) b.click(); });
  await p.waitForTimeout(1200); await p.locator('text=Test Plumbing').first().click({ timeout: 8000 }); await p.waitForTimeout(500);
}

(async () => {
  const b = await chromium.launch();
  await session(b, false, async (p, calls) => {          // rep books a preview in one tap
    await openLead(p);
    await p.locator('[data-qa=book]').click(); await p.waitForTimeout(1500);
    check(await p.locator('.ms-bookbox:not([data-nibox])').isVisible(), 'Booked a preview opens the booking box (not the not-interested box)');
    check(calls.some(c => c.action === 'bookNow'), 'one tap books it (bookNow)');
    check(await p.locator('.pv-ready').isVisible(), 'ready link box shows straight away');
    const sms = await p.getAttribute('.pv-ready a[href^="sms:"]', 'href');
    check(!!sms && decodeURIComponent(sms).includes('hi its test from maystone') && decodeURIComponent(sms).includes(LINK), 'Text them button has the right text and link');
    await p.waitForTimeout(6500);
    check(/opened it/i.test(await p.locator('[data-seen]').innerText()), '"They\'ve opened it" shows when they open the preview');
    await p.waitForTimeout(5500);
    check(/I like it/.test(await p.locator('[data-seen]').innerText()), '"They tapped I like it" shows when they tap the button');
    await p.locator('[data-otc="Opened it on the call"]').click();
    await p.locator('.ms-drawer [type=submit]').nth(1).click(); await p.waitForTimeout(700);
    check(await p.locator('.ms-drawer .ms-err').isVisible(), 'Save asks for their email first');
    await p.fill('input[name="Client email"]', 'customer@example.com'); await p.locator('[data-agreed=No]').click();
    await p.locator('.ms-drawer [type=submit]').nth(1).click(); await p.waitForTimeout(1200);
    const s = calls.filter(c => c.action === 'saveLead').pop();
    check(!!s && s.fields['On the call'] === 'Opened it on the call' && s.fields['Client email'] === 'customer@example.com' && s.fields['CB agreed'] === 'No', 'Save sends email, call-back answer and On the call');
  });
  await session(b, false, async (p, calls) => {          // rep marks not interested with a reason
    await openLead(p);
    await p.locator('[data-qa=ni]').click(); await p.waitForTimeout(300);
    await p.locator('[data-ni="Gets calls like this all the time"]').click(); await p.waitForTimeout(1000);
    const s = calls.filter(c => c.action === 'saveLead').pop();
    check(!!s && /not interested/i.test(s.fields.Outcome) && s.fields['No reason'] === 'Gets calls like this all the time', 'Not interested saves with the reason');
  });
  await session(b, true, async (p) => {                  // admin pages load
    await p.goto('http://localhost:' + PORT + '/login/#closer'); await p.waitForTimeout(3000);
    check(!/Loading your calls/.test(await p.locator('body').innerText()), 'Closer page loads');
  });
  await b.close(); srv.close();
  console.log(fails.length ? '\n' + fails.length + ' check(s) failed' : '\nAll checks passed');
  process.exit(fails.length ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
