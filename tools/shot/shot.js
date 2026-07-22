// Fund Dashboard visual capture: V18 (Classic ASP :80) vs V19 (Blazor :5207)
// Uses playwright-core with installed Edge/Chrome (no browser download).
const { chromium } = require('playwright-core');
const fs = require('fs');

const OUT = 'f:/网站制作/网站/网站二/docs/screenshots/compare';
const USER = 'Ray88', PASS = 'Ray123456';
const BPS = [
  { name: 'desktop', w: 1440, h: 900 },
  { name: 'tablet',  w: 900,  h: 1024 },
  { name: 'mobile',  w: 390,  h: 844 },
];
const log = (...a) => console.log('[shot]', ...a);
if (!fs.existsSync(OUT)) fs.mkdirSync(OUT, { recursive: true });

async function launch() {
  const attempts = [
    { channel: 'msedge', headless: true },
    { channel: 'chrome', headless: true },
    { executablePath: (process.env.LOCALAPPDATA || '') + '\\ms-playwright\\chromium-1223\\chrome-win\\chrome.exe', headless: true },
  ];
  for (const opt of attempts) {
    try { const b = await chromium.launch(opt); log('browser launched via', opt.channel || 'cached chromium'); return b; }
    catch (e) { log('launch failed:', (opt.channel||'cached'), '-', e.message); }
  }
  throw new Error('Could not launch any browser');
}

async function grab(page, sel) {
  try { const t = await page.locator(sel).first().innerText({ timeout: 2000 }); return t.replace(/\s+/g,' ').trim(); }
  catch { return '(missing)'; }
}

async function v18Login(page) {
  for (let i = 1; i <= 3; i++) {
    await page.goto('http://localhost/admin/login.asp', { waitUntil: 'domcontentloaded' });
    await page.fill('#username', USER);
    await page.fill('#password', PASS);
    await Promise.all([
      page.waitForLoadState('domcontentloaded').catch(()=>{}),
      page.click('button[type=submit]'),
    ]);
    await page.waitForTimeout(1000);
    const url = page.url();
    const body = await page.textContent('body').catch(()=> '');
    if (/login\.asp/i.test(url) && /(安全验证失败|请刷新)/.test(body)) { log('V18 CSRF fail, retry', i); continue; }
    if (/用户名或密码错误|账户已.*锁定|过于频繁/.test(body) && /login\.asp/i.test(url)) throw new Error('V18 login rejected: ' + body.slice(0,60));
    log('V18 login OK ->', url);
    return true;
  }
  throw new Error('V18 login failed after retries');
}

async function v19Login(page) {
  await page.goto('http://localhost:5207/login', { waitUntil: 'domcontentloaded' });
  await page.waitForTimeout(600);
  const u = page.locator('input[name="username"]');
  if (await u.count()) await u.first().fill(USER);
  else await page.locator('form input[type="text"], form input:not([type])').first().fill(USER);
  const p = page.locator('input[name="password"]');
  if (await p.count()) await p.first().fill(PASS);
  else await page.locator('input[type="password"]').first().fill(PASS);
  await Promise.all([
    page.waitForLoadState('domcontentloaded').catch(()=>{}),
    page.click('button[type=submit]'),
  ]);
  await page.waitForTimeout(1500);
  log('V19 after login ->', page.url());
  return true;
}

async function shoot(page, url, prefix, waitSel, extraWait, dumpSels) {
  const data = {};
  for (const bp of BPS) {
    await page.setViewportSize({ width: bp.w, height: bp.h });
    await page.goto(url, { waitUntil: 'domcontentloaded' });
    if (waitSel) { try { await page.waitForSelector(waitSel, { timeout: 20000 }); } catch { log(prefix, 'waitSel timeout:', waitSel); } }
    if (extraWait) await page.waitForTimeout(extraWait);
    const file = `${OUT}/${prefix}-${bp.name}.png`;
    await page.screenshot({ path: file, fullPage: true });
    log('saved', file);
    if (bp.name === 'desktop' && dumpSels) {
      for (const [k, sel] of Object.entries(dumpSels)) data[k] = await grab(page, sel);
    }
  }
  return data;
}

(async () => {
  const browser = await launch();
  const results = {};
  try {
    // ---- V18 ----
    const ctx18 = await browser.newContext({ ignoreHTTPSErrors: true });
    const p18 = await ctx18.newPage();
    await v18Login(p18);
    results.v18 = await shoot(
      p18, 'http://localhost/admin/finance/fund_dashboard.asp', 'v18-fund',
      '.stats-overview .stat-value', 800,
      {
        statsOverview: '.stats-overview',
        accountsGrid: '.accounts-grid',
        pendingSection: '.pending-section',
        alertSection: '.alert-section',
      }
    );
    results.v18.abnormal = await grab(p18, '.abnormal-alert');
    await ctx18.close();

    // ---- V19 ----
    const ctx19 = await browser.newContext({ ignoreHTTPSErrors: true });
    const p19 = await ctx19.newPage();
    await v19Login(p19);
    results.v19 = await shoot(
      p19, 'http://localhost:5207/admin/finance/fund-dashboard', 'v19-fund',
      '.stats-row', 3000,
      {
        statsRow: '.stats-row',
        table: '.table-container',
      }
    );
    // chart canvas presence
    try {
      const hasChart = await p19.locator('#fundFlowChart').count();
      results.v19.chartCanvas = hasChart ? 'present' : 'absent';
    } catch { results.v19.chartCanvas = 'error'; }
    await ctx19.close();

    console.log('\n===== EXTRACTED DATA =====');
    console.log(JSON.stringify(results, null, 2));
    console.log('===== DONE =====');
  } catch (e) {
    console.error('[shot] FATAL:', e.message);
    process.exitCode = 1;
  } finally {
    await browser.close();
  }
})();
