// Capture V19 fund dashboard after toggling the header theme button (dark -> light).
const { chromium } = require('playwright-core');
const OUT = 'f:/网站制作/网站/网站二/docs/screenshots/compare';
const USER = 'Ray88', PASS = 'Ray123456';
const log = (...a) => console.log('[theme]', ...a);

async function launch() {
  for (const opt of [{channel:'msedge',headless:true},{channel:'chrome',headless:true},{executablePath:(process.env.LOCALAPPDATA||'')+'\\ms-playwright\\chromium-1223\\chrome-win\\chrome.exe',headless:true}]) {
    try { return await chromium.launch(opt); } catch(e){ log('launch fail', e.message); }
  }
  throw new Error('no browser');
}

(async () => {
  const browser = await launch();
  try {
    const ctx = await browser.newContext();
    const page = await ctx.newPage();
    await page.setViewportSize({ width: 1440, height: 900 });
    // login
    await page.goto('http://localhost:5207/login', { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(600);
    const u = page.locator('input[name="username"]'); if (await u.count()) await u.first().fill(USER); else await page.locator('form input[type="text"]').first().fill(USER);
    const p = page.locator('input[name="password"]'); if (await p.count()) await p.first().fill(PASS); else await page.locator('input[type="password"]').first().fill(PASS);
    await Promise.all([page.waitForLoadState('domcontentloaded').catch(()=>{}), page.click('button[type=submit]')]);
    await page.waitForTimeout(1500);
    // dashboard
    await page.goto('http://localhost:5207/admin/finance/fund-dashboard', { waitUntil: 'domcontentloaded' });
    await page.waitForSelector('.stats-row', { timeout: 20000 }).catch(()=>{});
    await page.waitForTimeout(2500);
    // click theme toggle (last icon button in app bar)
    const toggle = page.locator('.mud-appbar button.mud-icon-button, header button.mud-icon-button').last();
    const n = await toggle.count();
    log('toggle candidates:', n);
    if (n) { await toggle.click(); log('clicked theme toggle'); }
    await page.waitForTimeout(1500);
    await page.screenshot({ path: `${OUT}/v19-fund-desktop-light.png`, fullPage: true });
    log('saved v19-fund-desktop-light.png');
    await ctx.close();
  } catch (e) { console.error('[theme] FATAL:', e.message); process.exitCode = 1; }
  finally { await browser.close(); }
})();
