// Smoke test: login V19 and capture purchase-review + payment-config pages.
const { chromium } = require('playwright-core');
const OUT = 'f:/网站制作/网站/网站二/docs/screenshots/compare';
const USER = 'Ray88', PASS = 'Ray123456';
const log = (...a) => console.log('[smoke]', ...a);

async function launch() {
  for (const opt of [{channel:'msedge',headless:true},{channel:'chrome',headless:true}]) {
    try { return await chromium.launch(opt); } catch(e){ log('launch fail', e.message); }
  }
  throw new Error('no browser');
}

async function grab(page, sel) {
  try { return (await page.locator(sel).first().innerText({ timeout: 2500 })).replace(/\s+/g,' ').trim().slice(0,400); }
  catch { return '(missing)'; }
}

(async () => {
  const browser = await launch();
  try {
    const ctx = await browser.newContext();
    const page = await ctx.newPage();
    await page.setViewportSize({ width: 1440, height: 900 });
    await page.goto('http://localhost:5207/login', { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(600);
    const u = page.locator('input[name="username"]'); if (await u.count()) await u.first().fill(USER); else await page.locator('form input[type="text"]').first().fill(USER);
    const p = page.locator('input[name="password"]'); if (await p.count()) await p.first().fill(PASS); else await page.locator('input[type="password"]').first().fill(PASS);
    await Promise.all([page.waitForLoadState('domcontentloaded').catch(()=>{}), page.click('button[type=submit]')]);
    await page.waitForTimeout(1500);
    log('after login ->', page.url());

    // Purchase Review
    await page.goto('http://localhost:5207/admin/finance/purchase-review', { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(2500);
    await page.screenshot({ path: `${OUT}/v19-purchase-review.png`, fullPage: true });
    log('purchase-review saved');
    log('  stats:', await grab(page, '.stats-row'));
    log('  tabs:', await grab(page, '.tab-nav'));
    log('  denied?:', await grab(page, '.alert-error'));

    // Payment Config
    await page.goto('http://localhost:5207/admin/finance/payment-config', { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(2500);
    await page.screenshot({ path: `${OUT}/v19-payment-config.png`, fullPage: true });
    log('payment-config saved');
    log('  channels:', await grab(page, '.config-container'));

    await ctx.close();
  } catch (e) { console.error('[smoke] FATAL:', e.message); process.exitCode = 1; }
  finally { await browser.close(); }
})();
