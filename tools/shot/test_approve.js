// Interaction test: approve a pending purchase order via the modal (write-path verification).
const { chromium } = require('playwright-core');
const OUT = 'f:/网站制作/网站/网站二/docs/screenshots/compare';
const USER = 'Ray88', PASS = 'Ray123456';
const log = (...a) => console.log('[approve]', ...a);

async function launch() {
  for (const opt of [{channel:'msedge',headless:true},{channel:'chrome',headless:true}]) {
    try { return await chromium.launch(opt); } catch(e){}
  }
  throw new Error('no browser');
}

(async () => {
  const browser = await launch();
  try {
    const ctx = await browser.newContext();
    const page = await ctx.newPage();
    await page.setViewportSize({ width: 1440, height: 900 });
    await page.goto('http://localhost:5207/login', { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(600);
    await page.locator('input[name="username"]').first().fill(USER);
    await page.locator('input[name="password"]').first().fill(PASS);
    await Promise.all([page.waitForLoadState('domcontentloaded').catch(()=>{}), page.click('button[type=submit]')]);
    await page.waitForTimeout(1500);

    await page.goto('http://localhost:5207/admin/finance/purchase-review', { waitUntil: 'domcontentloaded' });
    // wait for interactive circuit + data
    await page.waitForSelector('.btn-approve', { timeout: 20000 });
    // Blazor Server: wait for SignalR circuit to be live before clicking
    await page.waitForFunction(() => window.Blazor !== undefined, { timeout: 15000 }).catch(()=>{});
    await page.waitForTimeout(5000);
    log('clicking 通过...');
    await page.click('.btn-approve');
    // if modal did not open (circuit not ready), retry once
    try { await page.waitForSelector('.modal-overlay', { timeout: 6000 }); }
    catch { log('modal not open, retry click'); await page.waitForTimeout(3000); await page.click('.btn-approve'); await page.waitForSelector('.modal-overlay', { timeout: 8000 }); }
    await page.waitForTimeout(600);
    // select cost allocation
    await page.selectOption('.modal-content select', 'RawMaterial');
    await page.waitForTimeout(400);
    await page.click('.modal-content .btn-save');
    await page.waitForTimeout(2500);
    const msg = await page.locator('.alert').first().innerText().catch(()=>'(no alert)');
    log('result msg:', msg.replace(/\s+/g,' ').trim());
    const stats = await page.locator('.stats-row').first().innerText().catch(()=>'');
    log('stats after:', stats.replace(/\s+/g,' ').trim());
    await page.screenshot({ path: `${OUT}/v19-purchase-review-approved.png`, fullPage: true });
    log('screenshot saved');
    await ctx.close();
  } catch (e) { console.error('[approve] FATAL:', e.message); process.exitCode = 1; }
  finally { await browser.close(); }
})();
