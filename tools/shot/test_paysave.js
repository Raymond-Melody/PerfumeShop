const { chromium } = require('playwright-core');
const U='Ray88',P='Ray123456';
(async()=>{const b=await chromium.launch({channel:'msedge',headless:true});try{
const c=await b.newContext();const pg=await c.newPage();await pg.setViewportSize({width:1440,height:900});
await pg.goto('http://localhost:5207/login',{waitUntil:'domcontentloaded'});await pg.waitForTimeout(600);
await pg.locator('input[name="username"]').first().fill(U);await pg.locator('input[name="password"]').first().fill(P);
await Promise.all([pg.waitForLoadState('domcontentloaded').catch(()=>{}),pg.click('button[type=submit]')]);await pg.waitForTimeout(1500);
await pg.goto('http://localhost:5207/admin/finance/payment-config',{waitUntil:'domcontentloaded'});
await pg.waitForSelector('.btn-save',{timeout:20000});await pg.waitForFunction(()=>window.Blazor!==undefined,{timeout:15000}).catch(()=>{});await pg.waitForTimeout(5000);
await pg.click('.btn-save');await pg.waitForTimeout(2500);
const msg=await pg.locator('.alert').first().innerText().catch(()=>'(no alert)');
console.log('[paycfg] save result:',msg.replace(/\s+/g,' ').trim());
await c.close();}catch(e){console.error('FATAL',e.message);process.exitCode=1;}finally{await b.close();}})();
