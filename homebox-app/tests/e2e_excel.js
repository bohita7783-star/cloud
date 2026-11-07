// In-app Excel generation, end to end, in real Chromium with SYNTHETIC data.
// Serve a COPY of homebox-app whose trusted_approvers.json holds the test key.
// Usage: node e2e_excel.js <baseUrl> <outDir> <template.xlsx> <approved.json> <tampered.json> <photo.jpg>
const { chromium } = require('playwright');
const [base, out, tplPath, manPath, badManPath, photo] = process.argv.slice(2);
const step = m => console.log('✓', m);
const must = (c, m) => { if (!c) throw new Error(m); };
const secs = t0 => ((Date.now() - t0) / 1000).toFixed(1);

(async () => {
  // UTF-8 locale: without it, Chromium on Linux saves Japanese file names as "download".
  const b = await chromium.launch({ env: { ...process.env, LANG: 'C.UTF-8', LC_ALL: 'C.UTF-8' } });
  const ctx = await b.newContext({ viewport: { width: 390, height: 844 }, acceptDownloads: true });
  const p = await ctx.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  // first visit over a slow mobile-like link (about 10 Mbit/s, 60 ms)
  const cdp = await ctx.newCDPSession(p);
  await cdp.send('Network.emulateNetworkConditions', { offline: false, latency: 60, downloadThroughput: 10e6 / 8, uploadThroughput: 3e6 / 8 });
  await p.goto(base + '#settings'); await p.waitForSelector('#tplReg');

  // a tampered approval is refused
  await p.setInputFiles('#tplX', tplPath); await p.setInputFiles('#tplM', badManPath);
  let t0 = Date.now(); await p.click('#tplReg');
  await p.waitForFunction(() => /登録できません|登録しました/.test(document.querySelector('#warmState').textContent), null, { timeout: 120000 });
  const firstLoad = secs(t0);
  must((await p.textContent('#warmState')).includes('署名が一致しません') && !(await p.textContent('#warmState')).includes('ApprovalError'), 'tampered manifest should be refused: ' + await p.textContent('#warmState'));
  step(`tampered approval refused (first engine start incl. 19 MB download at ~10 Mbit/s: ${firstLoad}s)`);

  await p.setInputFiles('#tplM', manPath); t0 = Date.now(); await p.click('#tplReg');
  await p.waitForFunction(() => /照合して|照合しています/.test(document.querySelector('#warmState').textContent) || /登録しました/.test(document.querySelector('#warmState').textContent));
  await p.waitForFunction(() => /登録しました|登録できません/.test(document.querySelector('#warmState').textContent));
  must((await p.textContent('#warmState')).includes('登録しました'), await p.textContent('#warmState'));
  must((await p.textContent('#tplState')).includes('APV-'), 'approval id not shown');
  step(`template registered after signature check (${secs(t0)}s)`);
  await cdp.send('Network.emulateNetworkConditions', { offline: false, latency: 0, downloadThroughput: -1, uploadThroughput: -1 });

  // minimal case: equipment, one photo, case with caption
  await p.goto(base + '#eq-new'); await p.fill('#f-name', '架空 熱源水ポンプ'); await p.click('button[type=submit]');
  await p.waitForFunction(() => /^#eq\/EQ-/.test(location.hash));
  const eqId = decodeURIComponent((await p.evaluate(() => location.hash)).slice(4));
  await p.goto(base + '#shoot/' + eqId); await p.waitForSelector('#cam', { state: 'attached' });
  await p.setInputFiles('#cam', photo); await p.waitForSelector('#shots .ph');
  await p.goto(base + '#eq/' + eqId); await p.click('#mkcase'); await p.waitForSelector('#mkxlsx');
  const caseHash = await p.evaluate(() => location.hash);
  const field = l => p.locator('label.field', { hasText: l }).locator('input,textarea,select').first();
  await field('件名').fill('架空 熱源水ポンプ 異音'); await field('状況概要').fill('架空：軸受付近から異音。');
  await field('発生場所').fill('架空ビル B1 機械室');
  await p.locator('#pool .ph').first().click(); await p.waitForSelector('#captions textarea');
  await p.locator('#captions textarea').first().fill('架空：銘板を確認。');
  await p.waitForTimeout(900);

  // make the Excel on the device
  t0 = Date.now();
  const [dl] = await Promise.all([p.waitForEvent('download', { timeout: 120000 }), p.click('#mkxlsx')]);
  const name = dl.suggestedFilename(); await dl.saveAs(out + '/' + name);
  await p.waitForSelector('#xlsxResult b');
  must((await p.textContent('#xlsxResult b')).trim() === name, 'shown name differs from saved name');
  must(/会社書式_要確認_\d{12}_[0-9a-f]{8}\.xlsx$/.test(name), 'unexpected name ' + name);
  await p.waitForSelector('#dla');
  const [audit] = await Promise.all([p.waitForEvent('download'), p.click('#dla')]); await audit.saveAs(out + '/in_app_audit.json');
  step(`Excel made on the device: ${name} (${secs(t0)}s)`);

  // airplane mode: reload from the service worker cache and make it again
  await ctx.setOffline(true);
  await p.reload(); await p.waitForSelector('#mkxlsx');
  t0 = Date.now();
  const [dl2] = await Promise.all([p.waitForEvent('download', { timeout: 120000 }), p.click('#mkxlsx')]);
  await dl2.saveAs(out + '/offline_' + dl2.suggestedFilename());
  step(`offline: app reloaded from cache and made the Excel again (${secs(t0)}s incl. engine restart)`);
  await ctx.setOffline(false);

  // no template -> clear message, nothing produced
  await p.goto(base + '#settings'); await p.click('#tplDel'); await p.waitForSelector('dialog[open]'); await p.click('dialog button[value=yes]');
  await p.waitForFunction(() => document.querySelector('#tplState').textContent.includes('未登録'));
  await p.goto(base + caseHash); await p.waitForSelector('#mkxlsx'); await p.waitForTimeout(500);
  await p.click('#mkxlsx'); await p.waitForSelector('dialog[open]');
  must((await p.textContent('dialog')).includes('会社様式が登録されていません'), 'missing-template message');
  step('without a registered template the app explains what to do');
  console.log(JSON.stringify({ excel: name, errors }));
  await b.close();
})().catch(e => { console.error('FAILED:', e.message); process.exit(1); });
