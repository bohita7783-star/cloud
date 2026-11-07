// v0.3.0: the in-app Excel looks like completed company reports. v0.4.0: works with a BLANK form
// (pass blank_template.xlsx / blank_approved.json). Real Chromium, SYNTHETIC data.
// Serve a COPY of homebox-app with the test key (see make_test_approval.py).
// Usage: node e2e_report_layout.js <baseUrl> <outDir> <template.xlsx> <approved.json> <photo.jpg> <photo2.jpg>
const { chromium } = require('playwright');
const [base, out, tplPath, manPath, photo, photo2] = process.argv.slice(2);
const step = m => console.log('✓', m);
const must = (c, m) => { if (!c) throw new Error(m); };

(async () => {
  const b = await chromium.launch({ env: { ...process.env, LANG: 'C.UTF-8', LC_ALL: 'C.UTF-8' } });
  const ctx = await b.newContext({ viewport: { width: 390, height: 844 }, acceptDownloads: true });
  const p = await ctx.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));

  // settings: approved template + reporter profile + usual contacts
  await p.goto(base + '#settings'); await p.waitForSelector('#tplReg');
  await p.setInputFiles('#tplX', tplPath); await p.setInputFiles('#tplM', manPath); await p.click('#tplReg');
  await p.waitForFunction(() => /登録しました|登録できません/.test(document.querySelector('#warmState').textContent));
  const tplText = await p.textContent('#tplState');
  must(tplText.includes('現場異常速報（2026年10月の配置）') && tplText.includes('写真枠 3つ'), 'form shown at registration: ' + tplText);
  step('registered template shows its form, slots and photo method: ' + tplText.split('承認 ')[0].replace(/\s+/g, ' ').slice(0, 90));
  await p.fill('#pf-dept', '架空支店'); await p.fill('#pf-bldg', '架空駅名店街'); await p.fill('#pf-name', '架空　太郎'); await p.selectOption('#pf-role', '班長');
  for (const [kind, who, org] of [['オーナー・テナント', '架空', '架空不動産株式会社'], ['オーナー・テナント', '仮名', '架空不動産株式会社'], ['社内', '架空課長', '']]) {
    await p.click('#pf-add'); const row = p.locator('#pf-ct [data-pi]').last();
    await row.locator('select[data-k="区分"]').selectOption(kind); await row.locator('input[data-k="相手"]').fill(who); await row.locator('input[data-k="所属"]').fill(org);
  }
  await p.click('#pf-save'); await p.waitForSelector('text=報告者情報を保存しました');
  step('reporter profile and usual contacts saved');

  // equipment, two photos, case
  await p.goto(base + '#eq-new'); await p.fill('#f-name', '架空 天井'); await p.click('button[type=submit]');
  await p.waitForFunction(() => /^#eq\/EQ-/.test(location.hash));
  const eq = decodeURIComponent((await p.evaluate(() => location.hash)).slice(4));
  await p.goto(base + '#shoot/' + eq); await p.waitForSelector('#cam', { state: 'attached' });
  await p.setInputFiles('#cam', photo); await p.waitForSelector('#shots .ph');
  await p.setInputFiles('#cam', photo2); await p.waitForFunction(() => document.querySelectorAll('#shots .ph').length === 2);
  await p.goto(base + '#eq/' + eq); await p.click('#mkcase'); await p.waitForSelector('#mkxlsx');
  const caseHash = await p.evaluate(() => location.hash);
  must((await p.locator('#ct [data-ci]').count()) === 3, 'usual contacts should be pre-filled into the new case');
  const f = l => p.locator('label.field').filter({ hasText: new RegExp('^' + l) }).locator('input,textarea,select').first();
  await f('件名').fill('区画12「架空店」ホール天井漏水あり');
  await f('内容').selectOption('漏水');
  await f('発生場所').fill('ホール');
  await f('当事者').first().fill('架空、仮名');
  await p.locator('label.field', { hasText: '当事者の区分' }).locator('select').selectOption('【当社の社員】');
  await f('現場確認を実施').check(); await f('現場確認開始').fill('10:35'); await f('現場確認終了').fill('10:40');
  await f('最初の状況').fill('架空店の店長より、店舗内テーブル席の天井より漏水していると連絡有り。');
  await f('今後の予定').fill('修理を提案予定。');
  await f('原因').fill('経年使用によるものと思われます。');
  await f('処置報告の内容').fill('バケツとプラベニにて漏水を受ける');
  await p.click('#tl-add'); await p.locator('#tl textarea').first().fill('天井内部を確認時、足場板を設置した上にバケツとプラベニを用いて漏水受けを設置する。お客様の少ない時間に再調査すると店長へ説明し退店する。');
  await p.locator('#pool .ph').nth(0).click(); await p.locator('#pool .ph').nth(1).click(); await p.waitForSelector('#captions textarea');
  await p.locator('#captions textarea').nth(0).fill('ホール天井\n天井より漏水あり'); await p.locator('#captions textarea').nth(1).fill('天井内にバケツ設置');
  await p.waitForTimeout(900);
  step('case filled with the form fields; usual contacts came in automatically');

  // preview with the same engine
  await p.click('#pv'); await p.waitForSelector('#pvOut table, #pvOut .problems', { timeout: 60000 });
  const pv = await p.textContent('#pvOut');
  must(pv.includes('様式に収まります') && pv.includes('修理を提案予定。') && pv.includes('（第１報）'), 'preview: ' + pv.slice(0, 200));
  step('「様式に入るか確認」 shows the rows as they will be printed');

  // too many owner contacts -> the check says so before making the Excel
  for (let i = 0; i < 3; i++) await p.click('#ct-add');
  await p.waitForTimeout(400);
  must((await p.textContent('#probs')).includes('4件まで'), 'contact limit should be flagged');
  for (let i = 0; i < 3; i++) await p.locator('#ct [data-crm]').last().click();
  await p.waitForTimeout(400);
  step('the limit of 4 owner/tenant contacts is checked before export');

  // make the Excel and inspect it
  const [dl] = await Promise.all([p.waitForEvent('download', { timeout: 120000 }), p.click('#mkxlsx')]);
  const file = out + '/' + dl.suggestedFilename(); await dl.saveAs(file);
  step('Excel made: ' + dl.suggestedFilename());

  // follow-up report
  await p.click('#follow'); await p.waitForFunction(h => location.hash !== h, caseHash);
  must((await p.locator('label.field', { hasText: '報告の段階' }).locator('select').inputValue()) === '（第２報）', 'follow-up should be 第２報');
  must((await p.inputValue('label.field:has-text("件名") input')) === '区画12「架空店」ホール天井漏水あり', 'follow-up keeps the content');
  step('「続報を作る」 makes 第２報 with the same content');
  console.log(JSON.stringify({ file, errors }));
  await b.close();
})().catch(e => { console.error('FAILED:', e.message); process.exit(1); });
