// End-to-end: drive the HomeBox app in Chromium with SYNTHETIC data and save the
// report package + backup it produces.
// Usage: node e2e_app.js <baseUrl> <outDir> <jpeg1> <jpeg2>   (two DIFFERENT synthetic JPEGs;
//        jpeg1 should carry EXIF DateTimeOriginal 2026:11:01 09:15:00 +09:00)
const { chromium } = require('playwright');
const fs = require('fs');
const [base, out, jpeg, jpeg2] = process.argv.slice(2);
const step = m => console.log('✓', m);
const must = (cond, msg) => { if (!cond) throw new Error(msg); };

(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 390, height: 844 }, acceptDownloads: true, colorScheme: 'light' });
  const p = await ctx.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  p.on('console', m => m.type() === 'error' && errors.push(m.text()));
  await p.goto(base + '#home');
  await p.waitForSelector('#q');
  await p.screenshot({ path: out + '/shot-home-empty.png' });
  step('home renders');

  // CSV import (synthetic ledger, one row keeps a company ID)
  const csv = '﻿設備ID,設備名,分類,場所の区分,QR,製造者名,型式,機器番号,電動機出力,定格電流,揚程\r\n' +
    'EQ-000006,架空 熱源水ポンプ,ポンプ,架空ビル / B1機械室,,架空機工,HB-65P,P-6,7.5,15.6,32\r\n' +
    ',架空 冷却塔ファン,ファン・送風機,架空ビル / 屋上,,架空空調,CT-220F,F-11,3.7,8.2,\r\n';
  fs.writeFileSync(out + '/ledger.csv', csv);
  await p.goto(base + '#settings'); await p.waitForSelector('#csvin', { state: 'attached' });
  await p.setInputFiles('#csvin', out + '/ledger.csv');
  await p.waitForSelector('dialog[open]'); await p.click('dialog button[value=yes]');
  await p.waitForSelector('a.item[href="#eq/EQ-000006"]');
  const ids = await p.$$eval('a.item', as => as.map(a => a.getAttribute('href')));
  must(ids.includes('#eq/EQ-000006') && ids.includes('#eq/EQ-000007'), 'CSV ids wrong: ' + ids);
  step('CSV import keeps EQ-000006 and numbers the next one EQ-000007');

  // Ledger: mark 製造者名 confirmed with a source
  await p.goto(base + '#eq/EQ-000006'); await p.waitForSelector('#saveL');
  must((await p.inputValue('#lv-0')) === '架空機工', 'CSV value not imported');
  await p.selectOption('#ls-0', 'ok'); await p.fill('#lc-0', '銘板写真');
  await p.click('#saveL'); await p.waitForSelector('text=確認済 1/17');
  must(await p.locator('text=追加項目').count() > 0, 'extension section missing');
  must((await p.inputValue('#lv-' + 17)) === 'HB-65P', 'model should be the first extension item and keep its value');
  step('ledger saved with status and source');

  // Inspection record
  await p.click('#mkrec'); await p.waitForSelector('dialog[open]');
  const reads = await p.$$('#r-reads input'); await reads[0].fill('運転電流'); await reads[1].fill('15.2'); await reads[2].fill('A');
  await p.click('dialog button[value=yes]'); await p.waitForSelector('text=運転電流：15.2 A');
  step('inspection recorded');

  // Shoot 2 photos for the equipment
  await p.goto(base + '#shoot/EQ-000006'); await p.waitForSelector('#cam', { state: 'attached' });
  must((await p.textContent('#preview')).includes('EQ-000006_連番_銘板'), 'preview wrong');
  await p.setInputFiles('#cam', jpeg); await p.waitForSelector('#shots .ph');
  await p.click('label[for="k-2"]');
  await p.setInputFiles('#cam', jpeg2); await p.waitForFunction(() => document.querySelectorAll('#shots .ph').length === 2);
  const names = await p.$$eval('#shots .n', n => n.map(x => x.textContent));
  must(names.some(n => /_EQ-000006_001_銘板\.jpg$/.test(n)) && names.some(n => /_EQ-000006_002_圧力計\.jpg$/.test(n)), 'names: ' + names);
  step('photos saved with rule names: ' + names.join(' , '));
  const src = await p.evaluate(async () => { const all = await new Promise(r => { const q = indexedDB.open('homebox'); q.onsuccess = () => { const t = q.result.transaction('photos').objectStore('photos').getAll(); t.onsuccess = () => r(t.result); }; }); return all.map(x => [x.kind, x.takenSource, x.takenAt]); });
  must(src.some(([k, s, t]) => k === '銘板' && s === 'EXIF' && t.startsWith('2026-11-01T09:15:00+09:00')), 'EXIF capture time not used: ' + JSON.stringify(src));
  step('capture time taken from EXIF, registration time kept separately');
  await p.screenshot({ path: out + '/shot-shoot.png' });

  // Incident from the equipment
  await p.goto(base + '#eq/EQ-000006'); await p.waitForSelector('#mkcase'); await p.click('#mkcase');
  await p.waitForSelector('#export');
  const caseId = (await p.textContent('.eyebrow')).trim();
  const field = label => p.locator('label.field', { hasText: label }).locator('input,textarea,select').first();
  await field('件名').fill('架空 熱源水ポンプ 異音');
  await field('最初の状況').fill('架空：軸受付近から周期的な異音。漏水なし。');
  await field('現場確認を実施').check();
  await field('現場確認開始').fill('10:35');
  await p.waitForTimeout(800);
  must(!(await p.isDisabled('#export')), 'export should be possible with required fields and no photos');
  step('case ' + caseId + ' filled; export allowed');

  // Select both photos -> captions required -> export blocked
  const tiles = p.locator('#pool .ph');
  await tiles.nth(1).click(); await tiles.nth(0).click();
  await p.waitForSelector('#captions [data-cap]');
  await p.waitForTimeout(300);
  must(await p.isDisabled('#export'), 'export must be blocked until captions are written');
  must((await p.textContent('#probs')).includes('説明文が空'), 'problem list should mention empty caption');
  const caps = p.locator('#captions textarea');
  await caps.nth(0).fill('架空：銘板。型式 HB-65P を確認。');
  await caps.nth(1).fill('架空：吐出圧力計 0.32 MPa。');
  await p.waitForTimeout(300);
  // Switching the case to another equipment drops the photos of the old one
  await p.selectOption('#c-eq', 'EQ-000007'); await p.waitForSelector('text=別の設備の写真を選択から外しました');
  await p.waitForSelector('#export');
  await p.selectOption('#c-eq', 'EQ-000006'); await p.waitForSelector('#pool .ph');
  must((await p.locator('#captions [data-cap]').count()) === 0, 'selection should have been cleared');
  await tiles.nth(1).click(); await tiles.nth(0).click(); await p.waitForSelector('#captions [data-cap]');
  await caps.nth(0).fill('架空：銘板。型式 HB-65P を確認。'); await caps.nth(1).fill('架空：吐出圧力計 0.32 MPa。');
  await p.waitForTimeout(300);
  step('changing the equipment removes photos of the other equipment');
  // Timeline
  await p.click('#tl-add'); await p.locator('#tl textarea').first().fill('架空：運転電流を測定し、異音を確認');
  await p.locator('#tl input[data-k="測定数値など"]').first().fill('15.2 A');
  await p.waitForTimeout(900);
  must(!(await p.isDisabled('#export')), 'export should be enabled now: ' + await p.textContent('#probs'));
  await p.screenshot({ path: out + '/shot-case-photos.png', fullPage: false });
  const [dl] = await Promise.all([p.waitForEvent('download'), p.click('#export')]);
  const pkgPath = out + '/' + dl.suggestedFilename(); await dl.saveAs(pkgPath);
  must(/^HB-REPORT-CASE-\d{8}-\d{4}\.zip$/.test(dl.suggestedFilename()), 'package name ' + dl.suggestedFilename());
  step('report package downloaded: ' + dl.suggestedFilename());

  // Fourth photo cannot be selected beyond 3 (take 2 more, try to select)
  // Backup -> wipe -> restore round trip
  await p.goto(base + '#settings'); await p.waitForSelector('#bk');
  const [bk] = await Promise.all([p.waitForEvent('download'), p.click('#bk')]);
  const bkPath = out + '/' + bk.suggestedFilename(); await bk.saveAs(bkPath);
  // A damaged backup (one photo removed) must be refused without touching current data
  await p.setInputFiles('#rs', { name: 'broken.zip', mimeType: 'application/zip', buffer: require('child_process').execFileSync('python3', ['-I', '-c', `import sys,zipfile,io\nsrc=zipfile.ZipFile(sys.argv[1]);out=io.BytesIO();z=zipfile.ZipFile(out,'w')\nnames=[n for n in src.namelist() if n.startswith('photos/')]\n[z.writestr(n,src.read(n)) for n in src.namelist() if n!=names[0]]\nz.close();sys.stdout.buffer.write(out.getvalue())`, bkPath]) });
  await p.waitForSelector('dialog[open]'); must((await p.textContent('dialog')).includes('復元できません'), 'broken backup should be refused');
  await p.click('dialog button[value=yes]');
  await p.goto(base + '#eq/EQ-000006'); await p.waitForSelector('text=写真 2枚');
  await p.goto(base + '#settings'); await p.waitForSelector('#wipe');
  step('damaged backup refused; existing data untouched');
  await p.click('#wipe'); await p.waitForSelector('dialog[open]'); await p.click('dialog button[value=yes]');
  await p.waitForFunction(() => location.hash === '#home');
  await p.goto(base + '#eq'); await p.waitForSelector('.empty');
  await p.goto(base + '#settings'); await p.waitForSelector('#rs', { state: 'attached' });
  await p.setInputFiles('#rs', bkPath); await p.waitForSelector('dialog[open]');
  must((await p.textContent('dialog')).includes('写真 2枚'), 'restore preview');
  await p.click('dialog button[value=yes]'); await p.waitForFunction(() => location.hash === '#home');
  await p.goto(base + '#case/' + caseId); await p.waitForSelector('#captions [data-cap]');
  must((await p.locator('#captions textarea').first().inputValue()).includes('HB-65P'), 'caption lost in restore');
  step('backup -> wipe -> restore keeps case, photos and captions');

  // Dark mode + desktop width look
  await p.emulateMedia({ colorScheme: 'dark' });
  await p.goto(base + '#eq/EQ-000006'); await p.waitForSelector('#saveL');
  await p.screenshot({ path: out + '/shot-eq-dark.png' });
  const overflow = await p.evaluate(() => document.documentElement.scrollWidth - innerWidth);
  must(overflow <= 0, 'horizontal overflow ' + overflow);
  console.log(JSON.stringify({ pkgPath, bkPath, caseId, errors }));
  await b.close();
})().catch(e => { console.error('FAILED:', e.message); process.exit(1); });
