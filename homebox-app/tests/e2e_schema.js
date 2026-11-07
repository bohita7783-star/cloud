// Real-Chromium IndexedDB test of v0.1.3 field editing (synthetic data).
// Usage: node e2e_schema.js <baseUrl> <outDir>
const { chromium } = require('playwright');
const fs = require('fs');
const [base, out] = process.argv.slice(2);
const step = m => console.log('✓', m);
const must = (c, m) => { if (!c) throw new Error(m); };

(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 390, height: 844 }, acceptDownloads: true });
  const p = await ctx.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(base + '#eq-new'); await p.waitForSelector('#f-name');
  await p.fill('#f-name', '架空 給水ポンプ'); await p.click('button[type=submit]');
  await p.waitForFunction(() => /^#eq\/EQ-/.test(location.hash));
  const eqId = decodeURIComponent((await p.evaluate(() => location.hash)).slice(4));
  step('equipment ' + eqId);

  // add a field
  await p.goto(base + '#settings'); await p.waitForSelector('#fieldAdd');
  await p.click('#fieldAdd'); await p.waitForSelector('#schema-name');
  await p.fill('#schema-name', '法定点検区分'); await p.fill('#schema-rule', '架空の社内基準 第1条');
  await p.click('dialog button[value=yes]'); await p.waitForSelector('text=法定点検区分');
  step('field added');

  // value into the new field
  await p.goto(base + '#eq/' + eqId); await p.waitForSelector('#saveL');
  const idx = await p.evaluate(() => [...document.querySelectorAll('table.ledger th label')].findIndex(l => l.textContent === '法定点検区分'));
  must(idx >= 0, 'new field not on ledger');
  await p.fill(`#lv-${idx}`, '第2種（架空）'); await p.selectOption(`#ls-${idx}`, 'ok'); await p.click('#saveL');
  await p.waitForTimeout(400);
  step('value saved in new field');

  // rename, then archive, then reload, then restore -> value must come back
  await p.goto(base + '#settings'); await p.waitForSelector('#schemaUI');
  const key = await p.getAttribute('[data-op=edit][data-fk^="f_"]', 'data-fk');
  await p.click(`[data-op=edit][data-fk="${key}"]`); await p.waitForSelector('#schema-name');
  await p.fill('#schema-name', '法定点検の区分'); await p.click('dialog button[value=yes]'); await p.waitForSelector('text=法定点検の区分');
  await p.click(`[data-op=archive][data-fk="${key}"]`); await p.waitForSelector('dialog[open]'); await p.click('dialog button[value=yes]');
  await p.waitForSelector(`[data-op=restore][data-fk="${key}"]`);
  await p.reload(); await p.waitForSelector(`[data-op=restore][data-fk="${key}"]`);
  step('rename + archive survive a page reload (real IndexedDB)');
  await p.goto(base + '#eq/' + eqId); await p.waitForSelector('#saveL');
  must(!(await p.content()).includes('法定点検の区分'), 'archived field still shown');
  await p.click('#saveL'); await p.waitForTimeout(300); // saving while archived must not wipe the hidden value
  await p.goto(base + '#settings'); await p.waitForSelector(`[data-op=restore][data-fk="${key}"]`);
  await p.click(`[data-op=restore][data-fk="${key}"]`); await p.waitForSelector(`[data-op=archive][data-fk="${key}"]`);
  await p.goto(base + '#eq/' + eqId); await p.waitForSelector('#saveL');
  const idx2 = await p.evaluate(() => [...document.querySelectorAll('table.ledger th label')].findIndex(l => l.textContent === '法定点検の区分'));
  must(idx2 >= 0, 'restored field missing');
  must((await p.inputValue(`#lv-${idx2}`)) === '第2種（架空）', 'value lost after archive/save/restore: ' + await p.inputValue(`#lv-${idx2}`));
  step('value survives archive → save ledger → restore');

  // old header name still imports via alias
  const csv = '﻿設備ID,設備名,法定点検区分\r\n' + `${eqId},架空 給水ポンプ,第1種（架空）\r\n`;
  fs.writeFileSync(out + '/alias.csv', csv);
  await p.goto(base + '#settings'); await p.waitForSelector('#csvin', { state: 'attached' });
  await p.setInputFiles('#csvin', out + '/alias.csv'); await p.waitForSelector('dialog[open]'); console.log('DIALOG:', (await p.textContent('dialog')).replace(/\s+/g,' ').slice(0,300)); await p.click('dialog button[value=yes]'); await p.waitForTimeout(800); console.log('HASH after import:', await p.evaluate(()=>location.hash));
  await p.goto(base + '#eq/' + eqId); await p.waitForSelector('#saveL');
  must((await p.inputValue(`#lv-${idx2}`)) === '第1種（架空）', 'alias import failed');
  step('CSV with the old column name still imports into the renamed field');

  // official item can be archived? (governance check)
  await p.goto(base + '#settings'); await p.waitForSelector('#schemaUI');
  const officialArchivable = await p.locator('[data-op=archive][data-fk="maker"]').count();
  must(officialArchivable === 0, 'official field must not be deletable');
  step('official 17 show no edit/delete buttons');
  // Simulate a phone updated from v0.1.3 where 製造者名 had been hidden and renamed
  await p.goto(base + '#eq/' + eqId); await p.waitForSelector('#saveL');
  await p.fill('#lv-0', '架空機工'); await p.click('#saveL'); await p.waitForTimeout(300);
  await p.evaluate(() => new Promise((res, rej) => { const q = indexedDB.open('homebox'); q.onsuccess = () => {
    const st = q.result.transaction('meta', 'readwrite').objectStore('meta'); const g = st.get('ledger_fields_v1');
    g.onsuccess = () => { const r = g.result; const m = r.v.fields.find(f => f.key === 'maker'); m.active = false; m.label = 'メーカー名';
      r.v.fields.push({ key: 'f_legacystd', label: '架空の基本項目', extra: false, active: true, aliases: [] });
      st.put(r).onsuccess = res; }; }; q.onerror = rej; }));
  await p.reload(); await p.goto(base + '#eq/' + eqId); await p.waitForSelector('#saveL');
  const labels = await p.evaluate(() => [...document.querySelectorAll('table.ledger th label')].map(l => l.textContent));
  must(labels[0] === '製造者名', 'maker not restored: ' + labels[0]);
  must((await p.inputValue('#lv-0')) === '架空機工', 'maker value lost in migration');
  must(labels.indexOf('架空の基本項目') > 16, 'custom standard field should move to the extension group');
  step('v0.1.3 edits to the official 17 are undone on load; values kept; custom field moved to extension group');

  // backup -> wipe -> restore keeps the schema
  await p.goto(base + '#settings'); await p.waitForSelector('#bk');
  const [bk] = await Promise.all([p.waitForEvent('download'), p.click('#bk')]);
  const bkPath = out + '/schema-backup.zip'; await bk.saveAs(bkPath);
  await p.click('#wipe'); await p.waitForSelector('dialog[open]'); await p.click('dialog button[value=yes]'); await p.waitForFunction(() => location.hash === '#home');
  await p.goto(base + '#settings'); await p.waitForSelector('#schemaUI');
  must(!(await p.content()).includes('法定点検の区分'), 'wipe should reset schema');
  await p.setInputFiles('#rs', bkPath); await p.waitForSelector('dialog[open]'); await p.click('dialog button[value=yes]'); await p.waitForFunction(() => location.hash === '#home');
  await p.goto(base + '#settings'); await p.waitForSelector('text=法定点検の区分');
  step('schema restored from backup');
  console.log('errors:', JSON.stringify(errors));
  await b.close();
})().catch(e => { console.error('FAILED:', e.message); process.exit(1); });
