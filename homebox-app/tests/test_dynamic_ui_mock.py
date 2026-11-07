import asyncio
from pathlib import Path
from playwright.async_api import async_playwright
src=(Path(__file__).resolve().parent.parent / 'index.html').read_text()
MOCK=r'''(() => {
const maps={}; let initialized=false;
const request=(val, fn) => { const r={}; queueMicrotask(()=>{try{r.result=fn?fn():val;r.onsuccess?.({target:r});}catch(e){r.error=e;r.onerror?.({target:r});}}); return r; };
const store=name=>({createIndex(){return {}},index(i){return {getAll(v){return request(null,()=>[...maps[name].values()].filter(x=>x[i]===v))}}},get(k){return request(null,()=>maps[name].get(k))},getAll(){return request(null,()=>[...maps[name].values()])},put(v){return request(null,()=>{maps[name].set(v.key??v.id,v);return v.key??v.id;})},delete(k){return request(null,()=>maps[name].delete(k))},clear(){return request(null,()=>maps[name].clear())}});
const db={createObjectStore(name){maps[name]=new Map();return store(name)},transaction(name){return {objectStore(n){return store(n)}}}};
Object.defineProperty(window,'indexedDB',{value:{open(){const r={result:db};queueMicrotask(()=>{if(!initialized){initialized=true;r.onupgradeneeded?.()}r.onsuccess?.()});return r}},configurable:true});
})();'''
async def main():
 async with async_playwright() as p:
  b=await p.chromium.launch(executable_path='/usr/bin/chromium',args=['--no-sandbox'])
  pg=await b.new_page(viewport={'width':390,'height':844})
  errs=[];pg.on('pageerror',lambda e:errs.append(str(e)))
  await pg.evaluate(MOCK)
  await pg.set_content(src,wait_until='load')
  await pg.evaluate("location.hash='#settings'")
  await pg.wait_for_selector('#schemaUI [data-op]',timeout=10000)
  print('schema visible',await pg.locator('.schema-line').count())
  print('click add', flush=True);await pg.locator('#fieldAdd').evaluate('(el)=>el.click()');print('clicked add',flush=True)
  await pg.fill('dialog #schema-name','法定点検区分')
  await pg.fill('dialog #schema-rule','架空省令第1条')
  print('dialog submit',flush=True);await pg.locator('dialog button[value=yes]').evaluate('(el)=>el.click()')
  await pg.wait_for_timeout(250);print('dialog closed; schema entries',await pg.locator('.schema-line').count(),flush=True);await pg.wait_for_selector('.schema-line b:text-is("法定点検区分")',timeout=2500)
  print('after add',await pg.locator('.schema-line').count())
  await pg.evaluate("location.hash='#eq-new'")
  print('navigating eq-new',flush=True);await pg.wait_for_selector('#f-name',timeout=2500);print('eq-new visible',flush=True)
  await pg.fill('#f-name','架空設備1')
  await pg.locator('#f').evaluate('(el)=>el.dispatchEvent(new Event("submit",{bubbles:true,cancelable:true}))')
  print('waiting ledger',flush=True);await pg.wait_for_selector('#saveL',timeout=2500);print('ledger visible',flush=True)
  row=pg.locator('tr',has=pg.locator('label:text-is(\"法定点検区分\")'))
  await row.locator('input[type=text]').first.fill('第1種')
  await row.locator('select').select_option('ok')
  await row.locator('input[type=text]').nth(1).fill('架空基準')
  await pg.locator('#saveL').evaluate('(el)=>el.click()')
  await pg.wait_for_timeout(50)
  await pg.evaluate("location.hash='#settings'")
  await pg.wait_for_selector('#schemaUI [data-op]')
  print('equipment with field saved',flush=True)

  li=pg.locator('.schema-line',has=pg.locator('b:text-is("法定点検区分")'))
  print('edit',flush=True);await li.locator('[data-op=edit]').evaluate('(el)=>el.click()')
  await pg.fill('dialog #schema-name','点検区分')
  print('dialog submit',flush=True);await pg.locator('dialog button[value=yes]').evaluate('(el)=>el.click()')
  await pg.wait_for_selector('.schema-line b:text-is("点検区分")')
  li=pg.locator('.schema-line',has=pg.locator('b:text-is("点検区分")'))
  print('archive',flush=True);await li.locator('[data-op=archive]').evaluate('(el)=>el.click()')
  print('dialog submit',flush=True);await pg.locator('dialog button[value=yes]').evaluate('(el)=>el.click()')
  li=pg.locator('.schema-line',has=pg.locator('b:text-is("点検区分")'))
  await pg.wait_for_timeout(100)
  assert '非表示' in await li.inner_text()
  await pg.evaluate("location.hash='#eq/EQ-000001'")
  print('waiting ledger',flush=True);await pg.wait_for_selector('#saveL',timeout=2500);print('ledger visible',flush=True)
  assert await pg.locator('label:text-is("点検区分")').count()==0
  await pg.evaluate("location.hash='#settings'")
  await pg.wait_for_selector('#schemaUI [data-op]')
  li=pg.locator('.schema-line',has=pg.locator('b:text-is("点検区分")'))

  print('restore',flush=True);await li.locator('[data-op=restore]').evaluate('(el)=>el.click()')
  li=pg.locator('.schema-line',has=pg.locator('b:text-is("点検区分")'))
  assert '非表示' not in await li.inner_text()
  assert '架空省令第1条' in await li.inner_text()
  await pg.evaluate("location.hash='#eq/EQ-000001'")
  print('waiting ledger',flush=True);await pg.wait_for_selector('#saveL',timeout=2500);print('ledger visible',flush=True)
  row=pg.locator('tr',has=pg.locator('label:text-is("点検区分")'))
  assert await row.locator('input[type=text]').first.input_value()=='第1種'
  assert await row.locator('select').input_value()=='ok'
  assert await row.locator('input[type=text]').nth(1).input_value()=='架空基準'
  print('stored equipment field survived rename/archive/restore',flush=True)

  await pg.evaluate("location.hash='#settings'")
  await pg.wait_for_selector('#schemaUI [data-op]')
  cfg=await pg.evaluate('({schema:FIELD_FILE_SCHEMA, fields:fieldSchema, revision:fieldSchemaRevision})')
  import json
  fieldNew={**cfg['fields'][-1], 'label':'新・点検区分', 'aliases':['法定点検区分','点検区分']}
  cfg['fields'][-1]=fieldNew
  await pg.locator('#fieldImport').set_input_files({'name':'schema.json','mimeType':'application/json','buffer':json.dumps(cfg,ensure_ascii=False).encode()})
  await pg.wait_for_selector('dialog[open]')
  await pg.locator('dialog button[value=yes]').evaluate('(el)=>el.click()')
  await pg.wait_for_selector('.schema-line b:text-is("新・点検区分")')
  print('field schema JSON import success',flush=True)
  # Screenshots are optional for the mock UI test.
  print('BROWSER MOCK PASS, errs',errs)
  await b.close()
asyncio.run(main())
