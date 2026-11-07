// Test the real schema functions extracted from the PWA (no browser storage required).
// Run: node tests/test_dynamic_schema.js
'use strict';
const fs=require('fs'),vm=require('vm'),assert=require('assert'),crypto=require('crypto');
const src=fs.readFileSync(require('path').join(__dirname,'..','index.html'),'utf8');
const piece=src.slice(src.indexOf('const FIELDS = ['),src.indexOf('const STATUS = '));
assert(piece.startsWith('const FIELDS = [') && piece.includes('async function updateFieldSchema('));
const meta=new Map();
const ctx=vm.createContext({
  crypto:{randomUUID:()=>crypto.randomUUID()},
  nowIso:()=> '2026-11-07T12:00:00+09:00',
  get:async(n,k)=>meta.get(k),
  put:async(n,v)=>{if(n!=='meta')throw Error(n);meta.set(v.key,JSON.parse(JSON.stringify(v)));},
  rid:()=> 'tmp_fallback'
});
vm.runInContext(piece+`\nglobalThis.s={FIELDS,EXTRA_FIELDS,ALL_FIELDS,fieldSchema:()=>fieldSchema,fieldSchemaRevision:()=>fieldSchemaRevision,DEFAULT_FIELD_SCHEMA,validateFieldSchema,applyFieldSchema,updateFieldSchema,loadFieldSchema,validateNewField,makeFieldKey,normalizeFieldSchema};`,ctx);
const s=ctx.s;
async function test(){
 await s.loadFieldSchema();
 assert.strictEqual(s.FIELDS.length,17);assert.strictEqual(s.EXTRA_FIELDS.length,6);console.log('✓ default ledger: 17 + 6');
 const official=['製造者名','機器番号','製造番号','製造年月日','設置年月日','設置場所','種類と質量','電動機メーカー','電動機出力','定格電流','圧力タンク容量','口径','揚程','電動機形式','電圧区分','回転数','標準耐用年数'];
 assert.deepStrictEqual(Array.from(s.FIELDS,x=>x.label),official);console.log('✓ 17 original fields/order kept');
 const data={};data['maker']={v:'架空メーカー',st:'ok',src:'テスト銘板'};
 const k=s.makeFieldKey();const custom={key:k,label:'法定点検区分',rule:'架空規程第1条',extra:true,active:true,aliases:[]};
 await s.updateFieldSchema([...s.fieldSchema(),custom],'add');
 assert.strictEqual(s.ALL_FIELDS.length,24);data[k]={v:'第一種',st:'ok',src:'テスト規程'};
 console.log('✓ added custom regulatory field + source');
 let list=s.fieldSchema().map(x=>({...x,aliases:[...(x.aliases||[])]}));
 list.find(x=>x.key===k).label='点検区分';list.find(x=>x.key===k).aliases=['法定点検区分'];
 await s.updateFieldSchema(list,'rename');
 assert.strictEqual(s.EXTRA_FIELDS.find(x=>x.key===k).label,'点検区分');
 assert.strictEqual(data[k].v,'第一種'); console.log('✓ stable key on renaming, historical alias preserved');
 list=s.fieldSchema().map(x=>({...x}));list.find(x=>x.key===k).active=false;
 await s.updateFieldSchema(list,'archive');
 assert(!s.ALL_FIELDS.some(x=>x.key===k));assert.strictEqual(data[k].v,'第一種');console.log('✓ archive hides without deleting values');
 await s.loadFieldSchema();assert(!s.ALL_FIELDS.some(x=>x.key===k));console.log('✓ archive/metadata survived reload');
 list=s.fieldSchema().map(x=>({...x}));list.find(x=>x.key===k).active=true;
 await s.updateFieldSchema(list,'restore');
 assert(s.ALL_FIELDS.some(x=>x.key===k));assert.strictEqual(data[k].st,'ok');console.log('✓ restore reveals original value and status');
 let conflict=s.fieldSchema().map(x=>({...x}));conflict.find(x=>x.key===k).label='製造者名';
 assert.throws(()=>s.validateFieldSchema(conflict),/重複/);
 assert.throws(()=>s.validateNewField('設備ID','',k),/基本列/);
 console.log('✓ reject conflicting names, protect CSV identifiers');
 const invalid=s.fieldSchema().filter(x=>x.key!=='maker');
 assert.throws(()=>s.validateFieldSchema(invalid),/識別子/);console.log('✓ reject data-destructive schema');
 // v0.1.4: the official 17 are fixed
 const hideOfficial=s.fieldSchema().map(x=>({...x}));hideOfficial.find(x=>x.key==='maker').active=false;
 assert.throws(()=>s.validateFieldSchema(hideOfficial),/正式17項目/);
 const renameOfficial=s.fieldSchema().map(x=>({...x}));renameOfficial.find(x=>x.key==='maker').label='メーカー';
 assert.throws(()=>s.validateFieldSchema(renameOfficial),/正式17項目/);
 const reorder=s.fieldSchema().map(x=>({...x}));[reorder[0],reorder[1]]=[reorder[1],reorder[0]];
 assert.throws(()=>s.validateFieldSchema(reorder),/並び順/);
 const customStandard=[...s.fieldSchema().map(x=>({...x})),{key:'f_custom',label:'架空基本',extra:false,active:true,aliases:[]}];
 assert.throws(()=>s.validateFieldSchema(customStandard),/追加項目に/);
 const n=s.normalizeFieldSchema(hideOfficial.map(x=>x.key==='maker'?{...x,label:'メーカー'}:x));
 assert(n.changed);s.validateFieldSchema(n.fields);
 assert.strictEqual(n.fields.find(x=>x.key==='maker').label,'製造者名');assert(n.fields.find(x=>x.key==='maker').aliases.includes('メーカー'));
 console.log('✓ official 17 fixed; v0.1.3 edits normalised back (old name kept as CSV alias)');
 assert.strictEqual(meta.get('ledger_fields_v1').v.revision,4);
 console.log('PASS: 9 checks, actual schema code');
}
test().catch(e=>{console.error(e);process.exit(1)});
