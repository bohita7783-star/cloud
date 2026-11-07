// Run: node tests/check_field_order.js
const fs = require('fs');
const assert = require('assert');
const source = fs.readFileSync(require('path').join(__dirname, '..', 'index.html'), 'utf8');
const m = source.match(/const FIELDS = \[([\s\S]*?)\n\];/);
assert(m, 'FIELDS not found');
const rows = [...m[1].matchAll(/\{ key: '([^']+)', label: '([^']+)'(?:, unit: '([^']+)')? \}/g)];
const expected = [
  ['maker','製造者名'], ['equip_no','機器番号'], ['serial_no','製造番号'], ['made','製造年月日'],
  ['installed','設置年月日'], ['location','設置場所'], ['kind_mass','種類と質量'],
  ['motor_maker','電動機メーカー'], ['motor_kw','電動機出力'], ['rated_a','定格電流'],
  ['tank_volume','圧力タンク容量'], ['bore','口径'], ['head','揚程'],
  ['motor_type','電動機形式'], ['voltage_class','電圧区分'], ['rpm','回転数'], ['life','標準耐用年数'],
];
assert.strictEqual(rows.length,17);
assert.deepStrictEqual(rows.map(r=>[r[1],r[2]]),expected);
assert.strictEqual((source.match(/const EXTRA_FIELDS = \[/g)||[]).length,1);
assert(source.includes("const APP_VERSION = '0.4.0'"));
console.log('17-item template: order, names and keys MATCH');
