const fs=require('fs'); const vm=require('vm');
const js=(fs.readFileSync(require('path').join(__dirname,'..','index.html'),'utf8').match(/<script>([\s\S]*?)<\/script>/) || [,''])[1];
const fn=js.match(/  function renderName\(\) \{[\s\S]*?\n  \}\n  function renderEq/);
if(!fn) throw new Error('renderName not found');
const s=fn[0].slice(0,-'\n  function renderEq'.length);
const sel={'#filename':{innerHTML:''},'#fileStatus':{textContent:''}};
const EQ=[{id:'EQ-000006'},{id:'EQ-000011'}];const kinds=[['銘板'],['全景'],['圧力計']];
let ei=0,ki=0;
const shots={'EQ-000006':[], 'EQ-000011':[]};const counts={'EQ-000006':0,'EQ-000011':0};
const ctx={shots,counts,CASE:'CASE-20261107-0001',pad:n=>String(n).padStart(3,'0'),cur:()=>({e:EQ[ei],k:kinds[ki]}),$:key=>sel[key]};
vm.createContext(ctx);vm.runInContext(s+'\nthis.renderName = renderName;',ctx);
function assert(name,pattern){ctx.renderName(); if(!pattern.test(sel['#filename'].innerHTML))throw new Error(name+' failed: '+sel['#filename'].innerHTML); console.log('PASS',name, sel['#fileStatus'].textContent);}
assert('before photo / 001 銘板',/<span class="seg s-no">001<\/span>_<span class="seg s-kind">銘板/);
counts[EQ[0].id]++;shots[EQ[0].id].push({n:1,kind:'銘板'});assert('after first photo same 001 銘板',/<span class="seg s-no">001<\/span>_<span class="seg s-kind">銘板/);
ki=2;assert('change pending type does not retroactively rename previous photo',/<span class="seg s-no">001<\/span>_<span class="seg s-kind">銘板/);
counts[EQ[0].id]++;shots[EQ[0].id].push({n:2,kind:'圧力計'});assert('after second photo 002 圧力計',/<span class="seg s-no">002<\/span>_<span class="seg s-kind">圧力計/);
ei=1;ki=1;assert('switch eq empty history 001 全景',/<span class="seg s-no">001<\/span>_<span class="seg s-kind">全景/);
