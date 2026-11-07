// HomeBox report engine worker.
// Runs the Field Survey converter (Python, the same files as the Termux release)
// inside Pyodide, entirely on this device. Nothing is sent anywhere: the worker
// only reads files served by the HomeBox origin and the bytes the page passes in.
'use strict';
const PYODIDE_DIR = './pyodide/0.27.7/';
const ENGINE_FILES = ['engine_policy.py', 'approval.py', 'report_layout.py', 'converter_mobile.py', 'homebox_adapter.py', 'case_exports.py'];
importScripts(PYODIDE_DIR + 'pyodide.js');

let bootPromise = null;
const started = performance.now();
const ms = () => Math.round(performance.now() - started);

function post(id, type, payload) { postMessage({ id, type, ...payload }); }

async function fetchBytes(url) {
  const r = await fetch(url);
  if (!r.ok) throw new Error('読み込めません: ' + url);
  return new Uint8Array(await r.arrayBuffer());
}

function boot(id) {
  if (bootPromise) return bootPromise;
  bootPromise = (async () => {
    const t = {};
    post(id, 'progress', { stage: 'Pythonを起動しています' });
    const py = await loadPyodide({ indexURL: new URL(PYODIDE_DIR, location.href).href });
    t.python = ms();
    post(id, 'progress', { stage: '部品（lxml・Pillow・cryptography）を読み込んでいます' });
    await py.loadPackage(['lxml', 'pillow', 'cryptography'], { messageCallback: () => { }, errorCallback: () => { } });
    t.packages = ms();
    post(id, 'progress', { stage: '変換エンジンを準備しています' });
    py.FS.mkdirTree('/engine');
    for (const f of ENGINE_FILES) py.FS.writeFile('/engine/' + f, await fetchBytes('./engine/' + f));
    await py.runPythonAsync(`
import sys
sys.path.insert(0, '/engine')
import converter_mobile, approval, homebox_adapter, case_exports, engine_policy, report_layout
`);
    t.engine = ms();
    return { py, t };
  })();
  bootPromise.catch(() => { bootPromise = null; });
  return bootPromise;
}

function resetWork(py) {
  py.runPython(`
import shutil, os
shutil.rmtree('/work', ignore_errors=True)
os.makedirs('/work/out')
`);
}

// Python exceptions arrive as a long traceback; show only the final message.
function cleanError(e) {
  const text = String(e && e.message || e);
  const lines = text.trim().split('\n').filter(Boolean);
  const last = lines[lines.length - 1] || text;
  return last.replace(/^[\w.]+(Error|Exception): /, '').replace(/^(APPROVAL|HOMEBOX|SAFETY) STOP: /, '停止：');
}

onmessage = async (e) => {
  const { id, cmd } = e.data;
  try {
    const { py, t } = await boot(id);
    if (cmd === 'warm') return post(id, 'done', { result: { timings: t, version: py.runPython('converter_mobile.VERSION') } });
    if (cmd === 'layout') {
      // The same layout the Excel gets, without a template: for the "様式に入るか確認" preview.
      post(id, 'progress', { stage: '様式に当てはめています' });
      resetWork(py);
      py.FS.writeFile('/work/layout.json', e.data.payload);
      const json = await py.runPythonAsync(`
import json, converter_mobile as CM
d = json.load(open('/work/layout.json', encoding='utf-8'))
case = dict(d['case'], _contacts=d.get('contacts') or [])
r, p, w = CM.build_updates_with_notes(case, d.get('photos') or [], d.get('timeline') or [])
json.dumps({'cells': {k: (v if isinstance(v, str) else None) for k, v in r.items()}, 'warnings': w}, ensure_ascii=False)
`);
      return post(id, 'done', { result: JSON.parse(json) });
    }
    resetWork(py);
    py.FS.writeFile('/work/template.xlsx', e.data.template);
    py.FS.writeFile('/work/approved.json', e.data.manifest);
    py.FS.writeFile('/work/trusted.json', await fetchBytes('./trusted_approvers.json'));
    if (cmd === 'verify') {
      post(id, 'progress', { stage: '署名と原本を照合しています' });
      const json = await py.runPythonAsync(`
import json, approval as AP
m = AP.verify_manifest('/work/template.xlsx', '/work/approved.json', '/work/trusted.json')
json.dumps({k: m['approval'].get(k) for k in ('approval_id', 'approver', 'approved_at', 'valid_until')} | {'template_sha256': m['template']['sha256']}, ensure_ascii=False)
`);
      return post(id, 'done', { result: JSON.parse(json) });
    }
    if (cmd === 'convert') {
      post(id, 'progress', { stage: '会社様式のExcelを作り、全部品を検査しています' });
      py.FS.writeFile('/work/package.zip', e.data.package);
      const t0 = performance.now();
      const json = await py.runPythonAsync(`
import json, datetime, uuid, converter_mobile as CM, homebox_adapter as HB, case_exports as CE
case, _, _, _ = HB.read_package('/work/package.zip')
name = CE.output_report_name(case.get('件名'), datetime.datetime.now().strftime('%Y%m%d%H%M'), uuid.uuid4().hex[:8])
review = CM.make_report_from_homebox('/work/package.zip', '/work/template.xlsx', '/work/out/report.xlsx', '/work/out/audit.json',
                                     approval_manifest='/work/approved.json', trusted_keys='/work/trusted.json')
json.dumps({'name': name, 'approval_id': review['approval']['approval_id'], 'photo_count': review['photo_count'],
            'timeline_count': review['timeline_count'], 'unmapped': review.get('unmapped_nonempty_fields') or [],
            'version': review['version']}, ensure_ascii=False)
`);
      const result = JSON.parse(json);
      result.convertMs = Math.round(performance.now() - t0);
      result.timings = t;
      const xlsx = py.FS.readFile('/work/out/report.xlsx');
      const audit = py.FS.readFile('/work/out/audit.json');
      return post(id, 'done', { result, xlsx, audit });
    }
    throw new Error('不明な命令です');
  } catch (err) {
    post(id, 'error', { message: cleanError(err) });
  }
};
