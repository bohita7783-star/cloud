"""The in-app engine must be byte-identical to a released converter.

  python tests/check_engine_sync.py                      # engine files vs engine/ENGINE_SOURCE.json
  python tests/check_engine_sync.py <converter folder>   # also vs the converter release folder
"""
import hashlib
import json
import sys
from pathlib import Path

app = Path(__file__).resolve().parents[1]
src = json.loads((app / 'engine' / 'ENGINE_SOURCE.json').read_text(encoding='utf-8'))
bad = []
for name, digest in src['sha256'].items():
    actual = hashlib.sha256((app / 'engine' / name).read_bytes()).hexdigest()
    if actual != digest:
        bad.append(f'{name}: engine/ differs from ENGINE_SOURCE.json')
    if len(sys.argv) > 1:
        rel = Path(sys.argv[1]) / name
        if hashlib.sha256(rel.read_bytes()).hexdigest() != actual:
            bad.append(f'{name}: engine/ differs from {rel}')
extra = sorted(p.name for p in (app / 'engine').glob('*.py') if p.name not in src['sha256'])
if extra:
    bad.append('engine/ has files not listed in ENGINE_SOURCE.json: ' + ', '.join(extra))
print('engine =', src['converter_version'])
print('\n'.join(bad) if bad else 'OK: in-app engine matches the converter release')
sys.exit(1 if bad else 0)
