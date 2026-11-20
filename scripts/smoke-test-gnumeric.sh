#!/usr/bin/env bash
# Test real Gnumeric imports and SVG chart export (not merely mocked ssconvert).
set -euo pipefail
root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
source_file="${1:-$root/examples/Excel_Chart_Extraction_Stress_Test.xlsx}"
target="${2:-$root/smoke-test-gnumeric-output}"
python_cmd="${PYTHON:-python3}"
command -v ssconvert >/dev/null || { echo 'Install Gnumeric: sudo apt install gnumeric' >&2; exit 2; }
"$python_cmd" "$root/excel_charts_to_svg.py" "$source_file" \
    --engine gnumeric --format svg -o "$target"
"$python_cmd" - "$target/manifest.json" <<'PY2'
import json,sys
from pathlib import Path
s=json.loads(Path(sys.argv[1]).read_text())['summary']
print('Gnumeric chart extraction:',json.dumps(s,indent=2))
if s['charts_enumerated'] != 47 or s['errors']:
    raise SystemExit('Gnumeric import/export did not match the 47-chart fixture')
PY2
