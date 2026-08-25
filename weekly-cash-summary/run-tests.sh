#!/bin/sh
# Extracts the app's <script> block and runs the ledger test suite under Node.
set -e
cd "$(dirname "$0")"
python3 - <<'PY'
import io, re
s = io.open('Weekly_Cash_Summary_ledger.html', encoding='utf-8').read()
b = [x for x in re.findall(r'<script[^>]*>([\s\S]*?)</script>', s) if len(x) > 500]
io.open('app.js', 'w', encoding='utf-8').write('\n'.join(b))
PY
node --check app.js
for t in test1 test2 test4 test5; do printf '%-8s ' "$t"; node "$t.js" > /dev/null 2>&1 && echo PASS || { echo FAIL; node "$t.js"; exit 1; }; done
