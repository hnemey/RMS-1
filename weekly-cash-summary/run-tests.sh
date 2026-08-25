#!/bin/sh
# Extracts the app's <script> block and runs the ledger test suite under Node.
# test7 drives the real DOM and needs jsdom (npm i jsdom); it is skipped if absent.
set -e
cd "$(dirname "$0")"
python3 - <<'PY'
import io, re
s = io.open('Weekly_Cash_Summary_ledger.html', encoding='utf-8').read()
b = [x for x in re.findall(r'<script[^>]*>([\s\S]*?)</script>', s) if len(x) > 500]
io.open('app.js', 'w', encoding='utf-8').write('\n'.join(b))
PY
node --check app.js
for t in test1 test2 test4 test5 test6; do
  printf '%-8s ' "$t"
  node "$t.js" > /dev/null 2>&1 && echo PASS || { echo FAIL; node "$t.js"; exit 1; }
done
printf '%-8s ' test7
if node -e "require.resolve('jsdom')" 2>/dev/null; then
  node test7.js > /dev/null 2>&1 && echo PASS || { echo FAIL; node test7.js; exit 1; }
else
  echo "SKIP (npm i jsdom)"
fi
