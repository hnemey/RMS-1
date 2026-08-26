#!/bin/sh
# Extracts the app's <script> block and runs the ledger test suite under Node.
# test7/test8 drive a real DOM and need jsdom (npm i jsdom); they are skipped if absent.
# test8 also needs Rolling_Ave_Ledger.xlsx, produced by `node migrate.js`.
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
for t in test7 test8 test9; do
  printf '%-8s ' "$t"
  if ! node -e "require.resolve('jsdom')" 2>/dev/null; then echo "SKIP (npm i jsdom)"; continue; fi
  case "$t" in test8|test9) [ -f Rolling_Ave_Ledger.xlsx ] || { echo "SKIP (run: node migrate.js)"; continue; };; esac
  node "$t.js" > /dev/null 2>&1 && echo PASS || { echo FAIL; node "$t.js"; exit 1; }
done
