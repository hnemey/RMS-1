"""Check every index reference in an xlsx actually resolves.
A dangling index is invisible to XML validation and is the usual reason
Excel offers to 'recover' a file."""
import zipfile, re, sys, posixpath, xml.dom.minidom as md

f = sys.argv[1]
z = zipfile.ZipFile(f); names = z.namelist()
problems = []

st = z.read('xl/styles.xml').decode()
def count(tag):
    m = re.search(r'<%s count="(\d+)"' % tag, st)
    return int(m.group(1)) if m else 0
C = {t: count(t) for t in ['fonts','fills','borders','cellXfs','cellStyleXfs','dxfs']}
numfmt_ids = set(int(x) for x in re.findall(r'<numFmt numFmtId="(\d+)"', st))
print('styles:', C, '| custom numFmts:', sorted(numfmt_ids) or 'none')

# cellXfs internal references
cellxfs = re.search(r'<cellXfs\b.*?</cellXfs>', st, re.S)
if cellxfs:
    for i, xf in enumerate(re.findall(r'<xf\b[^>]*/?>', cellxfs.group(0))):
        for attr, key in (('fontId','fonts'), ('fillId','fills'), ('borderId','borders')):
            m = re.search(attr + r'="(\d+)"', xf)
            if m and int(m.group(1)) >= C[key]:
                problems.append(f'cellXfs[{i}] {attr}={m.group(1)} >= {key} count {C[key]}')
        m = re.search(r'numFmtId="(\d+)"', xf)
        if m and int(m.group(1)) >= 164 and int(m.group(1)) not in numfmt_ids:
            problems.append(f'cellXfs[{i}] numFmtId={m.group(1)} has no numFmt definition')

for n in names:
    if not n.endswith('.xml'):
        continue
    x = z.read(n).decode('utf8', 'replace')
    if n.startswith('xl/worksheets/sheet'):
        for s_ in set(re.findall(r'\bs="(\d+)"', x)):
            if int(s_) >= C['cellXfs']:
                problems.append(f'{n}: cell s="{s_}" >= cellXfs {C["cellXfs"]}')
    for m in re.finditer(r'\b(\w*[Dd]xfId)="(\d+)"', x):
        if int(m.group(2)) >= C['dxfs']:
            problems.append(f'{n}: {m.group(1)}="{m.group(2)}" >= dxfs {C["dxfs"]}')

# relationship ids referenced from sheets must exist in that sheet's rels
for n in names:
    if n.startswith('xl/worksheets/sheet') and n.endswith('.xml'):
        x = z.read(n).decode()
        used = set(re.findall(r'r:id="([^"]+)"', x))
        rel = 'xl/worksheets/_rels/' + posixpath.basename(n) + '.rels'
        have = set(re.findall(r'Id="([^"]+)"', z.read(rel).decode())) if rel in names else set()
        for u in used - have:
            problems.append(f'{n}: r:id="{u}" not in {rel}')

# tables: ref range must match the sheet, columns must match header cells
for n in [x for x in names if x.startswith('xl/tables/table')]:
    x = z.read(n).decode()
    ref = re.search(r'<table\b[^>]*\sref="([^"]+)"', x).group(1)
    af = re.search(r'<autoFilter ref="([^"]+)"', x)
    if af and af.group(1) != ref:
        problems.append(f'{n}: autoFilter ref {af.group(1)} != table ref {ref}')
    ncols = int(re.search(r'<tableColumns count="(\d+)"', x).group(1))
    got = len(re.findall(r'<tableColumn\b', x))
    if got != ncols:
        problems.append(f'{n}: tableColumns count={ncols} but {got} present')
    a, b_ = ref.split(':')
    w = (ord(re.match(r'([A-Z]+)', b_).group(1)) - ord(re.match(r'([A-Z]+)', a).group(1))) + 1
    if w != ncols:
        problems.append(f'{n}: ref {ref} spans {w} cols but declares {ncols}')
    ids = re.findall(r'<tableColumn id="(\d+)"', x)
    if len(set(ids)) != len(ids):
        problems.append(f'{n}: duplicate tableColumn ids')

# queryTable field ids must match their table's column ids
for tno in ['1','2','3']:
    tp, qp = f'xl/tables/table{tno}.xml', f'xl/queryTables/queryTable{tno}.xml'
    if tp in names and qp in names:
        tx, qx = z.read(tp).decode(), z.read(qp).decode()
        tcol = set(re.findall(r'queryTableFieldId="(\d+)"', tx))
        qfld = set(re.findall(r'<queryTableField id="(\d+)"', qx))
        if tcol and not tcol <= qfld:
            problems.append(f'{qp}: table references field ids {sorted(tcol-qfld)} that do not exist')

print('\nPROBLEMS:' if problems else '\nNo dangling references found.')
for p in problems:
    print('  -', p)
print(f'\ntotal: {len(problems)}')
