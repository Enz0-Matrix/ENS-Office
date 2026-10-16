"""Аудит формул в .xlsx без сторонних библиотек: функции, внешние ссылки, рискованные конструкции."""
import zipfile, sys, re, collections, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
path = sys.argv[1]
z = zipfile.ZipFile(path)
rd = lambda n: z.read(n).decode('utf-8')
print('app.xml:', re.findall(r'<Application>(.*?)</Application>|<AppVersion>(.*?)</AppVersion>', rd('docProps/app.xml')))
wb = rd('xl/workbook.xml')
rels = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rd('xl/_rels/workbook.xml.rels')))
rels.update({k: v for v, k in re.findall(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', rd('xl/_rels/workbook.xml.rels'))})
print('calcPr:', re.findall(r'<calcPr[^>]*>', wb))
for i in range(1, 20):
    n = f'xl/externalLinks/_rels/externalLink{i}.xml.rels'
    if n in z.namelist(): print(f'ext[{i}]:', re.findall(r'Target="([^"]+)"', rd(n)))
funcs = collections.Counter(); per_sheet = {}
risky = collections.defaultdict(list)
patterns = {
 'внешняя ссылка [n]': r'\[\d+\]',
 'INDIRECT/OFFSET (летучие)': r'\b(INDIRECT|OFFSET|NOW|TODAY|RAND)\(',
 'текст как число (\"123\"+, --)': r'--|\"\d+[.,]?\d*\"',
 'сравнение с пустой строкой': r'(<>|=)\"\"',
 'целые столбцы A:A': r'(?<![\w$])\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?![\w(])',
 'массивные/динамические': r'_xlfn\.|_xlws\.',
 'ROUND отсутствует при делении': r'/',
}
for sid, name, state, rid in re.findall(r'<sheet name="([^"]*)" sheetId="(\d+)" state="(\w+)" r:id="(rId\d+)"', wb) and [] or []:
    pass
sheets = re.findall(r'<sheet name="([^"]*)" sheetId="\d+" state="(\w+)" r:id="(rId\d+)"', wb)
for name, state, rid in sheets:
    target = 'xl/' + rels[rid].lstrip('/').replace('xl/', '')
    x = rd(target)
    cells = re.findall(r'<c r="([A-Z]+\d+)"([^>]*)>(.*?)</c>', x, re.S)
    nf = 0; arr = 0; shared = 0
    for ref, attrs, inner in cells:
        m = re.search(r'<f([^>]*)>(.*?)</f>|<f([^>]*)/>', inner, re.S)
        if not m: continue
        nf += 1
        fattr = (m.group(1) or m.group(3) or '')
        if 't="array"' in fattr: arr += 1
        if 't="shared"' in fattr: shared += 1
        f = (m.group(2) or '').replace('&quot;', '"').replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
        for fn in re.findall(r'([A-Z_][A-Z0-9_.]*)\(', f): funcs[fn] += 1
        for k, p in patterns.items():
            if re.search(p, f) and len(risky[k]) < 6: risky[k].append(f'{name}!{ref}: ={f[:150]}')
    per_sheet[name] = (state, len(cells), nf, arr, shared)
print('\nЛисты (состояние, ячеек, формул, массивных, общих):')
for k, v in per_sheet.items(): print(f'  {k}: {v}')
print('\nФункции:', funcs.most_common())
print('\nПримеры рискованных конструкций:')
for k, v in risky.items():
    print(' #', k); [print('   ', s) for s in v]
