"""Печать содержимого листов .xlsx как текста (строка = ячейки через ' | '), без сторонних библиотек."""
import zipfile, sys, re, io, html
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
z = zipfile.ZipFile(sys.argv[1]); rd = lambda n: z.read(n).decode('utf-8')
ss = []
if 'xl/sharedStrings.xml' in z.namelist():
    for si in re.findall(r'<si>(.*?)</si>', rd('xl/sharedStrings.xml'), re.S):
        ss.append(html.unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', si, re.S))))
wb = rd('xl/workbook.xml'); rels = rd('xl/_rels/workbook.xml.rels')
for name, rid in re.findall(r'<sheet [^>]*name="([^"]*)"[^>]*r:id="(rId\d+)"', wb):
    tgt = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    print('=====', html.unescape(name))
    x = rd('xl/' + tgt.group(1).replace('/xl/', '').lstrip('/'))
    for row in re.findall(r'<row[^>]*>(.*?)</row>', x, re.S):
        out = []
        for attrs, inner in re.findall(r'<c ([^>]*?)(?:/>|>(.*?)</c>)', row, re.S):
            v = re.search(r'<v>(.*?)</v>', inner or ''); t = re.search(r't="(\w+)"', attrs)
            if t and t.group(1) == 's' and v: out.append(ss[int(v.group(1))])
            elif t and t.group(1) == 'inlineStr': out.append(html.unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', inner, re.S))))
            elif v: out.append(html.unescape(v.group(1)))
        if out: print(' | '.join(s.replace('\n', ' ') for s in out))
