"""Создаёт tests/fixtures/sample.docx (python-docx): образец для проверки импорта и экспорта.
Запуск: python3 tests/make_fixtures.py"""
import io, os
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import struct, zlib

def png(w, h, rgb):
    raw = b''.join(b'\x00' + bytes(rgb) * w for _ in range(h))
    def chunk(t, d):
        c = struct.pack('>I', len(d)) + t + d
        return c + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b'')

d = Document()
sec = d.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3), Cm(1.5), Cm(2), Cm(2)

d.add_heading('Заголовок первого уровня', 1)
p = d.add_paragraph('Обычный абзац с ')
r = p.add_run('полужирным'); r.bold = True
p.add_run(', ')
r = p.add_run('курсивом'); r.italic = True
p.add_run(' и ')
r = p.add_run('подчёркиванием'); r.underline = True
p.add_run(', а также ')
r = p.add_run('красным шрифтом Arial 16'); r.font.name = 'Arial'; r.font.size = Pt(16); r.font.color.rgb = RGBColor(0xC0, 0, 0)
p.add_run('.')

d.add_heading('Заголовок второго уровня', 2)
c = d.add_paragraph('Абзац по центру'); c.alignment = WD_ALIGN_PARAGRAPH.CENTER
j = d.add_paragraph('Абзац по ширине. ' * 12); j.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
j.paragraph_format.first_line_indent = Cm(1.25)

d.add_paragraph('Первый маркированный пункт', style='List Bullet')
d.add_paragraph('Второй маркированный пункт', style='List Bullet')
d.add_paragraph('Вложенный пункт', style='List Bullet 2')
d.add_paragraph('Первый нумерованный пункт', style='List Number')
d.add_paragraph('Второй нумерованный пункт', style='List Number')

t = d.add_table(rows=3, cols=3)
t.style = 'Table Grid'
for i in range(3):
    for k in range(3):
        t.cell(i, k).text = f'Ячейка {i+1}.{k+1}'
a = t.cell(0, 0).merge(t.cell(0, 1))
a.text = 'Объединённая по горизонтали'
b = t.cell(1, 2).merge(t.cell(2, 2))
b.text = 'Объединённая по вертикали'

d.add_paragraph('Рисунок ниже:')
d.add_picture(io.BytesIO(png(120, 60, (30, 120, 200))), width=Cm(4))

# Гиперссылка
p = d.add_paragraph('Сайт: ')
rid = d.part.relate_to('https://example.org/', RT.HYPERLINK, is_external=True)
h = OxmlElement('w:hyperlink'); h.set(qn('r:id'), rid)
r = OxmlElement('w:r'); t_ = OxmlElement('w:t'); t_.text = 'example.org'; r.append(t_); h.append(r)
p._p.append(h)

pb = d.add_paragraph('Перед этим абзацем разрыв страницы')
pb.runs[0].add_break(WD_BREAK.PAGE)
d.add_paragraph('Текст на второй странице.')

out = os.path.join(os.path.dirname(__file__), 'fixtures')
os.makedirs(out, exist_ok=True)
d.save(os.path.join(out, 'sample.docx'))

# Документ с колонтитулами: верхний, нижний с номером страницы и числом страниц, особый на первой странице.
def field(p, instr, cached='1'):
    f = OxmlElement('w:fldSimple'); f.set(qn('w:instr'), f' {instr} ')
    r = OxmlElement('w:r'); t = OxmlElement('w:t'); t.text = cached; r.append(t); f.append(r)
    p._p.append(f)

d2 = Document()
sec2 = d2.sections[0]
sec2.different_first_page_header_footer = True
sec2.header.paragraphs[0].text = 'Колонтитул документа'
sec2.first_page_header.paragraphs[0].text = 'Титульный лист'
fp = sec2.footer.paragraphs[0]
fp.text = 'Страница '
field(fp, 'PAGE'); fp.add_run(' из '); field(fp, 'NUMPAGES', '9')
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
d2.add_paragraph('Текст с колонтитулами')
d2.save(os.path.join(out, 'header.docx'))

# Документ со сноской: версия 0.1 сноски не переносит и должна об этом сказать.
d3 = Document()
p3 = d3.add_paragraph('Текст со сноской')
r3 = OxmlElement('w:r'); fr = OxmlElement('w:footnoteReference'); fr.set(qn('w:id'), '1'); r3.append(fr); p3._p.append(r3)
d3.save(os.path.join(out, 'lossy.docx'))
print('ok')

# Три раздела: книжный, альбомный, книжный (страницы разной ориентации в одном файле).
from docx.enum.section import WD_ORIENT, WD_SECTION
d4 = Document()
d4.add_paragraph('Первая страница, книжная')
s4b = d4.add_section(WD_SECTION.NEW_PAGE)
s4b.orientation = WD_ORIENT.LANDSCAPE
s4b.page_width, s4b.page_height = s4b.page_height, s4b.page_width
d4.add_paragraph('Вторая страница, альбомная')
s4c = d4.add_section(WD_SECTION.NEW_PAGE)
s4c.orientation = WD_ORIENT.PORTRAIT
s4c.page_width, s4c.page_height = s4c.page_height, s4c.page_width
d4.add_paragraph('Третья страница, книжная')
d4.save(os.path.join(out, 'sections.docx'))

# Бланк с плавающими рисунками: логотип в колонтитуле «за текстом» от страницы и картинка с обтеканием в тексте.
import struct, zlib
from docx.shared import Cm
from docx.oxml import parse_xml

def png(path, w, h, rgb):
    raw = b''.join(b'\x00' + bytes(rgb) * w for _ in range(h))
    def chunk(t, d): return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    with open(path, 'wb') as f:
        f.write(b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b''))

WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

def to_anchor(run, behind, pos_h, pos_v, wrap):
    inline = run._r.xpath('.//wp:inline')[0]
    anchor = parse_xml(
        f'<wp:anchor xmlns:wp="{WP}" distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" '
        f'behindDoc="{1 if behind else 0}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>{pos_h}{pos_v}</wp:anchor>')
    for ch in list(inline):
        if ch.tag.endswith('}extent'):
            anchor.append(ch); anchor.append(parse_xml(f'<wp:effectExtent xmlns:wp="{WP}" l="0" t="0" r="0" b="0"/>')); anchor.append(parse_xml(wrap))
        else:
            anchor.append(ch)
    inline.getparent().replace(inline, anchor)

d5 = Document()
png(os.path.join(out, 'logo-blue.png'), 60, 30, (40, 90, 200))
png(os.path.join(out, 'pic-red.png'), 40, 40, (200, 40, 40))
h5 = d5.sections[0].header.paragraphs[0]
r5 = h5.add_run(); r5.add_picture(os.path.join(out, 'logo-blue.png'), width=Cm(6))
to_anchor(r5, True,
    f'<wp:positionH xmlns:wp="{WP}" relativeFrom="page"><wp:posOffset>720000</wp:posOffset></wp:positionH>',
    f'<wp:positionV xmlns:wp="{WP}" relativeFrom="page"><wp:posOffset>360000</wp:posOffset></wp:positionV>',
    f'<wp:wrapNone xmlns:wp="{WP}"/>')
h5.add_run('Шапка')
p5 = d5.add_paragraph('Текст, который обтекает картинку справа. ' * 6)
r5b = p5.add_run(); r5b.add_picture(os.path.join(out, 'pic-red.png'), width=Cm(3))
to_anchor(r5b, False,
    f'<wp:positionH xmlns:wp="{WP}" relativeFrom="margin"><wp:align>left</wp:align></wp:positionH>',
    f'<wp:positionV xmlns:wp="{WP}" relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
    f'<wp:wrapSquare xmlns:wp="{WP}" wrapText="right"/>')
d5.save(os.path.join(out, 'floats.docx'))
os.remove(os.path.join(out, 'logo-blue.png')); os.remove(os.path.join(out, 'pic-red.png'))
