"""Каталог функций по категориям для вкладки «Формулы» (src/funclib.ts).

Категории берутся из комментариев в core/ironcalc/base/src/functions/mod.rs (таблица
impl_function_lookup), русские и английские имена — из core/ironcalc/generate_language/src/languages.json.
Запуск: python3 scripts/gen-funclib.py
"""
import json
import re
from pathlib import Path

root = Path(__file__).resolve().parents[3]
mod = (root / 'core/ironcalc/base/src/functions/mod.rs').read_text(encoding='utf-8')
langs = json.loads((root / 'core/ironcalc/generate_language/src/languages.json').read_text(encoding='utf-8'))
ru, en = langs['ru']['functions'], langs['en']['functions']

CATEGORY = {
    'Logical': 'Логические',
    'Mathematical and trigonometry': 'Математические',
    'Information': 'Информационные',
    'Lookup and reference': 'Ссылки и массивы',
    'Dynamic array': 'Ссылки и массивы',
    'Text': 'Текстовые',
    'Statistical': 'Статистические',
    'More statistical': 'Статистические',
    'Compatibility': 'Совместимость',
    'Date and time': 'Дата и время',
    'Financial': 'Финансовые',
    'Engineering': 'Инженерные',
    'Database': 'Работа с базой данных',
}

start = mod.index('impl_function_lookup! {')
body = mod[start:]
body = body[: body.index('\n}\n')]
out: dict[str, list[tuple[str, str]]] = {}
cat = None
for line in body.splitlines()[1:]:
    line = line.strip()
    if line.startswith('//'):
        head = line[2:].strip().split(':')[0].split('—')[0].strip()
        cat = CATEGORY.get(head, cat)
        continue
    m = re.match(r'(?:r#)?(\w+)\s*=>', line)
    if not m or cat is None:
        continue
    key = m.group(1)
    if key not in ru or key not in en:
        continue
    pair = (ru[key], en[key])
    if pair not in out.setdefault(cat, []):
        out[cat].append(pair)

lines = ['// Создано scripts/gen-funclib.py: функции по категориям вкладки «Формулы» — [русское имя, английское].', '',
         'export const FUNCTION_CATEGORIES: Record<string, [string, string][]> = {']
for name, items in out.items():
    items.sort(key=lambda p: p[0].lower())
    lines.append(f'  {json.dumps(name, ensure_ascii=False)}: [')
    lines.append('    ' + ', '.join(f'[{json.dumps(a, ensure_ascii=False)}, {json.dumps(b)}]' for a, b in items) + ',')
    lines.append('  ],')
lines.append('};')
(root / 'app/table/src/funclib.ts').write_text('\n'.join(lines) + '\n', encoding='utf-8')
print({k: len(v) for k, v in out.items()})
