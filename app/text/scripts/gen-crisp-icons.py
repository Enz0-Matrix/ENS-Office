#!/usr/bin/env python3
# SPDX-License-Identifier: AGPL-3.0-or-later
# Создаёт src/icons-crisp.ts: для каждого значка ленты — варианты Fluent, нарисованные под нужный размер
# (32 или 28 px для больших кнопок, 20 px для средних и малых). Значок 24 px, растянутый до 32 px, выглядит
# размытым и тонким: линии попадают между пикселями. Запуск из app/text: python3 scripts/gen-crisp-icons.py
import os, re

D = 'node_modules/@fluentui/svg-icons/icons/'
keys = {}
for f in ['../table/src/icons.ts', 'src/icons.ts']:
    s = open(f, encoding='utf-8').read()
    imp = dict(re.findall(r"import (\w+) from '@fluentui/svg-icons/icons/(\w+)\.svg\?raw'", s))
    for k, v in re.findall(r"'?(\w+_(?:20|24))'?\s*:\s*(\w+)", s):
        if v in imp:
            keys[k] = imp[v]

lines, entries, n = [], [], 0
for key in sorted(keys):
    base = re.sub(r'_(16|20|24|28|32|48)_regular$', '', keys[key])
    large = next((f'{base}_{z}_regular' for z in (32, 28) if os.path.exists(f'{D}{base}_{z}_regular.svg')), None)
    small = f'{base}_20_regular' if os.path.exists(f'{D}{base}_20_regular.svg') else None
    if not large and not small:
        continue
    parts = []
    for kind, name in (('l', large), ('m', small)):
        if name:
            n += 1
            lines.append(f"import c{n} from '@fluentui/svg-icons/icons/{name}.svg?raw';")
            parts.append(f'{kind}: c{n}')
    entries.append(f"  '{key}': {{ {', '.join(parts)} }},")

open('src/icons-crisp.ts', 'w', encoding='utf-8').write(
    '// SPDX-License-Identifier: AGPL-3.0-or-later\n'
    '// Создано scripts/gen-crisp-icons.py — не править руками. Значки Fluent UI System Icons (MIT).\n'
    '// l — для больших кнопок ленты (32 или 28 px), m — для средних и малых (20 px).\n\n'
    + '\n'.join(lines) + '\n\nexport const CRISP: Record<string, { l?: string; m?: string }> = {\n' + '\n'.join(entries) + '\n};\n')
print(len(entries), 'значков,', n, 'файлов')
