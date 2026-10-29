"""Картинки установщика ENS-Text (NSIS): те же, что у ENS-Table, но в синих тонах.
Запуск: python3 make_images.py (после `tauri icon`: нужен icons/64x64.png)."""

import importlib.util
from pathlib import Path

HERE = Path(__file__).parent
spec = importlib.util.spec_from_file_location('table_images', HERE.parent.parent.parent / 'table' / 'src-tauri' / 'installer' / 'make_images.py')
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)
m.HERE = HERE
m.ICON = HERE.parent / 'icons' / '64x64.png'
m.TOP = (0x2b, 0x57, 0x9a)
m.BOTTOM = (0x1f, 0x3f, 0x73)
m.LIGHT = (0xf3, 0xf6, 0xfb)
m.main()
