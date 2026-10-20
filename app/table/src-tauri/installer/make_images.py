"""Картинки установщика ENS-Table (NSIS): боковая панель 164×314 и шапка 150×57.

Нефритовый градиент с тонкой сеткой, как в шапке программы, и значок
приложения. Без сторонних библиотек: python make_images.py
"""

import struct
import zlib
from pathlib import Path

HERE = Path(__file__).parent
ICON = HERE.parent / 'icons' / '64x64.png'

TOP = (0x2b, 0x7a, 0x6b)     # --brand: нефрит
BOTTOM = (0x1d, 0x5a, 0x4f)  # темнее
LIGHT = (0xf4, 0xfa, 0xf8)   # светлый фон шапки


def read_png(path):
    """RGBA 8 бит без чересстрочности (так сохраняет tauri icon)."""
    return decode_png(path.read_bytes())


def decode_png(data):
    pos, idat, w, h, ctype = 8, b'', 0, 0, 0
    while pos < len(data):
        n = struct.unpack('>I', data[pos:pos + 4])[0]
        kind, body = data[pos + 4:pos + 8], data[pos + 8:pos + 8 + n]
        if kind == b'IHDR':
            w, h, depth, ctype = struct.unpack('>IIBB', body[:10])
            assert depth == 8 and ctype in (2, 6), 'нужен PNG RGB/RGBA 8 бит'
        elif kind == b'IDAT':
            idat += body
        pos += 12 + n
    bpp = 4 if ctype == 6 else 3
    raw = zlib.decompress(idat)
    stride = w * bpp
    rows, prev, i = [], bytearray(stride), 0
    for _ in range(h):
        f, line = raw[i], bytearray(raw[i + 1:i + 1 + stride])
        i += 1 + stride
        for x in range(stride):
            a = line[x - bpp] if x >= bpp else 0
            b = prev[x]
            c = prev[x - bpp] if x >= bpp else 0
            if f == 1:
                line[x] = (line[x] + a) & 255
            elif f == 2:
                line[x] = (line[x] + b) & 255
            elif f == 3:
                line[x] = (line[x] + (a + b) // 2) & 255
            elif f == 4:
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                line[x] = (line[x] + (a if pa <= pb and pa <= pc else b if pb <= pc else c)) & 255
        rows.append(line)
        prev = line
    px = [[tuple(r[x * bpp:x * bpp + bpp]) + ((255,) if bpp == 3 else ()) for x in range(w)] for r in rows]
    return w, h, px


def mix(a, b, t):
    return tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def canvas(w, h, fill):
    return [[fill(x, y) for x in range(w)] for y in range(h)]


def paste(img, icon, ox, oy, scale=1):
    iw, ih, px = icon
    for y in range(ih * scale):
        for x in range(iw * scale):
            r, g, b, a = px[y // scale][x // scale]
            if a and 0 <= oy + y < len(img) and 0 <= ox + x < len(img[0]):
                img[oy + y][ox + x] = mix(img[oy + y][ox + x], (r, g, b), a / 255)


def write_bmp(path, img):
    h, w = len(img), len(img[0])
    stride = (w * 3 + 3) & ~3
    body = bytearray()
    for row in reversed(img):
        line = bytearray()
        for r, g, b in row:
            line += bytes((b, g, r))
        body += line + b'\0' * (stride - len(line))
    head = struct.pack('<2sIHHI', b'BM', 54 + len(body), 0, 0, 54)
    info = struct.pack('<IiiHHIIiiII', 40, w, h, 1, 24, 0, len(body), 2835, 2835, 0, 0)
    path.write_bytes(head + info + body)


def jade(w, h):
    def fill(x, y):
        c = mix(TOP, BOTTOM, (y / h) * 0.8 + (x / w) * 0.2)
        if x % 16 == 0 or y % 16 == 0:  # едва заметная сетка «как в таблице»
            c = mix(c, (255, 255, 255), 0.07)
        return c
    return canvas(w, h, fill)


def fix_ico(path):
    """Значки 16–64 — в классическом формате BMP: установщик NSIS и часть
    Windows не показывают маленькие значки, сжатые в PNG (как их пишет
    tauri icon). 256×256 остаётся в PNG."""
    data = path.read_bytes()
    count = struct.unpack('<HHH', data[:6])[2]
    images = []
    for i in range(count):
        w, h, _, _, _, _, size, off = struct.unpack('<BBBBHHII', data[6 + 16 * i:22 + 16 * i])
        blob = data[off:off + size]
        w, h = w or 256, h or 256
        if blob[1:4] == b'PNG' and w < 256:
            _, _, px = decode_png(blob)
            dib = bytearray(struct.pack('<IiiHHIIiiII', 40, w, h * 2, 1, 32, 0, 0, 0, 0, 0, 0))
            for row in reversed(px):
                for r, g, b, a in row:
                    dib += bytes((b, g, r, a))
            mask_stride = ((w + 31) // 32) * 4
            for row in reversed(px):
                bits = bytearray(mask_stride)
                for x, (_, _, _, a) in enumerate(row):
                    if a == 0:
                        bits[x // 8] |= 0x80 >> (x % 8)
                dib += bits
            blob = bytes(dib)
        images.append((w, h, blob))
    out = bytearray(struct.pack('<HHH', 0, 1, len(images)))
    off = 6 + 16 * len(images)
    for w, h, blob in images:
        out += struct.pack('<BBBBHHII', w % 256, h % 256, 0, 0, 1, 32, len(blob), off)
        off += len(blob)
    for _, _, blob in images:
        out += blob
    path.write_bytes(bytes(out))


def main():
    icon = read_png(ICON)
    side = jade(164, 314)
    paste(side, icon, (164 - 64) // 2, 90)
    write_bmp(HERE / 'sidebar.bmp', side)

    def light(x, y):
        c = LIGHT
        if x > 92:  # справа — нефритовая полоса с градиентом
            c = mix(LIGHT, TOP, min(1, (x - 92) / 40))
        return c
    header = canvas(150, 57, light)
    small = (icon[0] // 4 * 3, icon[1] // 4 * 3, [[icon[2][y * 4 // 3][x * 4 // 3] for x in range(48)] for y in range(48)])
    paste(header, small, 150 - 48 - 6, 4)
    write_bmp(HERE / 'header.bmp', header)
    fix_ico(HERE.parent / 'icons' / 'icon.ico')
    print('готово:', HERE / 'sidebar.bmp', HERE / 'header.bmp')


if __name__ == '__main__':
    main()
