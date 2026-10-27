// Встроенные свободные шрифты (OFL) с теми же размерами символов, что у шрифтов
// Office. В файле остаётся привычное имя (Calibri, Arial…), а если такого шрифта
// в системе нет (РЕД ОС, Linux), рисуется аналог — вёрстка не «плывёт».

import '@fontsource/carlito/400.css';
import '@fontsource/carlito/700.css';
import '@fontsource/carlito/400-italic.css';
import '@fontsource/carlito/700-italic.css';
import '@fontsource/caladea/400.css';
import '@fontsource/caladea/700.css';
import '@fontsource/arimo/400.css';
import '@fontsource/arimo/700.css';
import '@fontsource/tinos/400.css';
import '@fontsource/tinos/700.css';
import '@fontsource/cousine/400.css';
import '@fontsource/cousine/700.css';
import '@fontsource/roboto-condensed/400.css';
import '@fontsource/roboto-condensed/700.css';
import condensedCyrillic from '@fontsource/roboto-condensed/files/roboto-condensed-cyrillic-300-normal.woff2?url';
import condensedLatin from '@fontsource/roboto-condensed/files/roboto-condensed-latin-300-normal.woff2?url';

/** Инженерный шрифт: узкий, светлый, с кириллицей — открытый аналог Bahnschrift Light (OFL). */
export const ENGINEERING_FONT = 'Roboto Condensed Light';
for (const [url, range] of [
  [condensedCyrillic, 'U+0301,U+0400-045F,U+0490-0491,U+04B0-04B1,U+2116'],
  [condensedLatin, 'U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD'],
]) {
  const face = new FontFace(ENGINEERING_FONT, `url(${url}) format("woff2")`, { weight: '400', unicodeRange: range });
  document.fonts.add(face);
  void face.load();
}

/** Шрифт Office → встроенный аналог. */
const SUBSTITUTES: Record<string, string> = {
  calibri: 'Carlito',
  cambria: 'Caladea',
  arial: 'Arimo',
  helvetica: 'Arimo',
  'liberation sans': 'Arimo',
  'times new roman': 'Tinos',
  'liberation serif': 'Tinos',
  'courier new': 'Cousine',
  'liberation mono': 'Cousine',
  bahnschrift: ENGINEERING_FONT,
  'bahnschrift light': ENGINEERING_FONT,
};

export const OFFICE_FONTS = ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New', 'Verdana', 'Tahoma', 'Segoe UI', 'Inter', ENGINEERING_FONT, 'Roboto Condensed'];

/** Список шрифтов для CSS/Canvas: сам шрифт, его аналог, затем запасные. */
export function fontStack(name?: string): string {
  const parts: string[] = [];
  if (name) {
    parts.push(`"${name}"`);
    const sub = SUBSTITUTES[name.toLowerCase()];
    if (sub) parts.push(`"${sub}"`);
  }
  parts.push('"Carlito"', '"Arimo"', 'sans-serif');
  return parts.join(', ');
}

/** Загрузить встроенные шрифты до первой отрисовки Canvas. */
export async function loadFonts(): Promise<void> {
  const faces = ['Carlito', 'Caladea', 'Arimo', 'Tinos', 'Cousine', 'Roboto Condensed', ENGINEERING_FONT];
  const loads: Promise<unknown>[] = [];
  for (const face of faces) {
    for (const weight of ['400', '700']) {
      loads.push(document.fonts.load(`${weight} 15px "${face}"`, 'Ая'));
    }
  }
  loads.push(document.fonts.load('italic 400 15px "Carlito"', 'Ая'));
  await Promise.allSettled(loads);
}
