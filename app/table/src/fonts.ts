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
};

export const OFFICE_FONTS = ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New', 'Verdana', 'Tahoma', 'Segoe UI', 'Inter'];

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
  const faces = ['Carlito', 'Caladea', 'Arimo', 'Tinos', 'Cousine'];
  const loads: Promise<unknown>[] = [];
  for (const face of faces) {
    for (const weight of ['400', '700']) {
      loads.push(document.fonts.load(`${weight} 15px "${face}"`, 'Ая'));
    }
  }
  loads.push(document.fonts.load('italic 400 15px "Carlito"', 'Ая'));
  await Promise.allSettled(loads);
}
