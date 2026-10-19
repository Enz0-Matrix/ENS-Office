// Реестр команд (ADR-003): лента, меню, горячие клавиши и поиск вызывают
// одну и ту же команду по её command_id.

export interface Command {
  id: string;
  label: string;
  /** Синонимы и английские названия для поиска. */
  keywords?: string[];
  /** Сочетание в виде «Ctrl+Shift+KeyL»: по физической клавише (event.code),
   *  поэтому работает в любой раскладке — и в русской тоже. */
  keys?: string[];
  /** arg — значение из выпадающего списка ленты (например, числовой формат). */
  run: (arg?: string) => void;
  /** Включено ли (для переключателей — нажато ли). */
  checked?: () => boolean;
  /** Сочетание обрабатывает сам браузер (Ctrl+C/X/V — события буфера обмена);
   *  keys тогда только для подсказок. */
  nativeKeys?: boolean;
}

const commands = new Map<string, Command>();

export function register(...list: Command[]): void {
  for (const c of list) commands.set(c.id, c);
}

export function command(id: string): Command | undefined {
  return commands.get(id);
}

export function run(id: string, arg?: string): void {
  commands.get(id)?.run(arg);
}

export function allCommands(): Command[] {
  return [...commands.values()];
}

/** «Ctrl+Shift+KeyL» для события клавиатуры. */
export function keyOf(e: KeyboardEvent): string {
  const parts: string[] = [];
  if (e.ctrlKey || e.metaKey) parts.push('Ctrl');
  if (e.altKey) parts.push('Alt');
  if (e.shiftKey) parts.push('Shift');
  parts.push(e.code);
  return parts.join('+');
}

/** Человеческая запись сочетания: «Ctrl+Shift+L». */
export function keyLabel(key: string): string {
  return key
    .replace(/Key([A-Z])/g, '$1')
    .replace(/Digit(\d)/g, '$1')
    .replace('BracketLeft', '[')
    .replace('BracketRight', ']')
    .replace('ArrowLeft', '←')
    .replace('ArrowRight', '→');
}

/** Выполняет команду, привязанную к сочетанию; true, если такая нашлась. */
export function handleKey(e: KeyboardEvent): boolean {
  const key = keyOf(e);
  for (const c of commands.values()) {
    if (!c.nativeKeys && c.keys?.includes(key)) {
      e.preventDefault();
      c.run();
      return true;
    }
  }
  return false;
}

// --- Поиск команд и функций (строка поиска в заголовке окна) ---

function normalize(s: string): string {
  return s.toLowerCase().replace(/ё/g, 'е').replace(/[^\p{L}\p{N}]+/gu, ' ').trim();
}

/** Расстояние Левенштейна, ограниченное сверху: для опечаток в словах. */
function distance(a: string, b: string, limit: number): number {
  if (Math.abs(a.length - b.length) > limit) return limit + 1;
  let prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    for (let j = 1; j <= b.length; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = cur;
  }
  return prev[b.length];
}

/** Оценка совпадения запроса с текстом: 0 — нет совпадения. */
export function matchScore(query: string, text: string): number {
  const q = normalize(query);
  const t = normalize(text);
  if (!q) return 0;
  if (t.startsWith(q)) return 100;
  if (t.includes(q)) return 80;
  const words = t.split(' ');
  let score = 0;
  for (const qw of q.split(' ')) {
    let best = 0;
    for (const w of words) {
      if (w.startsWith(qw)) best = Math.max(best, 30);
      else if (qw.length >= 4 && distance(qw, w.slice(0, qw.length + 1), 2) <= (qw.length >= 7 ? 2 : 1)) best = Math.max(best, 15);
    }
    if (!best) return 0;
    score += best;
  }
  return score;
}
