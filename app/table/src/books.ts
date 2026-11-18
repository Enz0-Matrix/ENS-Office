// SPDX-License-Identifier: AGPL-3.0-or-later
// Открытые книги ENS-Table видят друг друга (ADR-012): каждое окно — своя книга, а общий
// канал BroadcastChannel связывает окна одной программы (одного источника страниц) —
// в Tauri, в окне Chromium на РЕД ОС 7 и во вкладках браузера одинаково.
// По каналу ходят только запросы «кто открыт» и листы, которые пользователь сам переносит.

export interface BookInfo { id: string; title: string; sheets: string[] }

type Msg =
  | { t: 'who'; from: string }
  | { t: 'here'; to: string; book: BookInfo }
  | { t: 'sheets'; to: string; from: string; req: string; pkgs: Uint8Array[]; before: number }
  | { t: 'sheets-done'; to: string; req: string; error?: string; names?: string[] }
  | { t: 'want'; to: string; from: string; req: string }
  | { t: 'values?'; to: string; from: string; req: string; sheet: string; rect: [number, number, number, number] }
  | { t: 'values'; to: string; req: string; bytes: Uint8Array | null };

/** Этой книге: что ответить на «кто открыт» и как принять листы. */
export interface BookHost {
  info(): Omit<BookInfo, 'id'>;
  /** Вставить листы (байты движка) перед листом `before` (-1 — в конец); вернуть имена новых листов. */
  takeSheets(pkgs: Uint8Array[], before: number): string[];
  /** Значения прямоугольника листа (для «Вставить связь» в другой книге); null — нет листа. */
  rangeValues(sheet: string, rect: [number, number, number, number]): Uint8Array | null;
}

export const bookId = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
const channel: BroadcastChannel | null = (() => {
  try { return new BroadcastChannel('ens-table-books'); } catch { return null; }
})();

let host: BookHost | null = null;
const waiters = new Map<string, (m: Msg) => void>();
/** Листы, ждущие новую книгу (она откроется и попросит их). */
const outbox = new Map<string, { pkgs: Uint8Array[]; done: (r: { error?: string; names?: string[] }) => void }>();

const post = (m: Msg) => channel?.postMessage(m);

channel?.addEventListener('message', (e: MessageEvent<Msg>) => {
  const m = e.data;
  if (m.t === 'who') {
    if (host && m.from !== bookId) post({ t: 'here', to: m.from, book: { id: bookId, ...host.info() } });
    return;
  }
  if (m.to !== bookId) return;
  if (m.t === 'here' || m.t === 'sheets-done' || m.t === 'values') { waiters.get(m.t === 'here' ? 'who' : m.req)?.(m); return; }
  if (m.t === 'values?') {
    let bytes: Uint8Array | null = null;
    try { bytes = host?.rangeValues(m.sheet, m.rect) ?? null; } catch { bytes = null; }
    post({ t: 'values', to: m.from, req: m.req, bytes });
    return;
  }
  if (m.t === 'want') {
    const o = outbox.get(m.req);
    if (!o) return;
    outbox.delete(m.req);
    void sendTo(m.from, o.pkgs, -1, m.req).then(o.done);
    return;
  }
  if (m.t === 'sheets') {
    if (!host) { post({ t: 'sheets-done', to: m.from, req: m.req, error: 'Книга ещё не готова' }); return; }
    try {
      const names = host.takeSheets(m.pkgs, m.before);
      post({ t: 'sheets-done', to: m.from, req: m.req, names });
    } catch (err) {
      post({ t: 'sheets-done', to: m.from, req: m.req, error: String(err) });
    }
  }
});

export function setBookHost(h: BookHost): void { host = h; }

export const booksAvailable = () => !!channel;

/** Другие открытые книги (ответы за 300 мс). */
export function otherBooks(): Promise<BookInfo[]> {
  if (!channel) return Promise.resolve([]);
  const found: BookInfo[] = [];
  return new Promise((resolve) => {
    waiters.set('who', (m) => { if (m.t === 'here') found.push(m.book); });
    post({ t: 'who', from: bookId });
    setTimeout(() => { waiters.delete('who'); resolve(found); }, 300);
  });
}

const newReq = () => `${bookId}:${Math.random().toString(36).slice(2, 10)}`;

/** Отправить листы в открытую книгу; ответ — её подтверждение (до 15 с). */
export function sendTo(to: string, pkgs: Uint8Array[], before: number, req = newReq()): Promise<{ error?: string; names?: string[] }> {
  return new Promise((resolve) => {
    const timer = setTimeout(() => { waiters.delete(req); resolve({ error: 'Другая книга не ответила' }); }, 15_000);
    waiters.set(req, (m) => {
      if (m.t !== 'sheets-done') return;
      clearTimeout(timer);
      waiters.delete(req);
      resolve({ error: m.error, names: m.names });
    });
    post({ t: 'sheets', to, from: bookId, req, pkgs, before });
  });
}

/** Листы для новой книги: она откроется с `?take=…&from=…` и попросит их (ждём до 60 с). */
export function holdForNewBook(pkgs: Uint8Array[]): { query: string; result: Promise<{ error?: string; names?: string[] }> } {
  const req = newReq();
  const result = new Promise<{ error?: string; names?: string[] }>((resolve) => {
    const timer = setTimeout(() => { outbox.delete(req); resolve({ error: 'Новая книга не открылась' }); }, 60_000);
    outbox.set(req, { pkgs, done: (r) => { clearTimeout(timer); resolve(r); } });
  });
  return { query: `new=1&take=${encodeURIComponent(req)}&from=${encodeURIComponent(bookId)}`, result };
}

/** Новая книга, открытая для переноса листов: просит их у книги-источника. */
export function claimSheets(): void {
  const params = new URLSearchParams(location.search);
  const req = params.get('take'), from = params.get('from');
  if (req && from) post({ t: 'want', to: from, from: bookId, req });
}

/** Свежие значения ячеек открытой книги `book` (до 2 с); null — книга закрыта или не ответила. */
export function askValues(book: string, sheet: string, rect: [number, number, number, number]): Promise<Uint8Array | null> {
  if (!channel) return Promise.resolve(null);
  const req = newReq();
  return new Promise((resolve) => {
    const timer = setTimeout(() => { waiters.delete(req); resolve(null); }, 2000);
    waiters.set(req, (m) => {
      if (m.t !== 'values') return;
      clearTimeout(timer);
      waiters.delete(req);
      resolve(m.bytes);
    });
    post({ t: 'values?', to: book, from: bookId, req, sheet, rect });
  });
}
