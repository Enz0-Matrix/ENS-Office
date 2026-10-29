// Платформа:
//  * настольная оболочка Tauri (Windows, РЕД ОС 8+): настоящие файлы,
//    восстановление после сбоя;
//  * версия для РЕД ОС 7 (app/table/redos7): окно Chromium в режиме приложения
//    и программа-помощник ens-table на 127.0.0.1 с теми же возможностями;
//  * обычный браузер: скачивание копии, без восстановления.

import { invoke } from '@tauri-apps/api/core';

export const isTauri = '__TAURI_INTERNALS__' in window;

/** Ключ помощника РЕД ОС 7: приходит в адресе окна (?t=…), без него помощник
 *  не отвечает — чужие страницы в браузере не доберутся до файлов. */
const localToken: string | null = (() => {
  if (isTauri || location.hostname !== '127.0.0.1') return null;
  const fromUrl = new URLSearchParams(location.search).get('t');
  try {
    if (fromUrl) sessionStorage.setItem('ens-token', fromUrl);
    return fromUrl ?? sessionStorage.getItem('ens-token');
  } catch {
    return fromUrl;
  }
})();

export const isLocalApp = !!localToken;
export const isDesktop = isTauri || isLocalApp;

/** Команда оболочки: в Tauri — invoke, в версии для РЕД ОС 7 — запрос к помощнику. */
async function call<T>(cmd: string, args?: Record<string, unknown> | Uint8Array, headers?: Record<string, string>): Promise<T> {
  if (isTauri) return invoke<T>(cmd, args as never, headers ? { headers } : undefined);
  const raw = args instanceof Uint8Array;
  const res = await fetch(`/api/${cmd}`, {
    method: 'POST',
    body: raw ? new Blob([args as Uint8Array<ArrayBuffer>]) : JSON.stringify(args ?? {}),
    headers: { 'x-ens-token': localToken ?? '', 'content-type': raw ? 'application/octet-stream' : 'application/json', ...(headers ?? {}) },
  });
  if (!res.ok) throw new Error(await res.text());
  if ((res.headers.get('content-type') ?? '').includes('octet-stream')) return (await res.arrayBuffer()) as T;
  return (await res.json()) as T;
}

export interface FileInfo {
  kind: string;
  size: number;
  modified_ms: number;
  readonly: boolean;
  name: string;
}

export interface DocMeta {
  source_path: string;
  title: string;
  source_modified_ms: number;
  snapshot_ms: number;
  journal_entries: number;
  /** Формат снимка (версия программы); у другой версии — восстанавливаем из xlsx. */
  format?: string;
  xlsx_ms?: number;
}

export interface RecoverableDoc {
  session: string;
  id: string;
  meta: DocMeta;
  source_exists: boolean;
  source_changed: boolean;
  crashed_ms: number;
}

const XLSX_FILTER = [
  { name: 'Книги Excel', extensions: ['xlsx', 'xlsm'] },
  { name: 'Все файлы', extensions: ['*'] },
];

/** Названия окон и типы файлов приложения. ENS-Text задаёт свои через setDialogProfile. */
let profile = { app: 'ENS-Table', open: 'Открыть книгу', filters: XLSX_FILTER };
export function setDialogProfile(p: { app: string; open: string; filters: { name: string; extensions: string[] }[] }): void {
  profile = p;
}

/** Окно выбора книги; dir — папка, с которой оно начинается. */
export async function pickOpen(dir?: string): Promise<string | null> {
  if (isLocalApp) return call<string | null>('pick_open', { title: profile.open, dir: dir ?? null });
  const { open } = await import('@tauri-apps/plugin-dialog');
  const path = await open({ multiple: false, directory: false, filters: profile.filters, title: profile.open, defaultPath: dir });
  return typeof path === 'string' ? path : null;
}

export async function pickSave(defaultName: string): Promise<string | null> {
  if (isLocalApp) return call<string | null>('pick_save', { title: 'Сохранить как', name: defaultName });
  const { save } = await import('@tauri-apps/plugin-dialog');
  return save({ defaultPath: defaultName, filters: profile.filters.slice(0, 1), title: 'Сохранить как' });
}

export async function ask(message: string, ok: string, cancel: string): Promise<boolean> {
  if (isLocalApp) return confirm(message);
  const { ask: dialogAsk } = await import('@tauri-apps/plugin-dialog');
  return dialogAsk(message, { title: profile.app, okLabel: ok, cancelLabel: cancel, kind: 'warning' });
}

/** Файл, с которым запустили программу («Открыть с помощью»). */
/** Имя пользователя в системе (для журнала изменений). */
/** Открыть адрес ссылки (сайт, почта, файл, папка) в программе по умолчанию. */
export const openTarget = (target: string) => call<void>('open_target', { target });
export const osUserName = () => (isDesktop ? call<string>('user_name').catch(() => '') : Promise.resolve(''));
export const startupFile = () => (isTauri ? invoke<string | null>('startup_file') : Promise.resolve(null));

/** Копии на диске (только ENS-Text): предыдущие версии файлов и документы, закрытые без сохранения. */
export interface HistoryEntry { kind: 'backup' | 'discarded'; name: string; path: string; ms: number; size: number }
export const historyStore = (kind: HistoryEntry['kind'], name: string, bytes: Uint8Array) =>
  call<string>('history_store', bytes, { kind: encodeURIComponent(kind), name: encodeURIComponent(name) });
export const historyList = () => call<HistoryEntry[]>('history_list');

export const inspectFile = (path: string) => call<FileInfo>('inspect_file', { path });

export async function readFile(path: string): Promise<Uint8Array> {
  return new Uint8Array(await call<ArrayBuffer>('read_file', { path }));
}

/** Атомарное сохранение. Ошибка 'CHANGED_ON_DISK' — файл изменили снаружи. */
export function saveFile(path: string, bytes: Uint8Array, expectedModified: number, force = false): Promise<number> {
  return call<number>('save_file', bytes, { path: encodeURIComponent(path), 'expected-modified': String(expectedModified), force: force ? '1' : '0' });
}

// --- Восстановление после сбоя (§7) ---

export const recovery = {
  heartbeat: () => call('recovery_heartbeat'),
  begin: (id: string, meta: DocMeta) => call('recovery_begin', { id, meta }),
  snapshot: (id: string, bytes: Uint8Array) => call('recovery_snapshot', bytes, { id }),
  snapshotXlsx: (id: string, bytes: Uint8Array) => call('recovery_snapshot_xlsx', bytes, { id }),
  loadXlsx: async (session: string, id: string) => new Uint8Array(await call<ArrayBuffer>('recovery_load_snapshot_xlsx', { session, id })),
  append: (id: string, bytes: Uint8Array) => call('recovery_append', bytes, { id }),
  forget: (id: string) => call('recovery_forget', { id }),
  endSession: () => call('recovery_end_session'),
  list: () => call<RecoverableDoc[]>('recovery_list'),
  root: () => call<string>('recovery_root'),
  setRoot: (path: string) => call('recovery_set_root', { path }),
  discard: (session: string, id: string) => call('recovery_discard', { session, id }),
  async load(session: string, id: string): Promise<{ snapshot: Uint8Array; journal: Uint8Array[] }> {
    const snapshot = new Uint8Array(await call<ArrayBuffer>('recovery_load_snapshot', { session, id }));
    const raw = new Uint8Array(await call<ArrayBuffer>('recovery_load_journal', { session, id }));
    const journal: Uint8Array[] = [];
    const view = new DataView(raw.buffer, raw.byteOffset, raw.byteLength);
    for (let at = 0; at + 4 <= raw.length; ) {
      const len = view.getUint32(at, true);
      journal.push(raw.slice(at + 4, at + 4 + len));
      at += 4 + len;
    }
    return { snapshot, journal };
  },
};

/** Буфер обмена: в настольной версии — системный (плагин Tauri), в браузере —
 *  Clipboard API, а если он недоступен — внутренний буфер программы. */
let internalClipboard = '';

export async function copyText(text: string): Promise<void> {
  internalClipboard = text;
  if (isTauri) {
    const { writeText } = await import('@tauri-apps/plugin-clipboard-manager');
    await writeText(text);
    return;
  }
  await navigator.clipboard.writeText(text).catch(() => {});
}

export async function readText(): Promise<string> {
  if (isTauri) {
    const { readText: read } = await import('@tauri-apps/plugin-clipboard-manager');
    return (await read()) ?? internalClipboard;
  }
  try {
    return await navigator.clipboard.readText();
  } catch {
    if (internalClipboard) return internalClipboard;
    throw new Error('буфер обмена недоступен');
  }
}

export async function pickFolder(title: string): Promise<string | null> {
  if (isLocalApp) return call<string | null>('pick_folder', { title });
  const { open } = await import('@tauri-apps/plugin-dialog');
  const path = await open({ directory: true, multiple: false, title });
  return typeof path === 'string' ? path : null;
}

/** Открыть ещё одно окно ENS-Table (например, исходную книгу рядом с восстановленной). */
/** Версия для РЕД ОС 7: окно закрыли — помощнику сообщают без ожидания ответа. */
export function beaconClose(docId: string, clean: boolean): void {
  if (!isLocalApp) return;
  const body = new Blob([JSON.stringify({ id: docId, clean })], { type: 'text/plain' });
  navigator.sendBeacon(`/api/window_closed?t=${encodeURIComponent(localToken ?? '')}`, body);
}

export async function openWindow(query: string, title: string): Promise<void> {
  if (isLocalApp) { await call('open_window', { query }); return; }
  const { WebviewWindow } = await import('@tauri-apps/api/webviewWindow');
  const label = `doc-${Date.now()}`;
  new WebviewWindow(label, { url: `index.html?${query}`, title, width: 1200, height: 800, minWidth: 640, minHeight: 400, decorations: false, shadow: true });
}
