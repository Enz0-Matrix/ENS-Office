// Платформа: настольная оболочка Tauri (настоящие файлы, восстановление после
// сбоя) или обычный браузер (скачивание копии, без восстановления).

import { invoke } from '@tauri-apps/api/core';

export const isDesktop = '__TAURI_INTERNALS__' in window;

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

export async function pickOpen(): Promise<string | null> {
  const { open } = await import('@tauri-apps/plugin-dialog');
  const path = await open({ multiple: false, directory: false, filters: XLSX_FILTER, title: 'Открыть книгу' });
  return typeof path === 'string' ? path : null;
}

export async function pickSave(defaultName: string): Promise<string | null> {
  const { save } = await import('@tauri-apps/plugin-dialog');
  return save({ defaultPath: defaultName, filters: XLSX_FILTER.slice(0, 1), title: 'Сохранить как' });
}

export async function ask(message: string, ok: string, cancel: string): Promise<boolean> {
  const { ask: dialogAsk } = await import('@tauri-apps/plugin-dialog');
  return dialogAsk(message, { title: 'ENS-Table', okLabel: ok, cancelLabel: cancel, kind: 'warning' });
}

/** Файл, с которым запустили программу («Открыть с помощью»). */
export const startupFile = () => invoke<string | null>('startup_file');

export const inspectFile = (path: string) => invoke<FileInfo>('inspect_file', { path });

export async function readFile(path: string): Promise<Uint8Array> {
  return new Uint8Array(await invoke<ArrayBuffer>('read_file', { path }));
}

/** Атомарное сохранение. Ошибка 'CHANGED_ON_DISK' — файл изменили снаружи. */
export function saveFile(path: string, bytes: Uint8Array, expectedModified: number, force = false): Promise<number> {
  return invoke<number>('save_file', bytes, {
    headers: { path: encodeURIComponent(path), 'expected-modified': String(expectedModified), force: force ? '1' : '0' },
  });
}

// --- Восстановление после сбоя (§7) ---

export const recovery = {
  heartbeat: () => invoke('recovery_heartbeat'),
  begin: (id: string, meta: DocMeta) => invoke('recovery_begin', { id, meta }),
  snapshot: (id: string, bytes: Uint8Array) => invoke('recovery_snapshot', bytes, { headers: { id } }),
  append: (id: string, bytes: Uint8Array) => invoke('recovery_append', bytes, { headers: { id } }),
  forget: (id: string) => invoke('recovery_forget', { id }),
  endSession: () => invoke('recovery_end_session'),
  list: () => invoke<RecoverableDoc[]>('recovery_list'),
  root: () => invoke<string>('recovery_root'),
  setRoot: (path: string) => invoke('recovery_set_root', { path }),
  discard: (session: string, id: string) => invoke('recovery_discard', { session, id }),
  async load(session: string, id: string): Promise<{ snapshot: Uint8Array; journal: Uint8Array[] }> {
    const snapshot = new Uint8Array(await invoke<ArrayBuffer>('recovery_load_snapshot', { session, id }));
    const raw = new Uint8Array(await invoke<ArrayBuffer>('recovery_load_journal', { session, id }));
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
  if (isDesktop) {
    const { writeText } = await import('@tauri-apps/plugin-clipboard-manager');
    await writeText(text);
    return;
  }
  await navigator.clipboard.writeText(text).catch(() => {});
}

export async function readText(): Promise<string> {
  if (isDesktop) {
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
  const { open } = await import('@tauri-apps/plugin-dialog');
  const path = await open({ directory: true, multiple: false, title });
  return typeof path === 'string' ? path : null;
}

/** Открыть ещё одно окно ENS-Table (например, исходную книгу рядом с восстановленной). */
export async function openWindow(query: string, title: string): Promise<void> {
  const { WebviewWindow } = await import('@tauri-apps/api/webviewWindow');
  const label = `doc-${Date.now()}`;
  new WebviewWindow(label, { url: `index.html?${query}`, title, width: 1200, height: 800 });
}
