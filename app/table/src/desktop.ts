// Документ в настольной версии: файл на диске, признак изменений, журнал для
// восстановления после сбоя (§7), вопрос «Сохранить?» при закрытии.
// В браузере те же функции работают без файлов: сохранение — скачиванием.

import { fromBytes, openXlsx, type Model } from './engine';
import { ask, inspectFile, isDesktop, startupFile, openWindow, pickFolder, pickOpen, pickSave, readFile, recovery, saveFile, type RecoverableDoc } from './platform';
import { localeId, saveSettings, settings } from './settings';

export interface Doc {
  /** Идентификатор для папки восстановления: латиница, цифры, дефис. */
  id: string;
  path: string;
  title: string;
  modifiedMs: number;
  dirty: boolean;
  /** Книга восстановлена после сбоя: временные файлы удаляются, когда её сохранят. */
  recoveredFrom?: { session: string; id: string };
}

export interface Host {
  model(): Model;
  setModel(model: Model, doc: Doc): void;
  doc(): Doc;
  onDocChanged(): void;
  toast(text: string): void;
}

const newId = () => `doc-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;

export function newDoc(title = 'Книга1'): Doc {
  return { id: newId(), path: '', title, modifiedMs: 0, dirty: false };
}

let host: Host;
let emptyQueue = 0;
let lastChange = 0;
let lastSnapshot = 0;
let changedSinceSnapshot = false;

const baseName = (path: string) => path.split(/[\\/]/).pop()!.replace(/\.(xlsx|xlsm)$/i, '');

/** Начать вести книгу: метаданные и первый снимок. */
async function track(): Promise<void> {
  const doc = host.doc();
  host.model().flushSendQueue();
  if (!isDesktop) return;
  await recovery.begin(doc.id, {
    source_path: doc.path, title: doc.title, source_modified_ms: doc.modifiedMs, snapshot_ms: 0, journal_entries: 0,
  });
  await snapshot();
}

async function snapshot(): Promise<void> {
  const m = host.model();
  const pending = m.flushSendQueue();
  if (pending.length > emptyQueue) markChanged();
  if (!isDesktop) return;
  await recovery.snapshot(host.doc().id, m.toBytes());
  lastSnapshot = Date.now();
  changedSinceSnapshot = false;
}

function markChanged(): void {
  lastChange = Date.now();
  changedSinceSnapshot = true;
  const doc = host.doc();
  if (!doc.dirty) {
    doc.dirty = true;
    host.onDocChanged();
  }
}

/** Раз в секунду: правки — в журнал; снимок — по расписанию и после простоя. */
async function tick(): Promise<void> {
  const m = host.model();
  const diffs = m.flushSendQueue();
  if (diffs.length > emptyQueue) {
    markChanged();
    if (isDesktop) await recovery.append(host.doc().id, diffs).catch((e) => host.toast(String(e)));
  }
  if (!isDesktop || !changedSinceSnapshot) return;
  const now = Date.now();
  const due = now - lastSnapshot >= settings.autosaveMinutes * 60_000;
  const idle = now - lastChange >= 30_000;
  if (due || idle) await snapshot().catch((e) => host.toast(String(e)));
}

export async function newWindow(): Promise<void> {
  await openWindow('new=1', 'Книга');
}

export async function closeBook(): Promise<void> {
  if (!isDesktop) return;
  const { getCurrentWindow } = await import('@tauri-apps/api/window');
  await getCurrentWindow().close(); // спросит «Сохранить изменения?»
}

export async function recoveryFolder(): Promise<string> {
  return isDesktop ? recovery.root() : '';
}

export async function chooseRecoveryFolder(): Promise<void> {
  const path = await pickFolder('Папка для временных файлов восстановления');
  if (!path) return;
  try {
    await recovery.setRoot(path);
    settings.recoveryRoot = path;
    saveSettings();
    await track();
  } catch (e) {
    host.toast(`Папка недоступна для записи: ${e}. Временные файлы остаются в прежней папке.`);
  }
}

export async function start(h: Host): Promise<void> {
  host = h;
  if (isDesktop && settings.recoveryRoot) {
    // Своя (например, сетевая) папка; если недоступна — остаёмся в профиле.
    await recovery.setRoot(settings.recoveryRoot).catch(() =>
      h.toast(`Папка восстановления «${settings.recoveryRoot}» недоступна — временные файлы пишутся в профиль пользователя.`));
  }
  emptyQueue = h.model().flushSendQueue().length;
  setInterval(() => void tick(), 1000);
  if (isDesktop) {
    setInterval(() => void recovery.heartbeat().catch(() => {}), 30_000);
    await setupClose();
  }
  const params = new URLSearchParams(location.search);
  const open = params.get('open');
  const recover = params.get('recover');
  if (open) await openPath(open);
  else if (recover) {
    const [session, id] = recover.split('/');
    await restore(session, id, params.get('title') ?? 'Книга');
  } else if (params.get('new')) {
    await track(); // новое окно: панель восстановления уже показана в первом
  } else {
    await track();
    const file = isDesktop ? await startupFile().catch(() => null) : null;
    if (file) await openPath(file);
    if (isDesktop) await offerRecovery();
  }
}

// --- Открытие ---

const KIND_NAMES: Record<string, string> = {
  docx: 'документ Word', pptx: 'презентация PowerPoint', odt: 'текстовый документ OpenDocument',
  ods: 'таблица OpenDocument (.ods)', 'xls-or-doc': 'файл старого формата (.xls или .doc)', csv: 'текст с разделителями (CSV)',
};

export async function openDialog(): Promise<void> {
  if (!isDesktop) {
    document.getElementById('file-input')!.click();
    return;
  }
  const path = await pickOpen();
  if (path) await openPath(path);
}

export async function openPath(path: string): Promise<void> {
  const doc = host.doc();
  if (doc.dirty || doc.path) {
    // Текущая книга не пустая — открываем рядом, в новом окне, как Excel.
    await openWindow(`open=${encodeURIComponent(path)}`, baseName(path));
    return;
  }
  try {
    const info = await inspectFile(path);
    if (info.kind !== 'xlsx' && info.kind !== 'xlsm') {
      const what = KIND_NAMES[info.kind] ?? 'файл неизвестного формата';
      host.toast(`«${info.name}» — это ${what}, а не книга Excel. ENS-Table открывает .xlsx и .xlsm` +
        (info.kind === 'docx' ? '; документы Word будет открывать ENS-Tekst.' : '.'));
      return;
    }
    const bytes = await readFile(path);
    const model = openXlsx(bytes, info.name, localeId());
    if (isDesktop) await recovery.forget(doc.id).catch(() => {});
    host.setModel(model, { id: newId(), path, title: baseName(path), modifiedMs: info.modified_ms, dirty: false });
    await track();
    if (info.readonly) host.toast('Файл только для чтения: изменения можно сохранить под другим именем.');
  } catch (e) {
    host.toast(String(e));
  }
}

// --- Сохранение ---

export async function save(): Promise<boolean> {
  const doc = host.doc();
  if (!isDesktop) return download();
  if (!doc.path) return saveAs();
  return writeTo(doc.path, false);
}

export async function saveAs(): Promise<boolean> {
  const doc = host.doc();
  if (!isDesktop) return download();
  const path = await pickSave(`${doc.title.replace(/ \(восстановлено\)$/, '')}.xlsx`);
  if (!path) return false;
  return writeTo(path.endsWith('.xlsx') || path.endsWith('.xlsm') ? path : `${path}.xlsx`, true);
}

async function writeTo(path: string, isNew: boolean): Promise<boolean> {
  const doc = host.doc();
  const bytes = host.model().toXlsx();
  const expected = isNew ? 0 : doc.modifiedMs;
  try {
    doc.modifiedMs = await saveFile(path, bytes, expected);
  } catch (e) {
    if (String(e).includes('CHANGED_ON_DISK')) {
      const overwrite = await ask(
        `Файл «${baseName(path)}» изменён другой программой или пользователем после того, как вы его открыли.\n\n` +
          'Перезаписать его вашей версией? Чтобы сохранить обе версии, выберите «Отмена» и затем «Сохранить как».',
        'Перезаписать', 'Отмена');
      if (!overwrite) return false;
      try {
        doc.modifiedMs = await saveFile(path, bytes, 0, true);
      } catch (e2) {
        host.toast(String(e2));
        return false;
      }
    } else {
      host.toast(String(e));
      return false;
    }
  }
  doc.path = path;
  doc.title = baseName(path);
  doc.dirty = false;
  rememberRecent(path, doc.title);
  if (doc.recoveredFrom) {
    await recovery.discard(doc.recoveredFrom.session, doc.recoveredFrom.id).catch(() => {});
    doc.recoveredFrom = undefined;
  }
  host.onDocChanged();
  // Сохранённая версия — новая точка отсчёта для восстановления.
  await recovery.begin(doc.id, {
    source_path: doc.path, title: doc.title, source_modified_ms: doc.modifiedMs, snapshot_ms: 0, journal_entries: 0,
  });
  await snapshot();
  return true;
}

/** Список «Последние» в меню «Файл». */
export function rememberRecent(path: string, title: string): void {
  settings.recent = [{ path, title }, ...settings.recent.filter((r) => r.path !== path)].slice(0, 12);
  saveSettings();
}

function download(): boolean {
  const doc = host.doc();
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([new Uint8Array(host.model().toXlsx())], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  }));
  a.download = `${doc.title}.xlsx`;
  a.click();
  URL.revokeObjectURL(a.href);
  doc.dirty = false;
  host.onDocChanged();
  return true;
}

// --- Восстановление после сбоя ---

async function restore(session: string, id: string, title: string): Promise<void> {
  try {
    const data = await recovery.load(session, id);
    const model = fromBytes(data.snapshot);
    for (const frame of data.journal) model.applyExternalDiffs(frame);
    model.evaluate();
    const list = await recovery.list();
    const meta = list.find((d) => d.session === session && d.id === id)?.meta;
    host.setModel(model, {
      id: newId(), path: meta?.source_path ?? '', title: `${meta?.title ?? title} (восстановлено)`,
      modifiedMs: meta?.source_modified_ms ?? 0, dirty: true, recoveredFrom: { session, id },
    });
    await track();
    markChanged();
  } catch (e) {
    host.toast(`Не удалось восстановить книгу: ${e}`);
  }
}

function fmtTime(ms: number): string {
  return new Date(ms).toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' });
}

async function offerRecovery(): Promise<void> {
  const list = await recovery.list().catch(() => [] as RecoverableDoc[]);
  if (!list.length) return;
  const panel = document.getElementById('recovery')!;
  const render = (docs: RecoverableDoc[]) => {
    if (!docs.length) { panel.hidden = true; return; }
    panel.innerHTML = `
      <div class="modal">
        <h2>Восстановление документов</h2>
        <p>Программа была закрыта аварийно. Эти книги можно восстановить на момент сбоя.
           Исходные файлы не изменялись — откройте обе версии и сохраните нужную.</p>
        <div class="rec-list"></div>
        <div class="rec-actions">
          <button data-act="later">Решить позже</button>
          <button data-act="discard-all" class="danger">Удалить все временные файлы</button>
        </div>
      </div>`;
    const host = panel.querySelector('.rec-list')!;
    for (const d of docs) {
      const row = document.createElement('div');
      row.className = 'rec-item';
      row.innerHTML = `
        <div class="rec-title">${escapeHtml(d.meta.title)}</div>
        <div class="rec-meta">${d.meta.source_path ? escapeHtml(d.meta.source_path) : 'новая книга, не сохранялась'}<br>
          сбой: ${fmtTime(d.crashed_ms)}, последний снимок: ${d.meta.snapshot_ms ? fmtTime(d.meta.snapshot_ms) : '—'}
          ${d.source_changed ? '<br><b>Исходный файл менялся после открытия.</b>' : ''}</div>
        <div class="rec-buttons">
          <button data-act="restore">Открыть восстановленную</button>
          ${d.source_exists ? '<button data-act="original">Открыть исходную рядом</button>' : ''}
          <button data-act="discard" class="danger">Удалить временные файлы</button>
        </div>`;
      row.addEventListener('click', async (e) => {
        const act = (e.target as HTMLElement).dataset.act;
        if (act === 'restore') {
          const doc = window.__ensDoc?.();
          if (doc && !doc.dirty && !doc.path) await restore(d.session, d.id, d.meta.title);
          else await openWindow(`recover=${d.session}/${d.id}&title=${encodeURIComponent(d.meta.title)}`, d.meta.title);
          render(docs.filter((x) => x !== d));
        } else if (act === 'original') {
          await openWindow(`open=${encodeURIComponent(d.meta.source_path)}`, `${d.meta.title} (исходная)`);
        } else if (act === 'discard') {
          if (await ask(`Удалить временные файлы книги «${d.meta.title}»? Останется только версия, сохранённая вами.`, 'Удалить', 'Отмена')) {
            await recovery.discard(d.session, d.id);
            render(docs.filter((x) => x !== d));
          }
        }
      });
      host.append(row);
    }
    panel.querySelector('[data-act="later"]')!.addEventListener('click', () => (panel.hidden = true));
    panel.querySelector('[data-act="discard-all"]')!.addEventListener('click', async () => {
      if (await ask('Удалить временные файлы всех книг из списка? Останутся только версии, сохранённые вами.', 'Удалить все', 'Отмена')) {
        for (const d of docs) await recovery.discard(d.session, d.id).catch(() => {});
        panel.hidden = true;
      }
    });
    panel.hidden = false;
  };
  render(list);
}

const escapeHtml = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);

declare global {
  interface Window { __ensDoc?: () => Doc }
}

// --- Закрытие окна ---

async function setupClose(): Promise<void> {
  const { getCurrentWindow } = await import('@tauri-apps/api/window');
  const { getAllWebviewWindows } = await import('@tauri-apps/api/webviewWindow');
  const win = getCurrentWindow();
  await win.onCloseRequested(async (event) => {
    event.preventDefault();
    await tick();
    const doc = host.doc();
    if (doc.dirty) {
      const choice = await askSave(doc.title);
      if (choice === 'cancel') return;
      if (choice === 'save' && !(await save())) return;
    }
    await recovery.forget(doc.id).catch(() => {});
    if ((await getAllWebviewWindows()).length <= 1) await recovery.endSession().catch(() => {});
    await win.destroy();
  });
}

/** «Сохранить изменения?» — три кнопки, как в Excel. */
function askSave(title: string): Promise<'save' | 'discard' | 'cancel'> {
  const panel = document.getElementById('recovery')!;
  panel.innerHTML = `
    <div class="modal small">
      <h2>Сохранить изменения в «${escapeHtml(title)}»?</h2>
      <p>Если не сохранить, изменения будут потеряны.</p>
      <div class="rec-actions">
        <button data-act="save" class="primary">Сохранить</button>
        <button data-act="discard">Не сохранять</button>
        <button data-act="cancel">Отмена</button>
      </div>
    </div>`;
  panel.hidden = false;
  return new Promise((resolve) => {
    panel.querySelectorAll<HTMLButtonElement>('button').forEach((b) =>
      b.addEventListener('click', () => {
        panel.hidden = true;
        resolve(b.dataset.act as 'save' | 'discard' | 'cancel');
      }));
  });
}
