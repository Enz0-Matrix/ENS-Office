// ENS-Tekst в настольной оболочке (Tauri: Windows, Linux) и в версии для РЕД ОС 7 (окно Chromium + помощник):
// настоящие файлы, атомарное сохранение (ADR-005), восстановление после сбоя, «Открыть с помощью»,
// вопрос «Сохранить?» при закрытии. В браузере всего этого нет: там main.ts скачивает копию.

import {
  ask, beaconClose, inspectFile, isDesktop, isTauri, openTarget, openWindow, pickOpen, pickSave, readFile, recovery, saveFile,
  setDialogProfile, startupFile, type RecoverableDoc,
} from '../../table/src/platform';
import { choose, dialog, el, toast } from './ui';
import { historyStore } from '../../table/src/platform';
import { APP_VERSION } from './version';

export { isDesktop, openTarget };

export interface Host {
  name(): string;
  isDirty(): boolean;
  /** Пустой нетронутый документ: его можно заменить, не открывая новое окно. */
  isPristine(): boolean;
  /** Собирает docx и проверяет его чтением; бросает ошибку, если файл не читается. */
  docxBytes(name: string): Uint8Array;
  /** Снимок для восстановления: документ, страница, имя (JSON). */
  snapshotJson(): string;
  /** Открывает docx в окне; возвращает, что не перенесено. info — откуда (для журнала и вопроса «что изменилось»). */
  openBytes(bytes: Uint8Array, name: string, info?: { key: string; mtime?: number }): string[];
  /** Версия и запись журнала при сохранении; null-результата нет: бросает, если файл не читается. */
  rollbackSave(): void;
  /** Файл записан на диск. */
  afterSaved(key: string): void;
  /** Копия документа для хранения на диске (с историей версий), без записи в журнал. */
  copyBytes(): Uint8Array;
  /** Открывает снимок восстановления (документ остаётся несохранённым). */
  restoreJson(json: string, title: string): void;
  /** Документ сохранён под этим именем. */
  markSaved(name: string): void;
}

const newId = () => `doc-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;

/** Файл на диске, с которым связан открытый документ. */
const file = {
  id: newId(),
  path: '',
  modifiedMs: 0,
  /** Что не перенеслось при открытии, и из какого файла: поверх него писать нельзя. */
  lossy: [] as string[],
  lossySource: '',
  recoveredFrom: undefined as { session: string; id: string } | undefined,
};

let host: Host;
let tracked = false;
let lastBackup = 0;

const baseName = (path: string) => path.split(/[\\/]/).pop()!.replace(/\.docx$/i, '');
const fmtTime = (ms: number) => new Date(ms).toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' });

const KIND_NAMES: Record<string, string> = {
  xlsx: 'книга Excel', xlsm: 'книга Excel', pptx: 'презентация PowerPoint', odt: 'текстовый документ OpenDocument (.odt)',
  ods: 'таблица OpenDocument', 'xls-or-doc': 'файл старого формата (.doc или .xls)', csv: 'текст с разделителями (CSV)',
};

export async function start(h: Host): Promise<void> {
  host = h;
  setDialogProfile({
    app: 'ENS-Tekst', open: 'Открыть документ',
    filters: [{ name: 'Документы Word', extensions: ['docx'] }, { name: 'Все файлы', extensions: ['*'] }],
  });
  setInterval(() => void recovery.heartbeat().catch(() => {}), 30_000);
  if (isTauri) {
    await setupClose();
    await setupWindowButtons();
  } else {
    setupCloseLocal();
  }
  const params = new URLSearchParams(location.search);
  const open = params.get('open');
  const recover = params.get('recover');
  if (open) await openPath(open);
  else if (recover) {
    const [session, id] = recover.split('/');
    await restore(session, id, params.get('title') ?? 'Документ');
  } else if (!params.get('new')) {
    const startup = await startupFile().catch(() => null);
    if (startup) await openPath(startup);
    await offerRecovery();
  }
}

// --- Открытие ---

export async function openDialog(): Promise<void> {
  const path = await pickOpen();
  if (path) await openPath(path);
}

export async function newWindow(): Promise<void> {
  if (host.isPristine() && !file.path) return;
  await openWindow('new=1', 'Документ');
}

export async function openPath(path: string): Promise<void> {
  if (!host.isPristine() || file.path) {
    // Текущий документ не пустой — открываем рядом, в новом окне, как Word.
    await openWindow(`open=${encodeURIComponent(path)}`, baseName(path));
    return;
  }
  try {
    const info = await inspectFile(path);
    if (info.kind !== 'docx') {
      const what = KIND_NAMES[info.kind] ?? 'файл неизвестного формата';
      toast(`«${info.name}» — это ${what}, а не документ Word. ENS-Tekst открывает .docx` +
        (info.kind === 'xlsx' || info.kind === 'xlsm' ? '; книги Excel открывает ENS-Table.' : '.'));
      return;
    }
    const bytes = await readFile(path);
    await recovery.forget(file.id).catch(() => {});
    tracked = false;
    const lossy = host.openBytes(bytes, baseName(path), { key: path, mtime: info.modified_ms });
    file.path = path;
    file.modifiedMs = info.modified_ms;
    file.lossy = lossy;
    file.lossySource = lossy.length ? path : '';
    if (info.readonly) toast('Файл только для чтения: изменения можно сохранить под другим именем.');
  } catch (e) {
    toast(`Не удалось открыть файл: ${(e as Error)?.message ?? e}`);
  }
}

/** Документ загружен вне этого модуля (например, «Создать»): забываем прежний файл. */
export function detach(): void {
  void recovery.forget(file.id).catch(() => {});
  file.id = newId();
  file.path = '';
  file.modifiedMs = 0;
  file.lossy = [];
  file.lossySource = '';
  file.recoveredFrom = undefined;
  tracked = false;
}

/** Копия с диска (предыдущее сохранение, закрытое без сохранения) открывается в новом окне. */
export async function openCopy(path: string): Promise<void> {
  await openWindow(`open=${encodeURIComponent(path)}`, baseName(path));
}

// --- Сохранение ---

export async function save(asNew = false): Promise<boolean> {
  if (!file.path || asNew) return saveAs();
  return writeTo(file.path, false);
}

async function saveAs(): Promise<boolean> {
  const path = await pickSave(`${host.name().replace(/ \(восстановлено\)$/, '')}.docx`);
  if (!path) return false;
  return writeTo(/\.docx$/i.test(path) ? path : `${path}.docx`, true);
}

async function writeTo(path: string, isNew: boolean): Promise<boolean> {
  // ADR-005: документ с потерями при открытии поверх исходного файла не пишем — потерянное не вернуть.
  if (file.lossy.length && path === file.lossySource) {
    const c = await choose('Поверх этого файла сохранять нельзя',
      `В «${baseName(path)}» есть то, что ENS-Tekst пока не умеет сохранять: ${file.lossy.join('; ')}. Если записать поверх, это пропадёт из файла. Сохраните копию — исходный файл останется как был.`,
      [{ id: 'copy', label: 'Сохранить как…', primary: true }, { id: 'no', label: 'Отмена' }]);
    return c === 'copy' ? saveAs() : false;
  }
  let bytes: Uint8Array;
  try {
    bytes = host.docxBytes(baseName(path));
  } catch (e) {
    toast(`Не удалось подготовить файл: ${(e as Error)?.message ?? e}. Документ в программе не потерян.`);
    return false;
  }
  const expected = isNew ? 0 : file.modifiedMs;
  const written = await (async (): Promise<boolean> => {
    // Страховка: прежнее содержимое файла остаётся в профиле, пока его не затёрли.
    await backupExisting(path);
    try {
      file.modifiedMs = await saveFile(path, bytes, expected);
    } catch (e) {
      if (String(e).includes('CHANGED_ON_DISK')) {
        const overwrite = await ask(
          `Файл «${baseName(path)}» изменён другой программой или пользователем после того, как вы его открыли.\n\n` +
            'Перезаписать его вашей версией? Чтобы сохранить обе версии, выберите «Отмена» и затем «Сохранить как».',
          'Перезаписать', 'Отмена');
        if (!overwrite) return false;
        try {
          file.modifiedMs = await saveFile(path, bytes, 0, true);
        } catch (e2) {
          toast(String(e2));
          return false;
        }
      } else {
        toast(String(e));
        return false;
      }
    }
    return true;
  })();
  if (!written) { host.rollbackSave(); return false; }
  host.afterSaved(path);
  file.path = path;
  if (path !== file.lossySource) { file.lossy = []; file.lossySource = ''; }
  host.markSaved(baseName(path));
  if (file.recoveredFrom) {
    await recovery.discard(file.recoveredFrom.session, file.recoveredFrom.id).catch(() => {});
    file.recoveredFrom = undefined;
  }
  // Сохранённая версия — точка отсчёта: временные файлы больше не нужны.
  await recovery.forget(file.id).catch(() => {});
  tracked = false;
  return true;
}

/** Прежнее содержимое файла — в «Копии на диске» (10 последних на файл): если перезапись оказалась ошибкой, оно не потеряно. */
async function backupExisting(path: string): Promise<void> {
  try {
    await inspectFile(path); // нет файла — копировать нечего
    await historyStore('backup', `${baseName(path)}.docx`, await readFile(path));
  } catch { /* копия — страховка, а не условие сохранения */ }
}

// --- Восстановление после сбоя ---

/** Вызывается через 1,5 с после правки: свежий снимок документа во временной папке. */
export async function persist(): Promise<void> {
  try {
    if (!host.isDirty()) {
      if (tracked) { await recovery.forget(file.id); tracked = false; }
      return;
    }
    if (!tracked) {
      await recovery.begin(file.id, {
        source_path: file.path, title: host.name(), source_modified_ms: file.modifiedMs, snapshot_ms: 0, journal_entries: 0, format: APP_VERSION,
      });
      tracked = true;
    }
    await recovery.snapshot(file.id, new TextEncoder().encode(host.snapshotJson()));
    // Запасная копия в docx: её откроет и другая версия программы (после обновления); не чаще раза в минуту.
    if (Date.now() - lastBackup > 60_000) {
      lastBackup = Date.now();
      try { await recovery.snapshotXlsx(file.id, host.copyBytes()); } catch { /* основной снимок уже записан */ }
    }
  } catch (e) {
    toast(`Не удалось записать копию для восстановления: ${(e as Error)?.message ?? e}`);
  }
}

async function restore(session: string, id: string, title: string): Promise<void> {
  try {
    const list = await recovery.list();
    const meta = list.find((d) => d.session === session && d.id === id)?.meta;
    let notice = '';
    try {
      if (meta && (meta.format ?? '') !== APP_VERSION) throw new Error('другая версия');
      const data = await recovery.load(session, id);
      host.restoreJson(new TextDecoder().decode(data.snapshot), `${meta?.title ?? title} (восстановлено)`);
    } catch (e) {
      // Другая версия (или повреждённый снимок): запасная копия в docx.
      if (!meta?.xlsx_ms) throw e;
      host.openBytes(await recovery.loadXlsx(session, id), `${meta.title} (восстановлено)`);
      notice = `Документ восстановлен из запасной копии на ${fmtTime(meta.xlsx_ms)} (снимок сделан версией ${meta.format || 'до 0.1'}). Правки после этого времени могли не сохраниться.`;
    }
    file.path = meta?.source_path ?? '';
    file.modifiedMs = meta?.source_modified_ms ?? 0;
    file.recoveredFrom = { session, id };
    tracked = false;
    await persist();
    if (notice) toast(notice);
  } catch (e) {
    // Файлы не удаляем: их можно восстановить в той версии, а исходный документ не менялся.
    toast(`Не удалось восстановить документ: ${(e as Error)?.message ?? e}. Временные файлы сохранены, исходный файл не изменялся.`);
  }
}

async function offerRecovery(): Promise<void> {
  const list = await recovery.list().catch(() => [] as RecoverableDoc[]);
  if (!list.length) return;
  const render = (docs: RecoverableDoc[]) => dialog((box, close) => {
    if (!docs.length) { close(); return; }
    box.append(
      el('h2', { textContent: 'Восстановление документов' }),
      el('p', { textContent: 'Программа была закрыта аварийно. Эти документы можно восстановить на момент сбоя. Исходные файлы не изменялись — откройте обе версии и сохраните нужную.' }),
    );
    for (const d of docs) {
      const open = async () => {
        close();
        if (host.isPristine() && !file.path) await restore(d.session, d.id, d.meta.title);
        else await openWindow(`recover=${d.session}/${d.id}&title=${encodeURIComponent(d.meta.title)}`, d.meta.title);
        render(docs.filter((x) => x !== d));
      };
      box.append(el('div', { class: 'rec-item' },
        el('div', { class: 'rec-title', textContent: d.meta.title }),
        el('div', { class: 'rec-meta', textContent: `${d.meta.source_path || 'новый документ, не сохранялся'} · сбой: ${fmtTime(d.crashed_ms)}, последняя копия: ${d.meta.snapshot_ms ? fmtTime(d.meta.snapshot_ms) : '—'}${d.source_changed ? ' · исходный файл менялся после открытия' : ''}` }),
        el('div', { class: 'rec-buttons' },
          el('button', { class: 'primary', textContent: 'Открыть восстановленный', onclick: open }),
          d.source_exists ? el('button', { textContent: 'Открыть исходный рядом', onclick: () => void openWindow(`open=${encodeURIComponent(d.meta.source_path)}`, `${d.meta.title} (исходный)`) }) : '',
          el('button', {
            class: 'danger', textContent: 'Удалить временные файлы',
            onclick: async () => {
              if (!(await ask(`Удалить временные файлы документа «${d.meta.title}»? Останется только версия, сохранённая вами.`, 'Удалить', 'Отмена'))) return;
              await recovery.discard(d.session, d.id);
              close();
              render(docs.filter((x) => x !== d));
            },
          }))));
    }
    box.append(el('div', { class: 'rec-actions' }, el('button', { textContent: 'Решить позже', onclick: close })));
  }, false);
  render(list);
}

// --- Окно ---

async function setupWindowButtons(): Promise<void> {
  const { getCurrentWindow } = await import('@tauri-apps/api/window');
  const win = getCurrentWindow();
  const box = document.getElementById('window-buttons')!;
  box.hidden = false;
  const max = box.querySelector<HTMLButtonElement>('[data-win="maximize"]')!;
  const sync = async () => {
    const on = await win.isMaximized().catch(() => false);
    max.title = on ? 'Свернуть в окно' : 'Развернуть';
    max.innerHTML = on
      ? '<svg viewBox="0 0 10 10"><rect x="0.5" y="2.5" width="7" height="7"/><path d="M2.5 2.5V0.5h7v7h-2"/></svg>'
      : '<svg viewBox="0 0 10 10"><rect x="0.5" y="0.5" width="9" height="9"/></svg>';
  };
  box.querySelector('[data-win="minimize"]')!.addEventListener('click', () => void win.minimize());
  max.addEventListener('click', () => void win.toggleMaximize());
  box.querySelector('[data-win="close"]')!.addEventListener('click', () => void win.close());
  await win.onResized(() => void sync());
  await sync();
}

/** Версия для РЕД ОС 7: окно закрыли — браузер спросит «Покинуть сайт?», а копия для восстановления останется. */
function setupCloseLocal(): void {
  window.addEventListener('beforeunload', (e) => {
    if (host.isDirty()) { void persist(); e.preventDefault(); e.returnValue = ''; }
  });
  window.addEventListener('pagehide', () => beaconClose(file.id, !host.isDirty()));
}

async function setupClose(): Promise<void> {
  const { getCurrentWindow } = await import('@tauri-apps/api/window');
  const { getAllWebviewWindows } = await import('@tauri-apps/api/webviewWindow');
  const win = getCurrentWindow();
  await win.onCloseRequested(async (event) => {
    event.preventDefault();
    if (host.isDirty()) {
      await persist();
      const c = await choose('Сохранить изменения?', `Если не сохранить, изменения в «${host.name()}» будут потеряны.`,
        [{ id: 'save', label: 'Сохранить', primary: true }, { id: 'discard', label: 'Не сохранять' }, { id: 'cancel', label: 'Отмена' }]);
      if (c === 'cancel') return;
      if (c === 'save' && !(await save())) return;
      // «Не сохранять» — частая ошибка: копия документа остаётся в «Копиях на диске» на 14 дней.
      if (c === 'discard') await historyStore('discarded', `${host.name()}.docx`, host.copyBytes()).catch(() => {});
    }
    await recovery.forget(file.id).catch(() => {});
    if ((await getAllWebviewWindows()).length <= 1) await recovery.endSession().catch(() => {});
    await win.destroy();
  });
}
