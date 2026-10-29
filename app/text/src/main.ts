// ENS-Text: текстовый редактор ENS-Office. Лента, команды, файлы, восстановление.

import '../../table/src/styles.css';
import './text.css';
import './icons';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { TextSelection } from 'prosemirror-state';
import { redo, redoDepth, undo, undoDepth } from 'prosemirror-history';
import {
  addColumnAfter, addColumnBefore, addRowAfter, addRowBefore, deleteColumn, deleteRow, deleteTable, isInTable, mergeCells, splitCell,
} from 'prosemirror-tables';
import type { Node as PMNode } from 'prosemirror-model';
import { register, run, handleKey } from '../../table/src/commands';
import { FONT_SIZES, FONTS, lastColor, renderRibbon, type RibbonTab } from '../../table/src/ribbon';
import { fontStack, loadFonts, OFFICE_FONTS } from '../../table/src/fonts';
import { ICONS } from '../../table/src/icons';
import { APP_VERSION } from './version';
import { schema } from './schema';
import {
  clearFormatting, createState, insertField, currentLevel, currentStyle, fmtActive, findKey, findStep, indent, insertImage, insertPageBreak, insertTable, linkAt,
  listKind, markActive, paraAttr, replaceAll, replaceMatch, setLink, setParaAttrs, setStyle, setTextStyle, toggle, toggleList, toggles,
} from './editor';
import { A4, exportDocx, importDocx, newMeta, type DocMeta, type HfDoc, type PageSetup } from './docx';
import { clearRecovery, listRecovery, saveRecovery, type RecoveryRecord } from './recovery';
import * as desktop from './desktop';
import * as V from './versions';
import { askNamedVersion, fmtWhen, openHistory, type HistoryHost } from './history-ui';
import { saveSettings, settings } from '../../table/src/settings';
import { osUserName } from '../../table/src/platform';
import { readText } from '../../table/src/platform';
import { dialog, el, setRestoreFocus, toast } from './ui';
import { initThemes, themeDialog } from './themes-ui';
import { checkDialog, renderInspector, toggleInspector } from './inspector-ui';
import type { CheckHost } from './checks';
import { applyScheme, continueList, convertManual, currentStart, setListStart } from './listtools';
import { PRESETS, presetOf, type Scheme } from './numbering';
import { schemeDialog } from './lists-ui';
import { describe as describeTypo, liveQuotes, typograph } from './typograph';
import { focusKey, jumpFlag, listFlags, toggleFlag } from './editor';
import { allCommands, matchScore } from '../../table/src/commands';
import { PX, geometry, layout, pageCount, pagination, replan, stats as pagStats } from './pagination';
import { hfFor, measureHf, openHfEditor, renderHf, setHf, type HfKind, type HfSession } from './hf';

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

// ---------------------------------------------------------------- состояние документа

let meta: DocMeta = newMeta();
let savedMeta = JSON.stringify(meta.page) + JSON.stringify(meta.defaults) + JSON.stringify(meta.hf);
let docName = 'Документ1';
let savedDoc: PMNode = schema.nodes.doc.createAndFill()!;
const docId = `doc-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
let zoom = 1;
/** Открытый для правки колонтитул (если есть). Команды форматирования действуют на него, а не на основной текст. */
let hfSession: HfSession | null = null;

function emptyDoc(): PMNode {
  return schema.nodes.doc.createAndFill()!;
}

const metaKey = () => JSON.stringify(meta.page) + JSON.stringify(meta.defaults) + JSON.stringify(meta.hf);
const isDirty = () => !view.state.doc.eq(savedDoc) || metaKey() !== savedMeta;


// ---------------------------------------------------------------- версии, журнал, защита от потерь

/** История версий и журнал изменений: хранятся в самом файле (customXml), в снимке восстановления и здесь. */
let hist: V.History = V.emptyHistory();
let histRev = 0;
let histCache = { rev: -1, text: '' };
let osUser = '';
void osUserName().then((n) => { osUser = n; });
const me = () => settings.userName.trim() || osUser || 'Пользователь';
let lastSnapDoc: PMNode | null = null;
let lastVersionAt = Date.now();
let dirtySince = 0;
let lastSavedAt = 0;
let lastReminder = 0;

const histText = (): string => {
  if (histCache.rev !== histRev) histCache = { rev: histRev, text: V.serialize(hist) };
  return histCache.text;
};

/** Записать версию (если содержимое изменилось с последней). Та же версия получает новый вид (автоверсия → сохранение, именная). */
function snap(kind: V.VersionKind, label = ''): V.Version | null {
  const hash = V.docHash(view.state.doc, meta);
  const last = hist.versions[hist.versions.length - 1];
  if (last && last.hash === hash) {
    if (kind === 'named' || kind === 'save') {
      last.kind = kind === 'save' && last.kind === 'named' ? 'named' : kind;
      last.t = Date.now();
      last.u = me();
      if (label) last.label = label;
      histRev++;
      return last;
    }
    return null;
  }
  const v = V.addVersion(hist, view.state.doc, meta, { u: me(), kind, label });
  lastSnapDoc = view.state.doc;
  lastVersionAt = Date.now();
  V.trim(hist);
  histRev++;
  return v;
}

let saveToken: { versions: V.Version[]; log: V.LogEntry[] } | null = null;

/** Собирает docx для записи на диск: версия «Сохранение», запись в журнале, история внутри файла. Проверяет файл чтением. */
function buildDocx(name: string): Uint8Array {
  const plain = exportDocx(view.state.doc, meta, name);
  const back = importDocx(plain); // проверка: только что собранный файл должен читаться
  const hash = V.docHash(back.doc, back.meta);
  saveToken = { versions: hist.versions.map((v) => ({ ...v })), log: [...hist.log] };
  const prevBase = hist.versions[hist.versions.length - 1];
  const v = snap('save') ?? hist.versions[hist.versions.length - 1];
  const idx = hist.versions.indexOf(v);
  const base = hist.versions[idx - 1] ?? (prevBase !== v ? prevBase : undefined);
  const ch = V.summarize(hist, base?.blocks ?? [], v.blocks, base?.words ?? 0, v.words);
  if (base && JSON.stringify(base.meta) !== JSON.stringify(v.meta)) ch.notes.push('изменены параметры страницы или колонтитулы');
  const last = hist.log[hist.log.length - 1];
  const changed = ch.added || ch.removed || ch.changed || ch.notes.length;
  if (changed || !last) V.addLog(hist, { u: me(), t: Date.now(), app: `ENS-Text ${APP_VERSION}`, kind: 'save', ch, h: hash });
  else { last.h = hash; last.t = Date.now(); }
  histRev++;
  return exportDocx(view.state.doc, meta, name, histText());
}

/** Запись на диск не удалась: версия и запись журнала «Сохранение» не должны остаться. */
function rollbackSave(): void {
  if (!saveToken) return;
  hist.versions = saveToken.versions;
  hist.log = saveToken.log;
  V.collect(hist);
  saveToken = null;
  histRev++;
}

/** Файл записан: отметка «когда сохранено», «последнее открытие» для вопроса «что изменилось». */
function afterSaved(key: string): void {
  saveToken = null;
  lastSavedAt = Date.now();
  dirtySince = 0;
  markOpened(key);
}

function markOpened(key: string): void {
  if (!key) return;
  const all = { ...settings.lastOpened, [key]: Date.now() };
  const keys = Object.keys(all).sort((a, b) => all[b] - all[a]).slice(0, 200);
  settings.lastOpened = Object.fromEntries(keys.map((k) => [k, all[k]]));
  saveSettings();
}

/** Открыт файл: прочитать историю, заметить правку в другой программе, предложить показать, что изменилось. */
function adoptHistory(raw: string, info: { key: string; mtime?: number }): { text: string; buttons: { label: string; run: () => void }[] } | null {
  hist = V.parse(raw);
  histRev++;
  lastSnapDoc = view.state.doc;
  const hash = V.docHash(view.state.doc, meta);
  const last = hist.log[hist.log.length - 1];
  const since = settings.lastOpened[info.key] ?? 0;
  markOpened(info.key);
  let outside: V.Version | null = null;
  if (last && last.h !== hash) {
    const prev = hist.versions[hist.versions.length - 1];
    outside = snap('outside', 'Состояние файла при открытии') ?? null;
    const ch = V.summarize(hist, prev?.blocks ?? [], outside?.blocks ?? [], prev?.words ?? 0, outside?.words ?? 0);
    V.addLog(hist, { u: 'неизвестно', t: info.mtime ?? Date.now(), app: 'другая программа', kind: 'outside', ch, h: hash });
  } else if (!hist.versions.length) {
    snap('open', 'Версия при первом открытии');
  }
  const fresh = hist.log.filter((e) => e.kind === 'save' && e.t > since && e.u !== me());
  if (outside) {
    return {
      text: 'Файл менялся в другой программе после последнего сохранения в ENS-Text. Прежняя версия сохранена в истории.',
      buttons: [{ label: 'Что изменилось', run: () => showHistory('versions', outside!.id, 'previous') }],
    };
  }
  if (since && fresh.length) {
    const who = [...new Set(fresh.map((e) => e.u))].join(', ');
    const base = [...hist.versions].reverse().find((v) => v.t <= since) ?? hist.versions[0];
    return {
      text: `С вашего прошлого открытия документ сохраняли: ${who} (${fresh.length}).`,
      buttons: [{ label: 'Что изменилось', run: () => showHistory('versions', base?.id) }],
    };
  }
  return null;
}

function showHistory(tab: 'versions' | 'disk' = 'versions', select?: string, against?: 'current' | 'previous'): void {
  openHistory(historyHost, tab, select, against);
}

const historyHost: HistoryHost = {
  history: () => hist,
  doc: () => view.state.doc,
  meta: () => meta,
  user: me,
  docName: () => docName,
  saveNamed(label) { snap('named', label); updateUi(); },
  restore(v) {
    snap('before', `Перед возвратом к версии от ${fmtWhen(v.t)}`);
    const d = V.versionDoc(hist, v);
    view.dispatch(view.state.tr.replaceWith(0, view.state.doc.content.size, d.content));
    const vm = V.versionMeta(v);
    meta.page = vm.page;
    meta.defaults = vm.defaults;
    meta.hf = vm.hf;
    hfPx = { header: measureHf(meta.hf, meta.page, 'header'), footer: measureHf(meta.hf, meta.page, 'footer') };
    applyPage();
    replan(view);
    updateUi();
    toast('Документ возвращён к выбранной версии. Отменить можно Ctrl+Z.');
  },
  revert(range, nodes) {
    snap('before', 'Перед возвратом фрагмента');
    const doc = view.state.doc;
    const offs: number[] = [];
    doc.forEach((_n, off) => offs.push(off));
    const at = (i: number) => (i >= offs.length ? doc.content.size : offs[i]);
    const insert = nodes.length || range[1] < doc.childCount || range[0] > 0 ? nodes : [schema.nodes.paragraph.create()];
    view.dispatch(view.state.tr.replaceWith(at(range[0]), at(range[1]), insert));
  },
  openCopy: (path) => { void desktop.openCopy(path); },
};

setInterval(() => {
  // Автоверсия: раз в N минут, если с последней версии документ менялся.
  if (view.state.doc !== lastSnapDoc && isDirty() && Date.now() - lastVersionAt >= Math.max(1, settings.autosaveMinutes) * 60_000) { snap('auto'); updateUi(); }
  // Напоминание: давно не сохраняли.
  const mins = settings.saveReminderMinutes;
  if (mins > 0 && dirtySince && Date.now() - dirtySince >= mins * 60_000 && Date.now() - lastReminder >= mins * 60_000 && $('notice').hidden) {
    lastReminder = Date.now();
    showNotice(`Документ не сохранялся ${Math.round((Date.now() - dirtySince) / 60_000)} мин.`, [
      { label: 'Сохранить', run: () => { hideNotice(); run('file.save'); } },
      { label: 'Напомнить позже', run: hideNotice },
    ]);
  }
  updateSavedStatus();
}, 20_000);

function updateSavedStatus(): void {
  const el2 = $('status-saved');
  if (isDirty()) {
    const m = dirtySince ? Math.floor((Date.now() - dirtySince) / 60_000) : 0;
    el2.textContent = m >= 1 ? `Не сохранено ${m} мин` : 'Есть несохранённые изменения';
    el2.className = m >= Math.max(1, settings.saveReminderMinutes || 10) ? 'warn' : '';
  } else {
    el2.textContent = lastSavedAt ? `Сохранено в ${new Date(lastSavedAt).toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' })}` : '';
    el2.className = '';
  }
}

// ---------------------------------------------------------------- страница

/** Измеренная высота колонтитулов, px: если они выше поля страницы, текст начинается ниже (как в Word). */
let hfPx = { header: 0, footer: 0 };

/** Страница с учётом колонтитулов: top и bottom — где на самом деле начинается и кончается текст. */
function effPage(): PageSetup {
  const p = meta.page;
  const hf = meta.hf;
  return {
    ...p,
    top: hfPx.header ? Math.max(p.top, hf.headerDist + hfPx.header / PX) : p.top,
    bottom: hfPx.footer ? Math.max(p.bottom, hf.footerDist + hfPx.footer / PX) : p.bottom,
  };
}

/** Колонтитулы изменились: перемерить; если поля текста сдвинулись — пересчитать страницы. */
function refreshHf(): void {
  const before = JSON.stringify([effPage().top, effPage().bottom]);
  hfPx = { header: measureHf(meta.hf, meta.page, 'header'), footer: measureHf(meta.hf, meta.page, 'footer') };
  if (before !== JSON.stringify([effPage().top, effPage().bottom])) {
    applyPage();
    replan(view);
  } else {
    paintHf();
  }
}

function applyPage(): void {
  const p = meta.page;
  const e = effPage();
  const sheet = $('sheet');
  sheet.style.width = `${p.w}pt`;
  view.dom.style.padding = `${e.top}pt ${e.right}pt ${e.bottom}pt ${e.left}pt`;
  drawPapers(pageCount(view));
  const d = meta.defaults;
  const root = document.documentElement.style;
  root.setProperty('--doc-font', fontStack(d.font));
  root.setProperty('--doc-size', `${d.size}pt`);
  root.setProperty('--doc-line', d.line === 1 ? 'normal' : String((d.line * 1.15).toFixed(3)));
  root.setProperty('--doc-after', `${d.after}pt`);
  root.setProperty('--doc-before', `${d.before}pt`);
  // Печать идёт по нашей вёрстке: поля уже внутри страницы, поэтому у листа бумаги полей нет.
  let st = document.getElementById('print-page') as HTMLStyleElement | null;
  if (!st) { st = document.createElement('style'); st.id = 'print-page'; document.head.append(st); }
  st.textContent = `@page { size: ${p.w}pt ${p.h}pt; margin: 0; }`;
}

/** Листы-подложки: по одному на страницу; высота страницы и зазоры — как у расчёта разрывов. */
function drawPapers(pages: number): void {
  const g = geometry(effPage());
  const papers = $('papers');
  if (papers.childElementCount !== pages || papers.dataset.h !== String(g.H) || papers.dataset.gap !== String(layout.gap)) {
    papers.replaceChildren(...Array.from({ length: pages }, (_, i) => {
      const d = document.createElement('div');
      d.className = 'paper';
      d.style.cssText = `top:${i * g.P}px;height:${g.H}px`;
      return d;
    }));
    papers.dataset.h = String(g.H);
    papers.dataset.gap = String(layout.gap);
  }
  const total = pages * g.H + (pages - 1) * layout.gap;
  $('sheet').style.minHeight = `${total}px`;
  $('sheet').style.setProperty('--sheet-h', `${total}px`);
  view.dom.style.minHeight = `${total}px`;
  paintHf(pages);
  updatePageStatus(pages);
}

/** Колонтитулы на листах: свои на каждой странице (номер страницы подставляется). */
function paintHf(pages = pageCount(view)): void {
  const papers = $('papers');
  [...papers.children].forEach((paper, i) => {
    paper.querySelectorAll('.hf').forEach((e) => e.remove());
    for (const kind of ['header', 'footer'] as const) {
      if (hfSession && hfSession.kind === kind && hfSession.pageIndex === i) continue; // здесь идёт правка
      const e = renderHf(hfFor(meta.hf, kind, i), kind, i + 1, pages);
      if (!e) continue;
      e.style.left = `${meta.page.left}pt`;
      e.style.right = `${meta.page.right}pt`;
      if (kind === 'header') e.style.top = `${meta.hf.headerDist}pt`;
      else e.style.bottom = `${meta.hf.footerDist}pt`;
      paper.append(e);
    }
  });
}

/** Страница (с 1), на которой стоит курсор. */
function currentPage(n = pageCount(view)): number {
  try {
    const box = view.dom.getBoundingClientRect();
    const scale = box.width / (view.dom as HTMLElement).offsetWidth || 1;
    const c = view.coordsAtPos(view.state.selection.head);
    return Math.min(n, Math.max(1, Math.floor((((c.top - box.top) / scale) + 1) / geometry(effPage()).P) + 1));
  } catch {
    return 1; // вид ещё не готов
  }
}

/** «Страница X из N» по положению курсора. */
function updatePageStatus(pages = pageCount(view)): void {
  $('status-page').textContent = `Страница ${currentPage(pages)} из ${pages}`;
}

function setZoom(z: number): void {
  zoom = Math.min(3, Math.max(0.3, Math.round(z * 10) / 10));
  ($('sheet-wrap').style as CSSStyleDeclaration & { zoom: string }).zoom = String(zoom);
  $('zoom-val').textContent = `${Math.round(zoom * 100)}%`;
}

// ---------------------------------------------------------------- редактор

const paginationPlugin = pagination(effPage, (n) => drawPapers(n));
let liveQuotesOn = (() => { try { return localStorage.getItem('ens-text.liveQuotes') !== '0'; } catch { return true; } })();
const quotesPlugin = liveQuotes(() => liveQuotesOn);

const view = new EditorView($('sheet'), {
  state: createState(emptyDoc(), [paginationPlugin, quotesPlugin]),
  editable: () => !hfSession,
  attributes: { spellcheck: 'true', lang: 'ru', 'aria-label': 'Текст документа' },
  dispatchTransaction(tr) {
    const before = view.state;
    view.updateState(before.apply(tr));
    afterChange(tr.docChanged);
  },
  handlePaste(v, event) {
    const files = [...(event.clipboardData?.files ?? [])].filter((f) => f.type.startsWith('image/'));
    if (!files.length) return false;
    void insertImageFile(files[0]);
    return true;
  },
  handleDrop(v, event) {
    const files = [...(event.dataTransfer?.files ?? [])].filter((f) => f.type.startsWith('image/'));
    if (!files.length) return false;
    void insertImageFile(files[0]);
    return true;
  },
  handleClick(v, pos, event) {
    if (!event.ctrlKey && !event.metaKey) return false;
    const $pos = v.state.doc.resolve(pos);
    const link = schema.marks.link.isInSet($pos.marks());
    if (!link) return false;
    if (desktop.isDesktop) {
      desktop.openTarget(link.attrs.href).catch((e) => toast(String((e as Error)?.message ?? e)));
      return true;
    }
    const a = document.createElement('a');
    a.href = link.attrs.href;
    a.target = '_blank';
    a.rel = 'noopener noreferrer';
    a.click();
    return true;
  },
});

let statsTimer = 0;
let recoveryTimer = 0;

function afterChange(docChanged: boolean): void {
  updateUi();
  renderInspector(inspectorSource, inspectorActions);
  if (docChanged) {
    clearTimeout(statsTimer);
    statsTimer = window.setTimeout(updateStats, 250);
    scheduleRecovery();
  } else {
    updateSelStats();
  }
}

function updateStats(): void {
  const text = view.state.doc.textBetween(0, view.state.doc.content.size, '\n', ' ');
  const words = (text.match(/[\p{L}\p{N}][\p{L}\p{N}'’\-]*/gu) ?? []).length;
  $('status-words').textContent = `Слов: ${words}  ·  Знаков: ${text.replace(/\n/g, '').length}`;
  updateSelStats();
}

function updateSelStats(): void {
  updatePageStatus();
  const { from, to, empty } = view.state.selection;
  if (empty) { $('status-sel').textContent = ''; return; }
  const text = view.state.doc.textBetween(from, to, '\n', ' ');
  const words = (text.match(/[\p{L}\p{N}][\p{L}\p{N}'’\-]*/gu) ?? []).length;
  $('status-sel').textContent = `Выделено слов: ${words}  ·  знаков: ${text.replace(/\n/g, '').length}`;
}

// ---------------------------------------------------------------- команды

const active = (): EditorView => hfSession?.view ?? view;

const dispatchCmd = (cmd: (s: EditorState, d?: typeof view.dispatch) => boolean): void => {
  cmd(active().state, active().dispatch);
  active().focus();
};

const ALIGN_LABEL = { left: 'По левому краю', center: 'По центру', right: 'По правому краю', justify: 'По ширине' } as const;
const FORMAT_LABEL: Record<string, string> = {
  bold: 'Полужирный', italic: 'Курсив', underline: 'Подчёркнутый', strike: 'Зачёркнутый', superscript: 'Надстрочный', subscript: 'Подстрочный',
};
const FORMAT_KEYS: Record<string, string[]> = { bold: ['Ctrl+KeyB'], italic: ['Ctrl+KeyI'], underline: ['Ctrl+KeyU'] };

for (const name of Object.keys(toggles)) {
  register({
    id: `format.${name}`, label: FORMAT_LABEL[name], keys: FORMAT_KEYS[name],
    run: () => dispatchCmd(toggle(name)),
    checked: () => markActive(view.state, toggles[name]),
  });
}
register({ id: 'format.superscript', label: 'Надстрочный', keys: ['Ctrl+Shift+Equal'], run: () => dispatchCmd(toggle('superscript')), checked: () => markActive(view.state, toggles.superscript) });
register({ id: 'format.subscript', label: 'Подстрочный', keys: ['Ctrl+Equal'], run: () => dispatchCmd(toggle('subscript')), checked: () => markActive(view.state, toggles.subscript) });

register(
  { id: 'edit.undo', label: 'Отменить', keys: ['Ctrl+KeyZ'], run: () => { undo(active().state, active().dispatch); active().focus(); } },
  { id: 'edit.redo', label: 'Повторить', keys: ['Ctrl+KeyY', 'Ctrl+Shift+KeyZ'], run: () => { redo(active().state, active().dispatch); active().focus(); } },
  {
    id: 'clipboard.paste', label: 'Вставить (Ctrl+V)',
    run: () => {
      active().focus();
      (desktop.isDesktop ? readText() : navigator.clipboard.readText()).then((t) => { if (t) active().dispatch(active().state.tr.insertText(t).scrollIntoView()); })
        .catch(() => toast('Браузер не дал доступ к буферу обмена — нажмите Ctrl+V'));
    },
  },
  { id: 'clipboard.cut', label: 'Вырезать (Ctrl+X)', run: () => { active().focus(); document.execCommand('cut'); } },
  { id: 'clipboard.copy', label: 'Копировать (Ctrl+C)', run: () => { active().focus(); document.execCommand('copy'); } },
  { id: 'format.fontName', label: 'Шрифт', run: (v) => v && dispatchCmd(setTextStyle({ font: v })) },
  { id: 'format.fontSize', label: 'Размер шрифта', run: (v) => v && dispatchCmd(setTextStyle({ size: parseFloat(v) })) },
  { id: 'format.fontColor', label: 'Цвет текста', run: (v) => dispatchCmd(setTextStyle({ color: v || null })) },
  { id: 'format.highlight', label: 'Цвет выделения', run: (v) => dispatchCmd(setTextStyle({ bg: v || null })) },
  { id: 'format.clear', label: 'Очистить форматирование', keys: ['Ctrl+Space'], run: () => dispatchCmd(clearFormatting) },
  { id: 'format.sizeUp', label: 'Увеличить размер', keys: ['Ctrl+Shift+Period'], run: () => stepSize(1) },
  { id: 'format.sizeDown', label: 'Уменьшить размер', keys: ['Ctrl+Shift+Comma'], run: () => stepSize(-1) },
);

function stepSize(dir: 1 | -1): void {
  const sizes = FONT_SIZES.map((o) => +o.value);
  const cur = currentStyle(active().state).size ?? meta.defaults.size;
  const next = dir > 0 ? sizes.find((s) => s > cur) : [...sizes].reverse().find((s) => s < cur);
  if (next) dispatchCmd(setTextStyle({ size: next }));
}

for (const a of ['left', 'center', 'right', 'justify'] as const) {
  register({
    id: `align.${a}`, label: ALIGN_LABEL[a], keys: [{ left: 'Ctrl+KeyL', center: 'Ctrl+KeyE', right: 'Ctrl+KeyR', justify: 'Ctrl+KeyJ' }[a]],
    run: () => dispatchCmd(setParaAttrs({ align: a === 'left' ? null : a })),
    checked: () => (paraAttr<string>(active().state, 'align') ?? 'left') === a,
  });
}
register(
  { id: 'list.bullet', label: 'Маркированный список', keys: ['Ctrl+Shift+KeyL'], run: () => dispatchCmd(toggleList(schema.nodes.bullet_list)), checked: () => listKind(active().state) === 'bullet' },
  { id: 'list.ordered', label: 'Нумерованный список', run: () => dispatchCmd(toggleList(schema.nodes.ordered_list)), checked: () => listKind(active().state) === 'ordered' },
  { id: 'indent.inc', label: 'Увеличить отступ', keys: ['Ctrl+KeyM'], run: () => dispatchCmd(indent(1)) },
  { id: 'indent.dec', label: 'Уменьшить отступ', keys: ['Ctrl+Shift+KeyM'], run: () => dispatchCmd(indent(-1)) },
  { id: 'para.line', label: 'Межстрочный интервал', run: (v) => dispatchCmd(setParaAttrs({ line: v && v !== '1' ? parseFloat(v) : null })) },
  { id: 'para.after', label: 'Интервал после абзаца', run: (v) => dispatchCmd(setParaAttrs({ after: v === '' || v === undefined ? null : parseFloat(v) })) },
  { id: 'para.style', label: 'Стиль абзаца', run: (v) => dispatchCmd(setStyle(+(v ?? 0))) },
  { id: 'edit.find', label: 'Найти', keys: ['Ctrl+KeyF'], run: () => openFind(false) },
  { id: 'edit.replace', label: 'Заменить', keys: ['Ctrl+KeyH'], run: () => openFind(true) },
  { id: 'insert.table', label: 'Таблица', run: () => tableDialog() },
  { id: 'insert.image', label: 'Рисунок', run: () => $('image-input').click() },
  { id: 'insert.link', label: 'Гиперссылка', keys: ['Ctrl+KeyK'], run: () => linkDialog() },
  { id: 'insert.pageBreak', label: 'Разрыв страницы', run: () => dispatchCmd(insertPageBreak) },
  { id: 'table.rowBefore', label: 'Вставить строку выше', run: () => dispatchCmd(addRowBefore) },
  { id: 'table.rowAfter', label: 'Вставить строку ниже', run: () => dispatchCmd(addRowAfter) },
  { id: 'table.colBefore', label: 'Вставить столбец слева', run: () => dispatchCmd(addColumnBefore) },
  { id: 'table.colAfter', label: 'Вставить столбец справа', run: () => dispatchCmd(addColumnAfter) },
  { id: 'table.delRow', label: 'Удалить строку', run: () => dispatchCmd(deleteRow) },
  { id: 'table.delCol', label: 'Удалить столбец', run: () => dispatchCmd(deleteColumn) },
  { id: 'table.delete', label: 'Удалить таблицу', run: () => dispatchCmd(deleteTable) },
  { id: 'table.merge', label: 'Объединить ячейки', run: () => dispatchCmd(mergeCells) },
  { id: 'table.split', label: 'Разделить ячейку', run: () => dispatchCmd(splitCell) },
  { id: 'table.borders', label: 'Границы таблицы', run: () => toggleBorders() },
  { id: 'page.size', label: 'Размер страницы', run: (v) => setPageSize(v ?? 'a4') },
  { id: 'page.orientation', label: 'Ориентация', run: (v) => setOrientation(v === 'landscape') },
  { id: 'page.margins', label: 'Поля', run: (v) => setMargins(v ?? 'normal') },
  { id: 'view.zoomIn', label: 'Увеличить масштаб', run: () => setZoom(zoom + 0.1) },
  { id: 'view.zoomOut', label: 'Уменьшить масштаб', run: () => setZoom(zoom - 0.1) },
  { id: 'view.zoomReset', label: 'Масштаб 100%', run: () => setZoom(1) },
  {
    id: 'list.scheme', label: 'Нумерация пунктов',
    run: (v) => {
      if (!v) return;
      if (v === 'custom') { schemeDialog(currentScheme(), (s) => { dispatchCmd(applyScheme(s)); }); setSelect('list.scheme', ''); return; }
      dispatchCmd(applyScheme(PRESETS.find((p) => p.id === v)?.scheme ?? null));
    },
  },
  { id: 'list.restart', label: 'Начать нумерацию заново', run: () => { if (!setListStart(1)(active().state, active().dispatch)) toast('Поставьте курсор в нумерованный список'); active().focus(); } },
  { id: 'list.startAt', label: 'Начать нумерацию с…', run: () => startAtDialog() },
  { id: 'list.continue', label: 'Продолжить нумерацию предыдущего списка', run: () => { if (!continueList(active().state, active().dispatch)) toast('Выше нет нумерованного списка, который можно продолжить'); active().focus(); } },
  {
    id: 'list.fromManual', label: 'Номера, набранные вручную, в список',
    run: () => {
      snap('before', 'Перед преобразованием ручных номеров в список');
      const n = convertManual(active().state, (tr) => active().dispatch(tr));
      toast(n ? `Преобразовано пунктов: ${n}. Отменить можно Ctrl+Z.` : 'Ручной нумерации не найдено: нужны подряд идущие абзацы вида «1. …», «1.1 …», «а) …».');
      active().focus();
    },
  },
  {
    id: 'typo.apply', label: 'Типограф: кавычки, тире, неразрывные пробелы', keywords: ['типограф', 'кавычки', 'ёлочки', 'тире', 'пробелы', 'typography'],
    run: () => {
      const { tr, report } = typograph(active().state);
      if (report.total) { snap('before', 'Перед типографом'); active().dispatch(tr); }
      toast(describeTypo(report));
      active().focus();
    },
  },
  {
    id: 'typo.live', label: 'Кавычки «ёлочки» при наборе', checked: () => liveQuotesOn,
    run: () => { liveQuotesOn = !liveQuotesOn; try { localStorage.setItem('ens-text.liveQuotes', liveQuotesOn ? '1' : '0'); } catch { /* без памяти браузера */ } updateUi(); },
  },
  { id: 'mark.red', label: 'Метка «проверить» (красная)', keywords: ['метка', 'флажок', 'проверить'], keys: ['Ctrl+Alt+KeyR'], run: () => dispatchCmd(toggleFlag('red')) },
  { id: 'mark.green', label: 'Метка «готово» (зелёная)', keywords: ['метка', 'флажок', 'готово'], keys: ['Ctrl+Alt+KeyG'], run: () => dispatchCmd(toggleFlag('green')) },
  { id: 'mark.next', label: 'К следующей метке', keys: ['Ctrl+Alt+ArrowDown'], run: () => dispatchCmd(jumpFlag(1)) },
  { id: 'mark.prev', label: 'К предыдущей метке', keys: ['Ctrl+Alt+ArrowUp'], run: () => dispatchCmd(jumpFlag(-1)) },
  { id: 'mark.list', label: 'Список меток', run: () => flagsDialog() },
  {
    id: 'view.focusPara', label: 'Фокус абзаца', keywords: ['фокус', 'читать', 'вычитка'], keys: ['Ctrl+Shift+KeyK'],
    checked: () => !!focusKey.getState(view.state),
    run: () => { view.dispatch(view.state.tr.setMeta(focusKey, !focusKey.getState(view.state))); updateUi(); },
  },
  { id: 'review.check', label: 'Проверка оформления (ГОСТ)', keywords: ['гост', 'оформление', 'проверка', 'поля', 'шрифт', 'интервал'], run: () => checkDialog(checkHost, (l) => { snap('before', l); }, () => view) },
  { id: 'review.inspector', label: 'Рентген форматирования', keywords: ['формат', 'инспектор', 'откуда'], keys: ['Ctrl+Alt+KeyI'], run: () => toggleInspector(inspectorSource, inspectorActions) },
  { id: 'view.theme', label: 'Тема оформления', keywords: ['theme', 'тема', 'оформление', 'пустошь', 'матрица'], run: () => themeDialog() },
  { id: 'ver.history', label: 'История версий', keys: ['Ctrl+Alt+KeyH'], run: () => showHistory() },
  { id: 'ver.named', label: 'Записать именную версию', run: () => askNamedVersion(historyHost) },
  { id: 'ver.disk', label: 'Копии на диске', run: () => showHistory('disk') },
  { id: 'ver.reminder', label: 'Напоминать о сохранении', run: (v) => { settings.saveReminderMinutes = Math.max(0, parseInt(v ?? '10', 10) || 0); saveSettings(); lastReminder = 0; } },
  { id: 'hf.editHeader', label: 'Верхний колонтитул', run: () => openHf('header') },
  { id: 'hf.editFooter', label: 'Нижний колонтитул', run: () => openHf('footer') },
  { id: 'hf.close', label: 'Закрыть колонтитулы', run: () => closeHf() },
  { id: 'hf.page', label: 'Номер страницы (поле)', run: () => dispatchCmd(insertField('page')) },
  { id: 'hf.pages', label: 'Число страниц (поле)', run: () => dispatchCmd(insertField('pages')) },
  { id: 'hf.differentFirst', label: 'Особый колонтитул первой страницы', checked: () => meta.hf.differentFirst, run: () => { closeHf(); meta.hf.differentFirst = !meta.hf.differentFirst; refreshHf(); updateUi(); scheduleRecovery(); } },
  { id: 'insert.pageNumber', label: 'Номер страницы', run: (v) => { if (v) insertPageNumber(v); setSelect('insert.pageNumber', ''); } },
  { id: 'file.new', label: 'Создать', keys: ['Ctrl+KeyN'], run: () => (desktop.isDesktop ? void desktop.newWindow() : void guardUnsaved(newDocument)) },
  { id: 'file.open', label: 'Открыть', keys: ['Ctrl+KeyO'], run: () => (desktop.isDesktop ? void desktop.openDialog() : void guardUnsaved(() => $('file-input').click())) },
  { id: 'file.save', label: 'Сохранить', keys: ['Ctrl+KeyS'], run: () => void save() },
  { id: 'file.saveAs', label: 'Сохранить как', keys: ['Ctrl+Shift+KeyS'], run: () => void save(true) },
  { id: 'file.history', label: 'История версий', run: () => showHistory() },
  { id: 'file.print', label: 'Печать', keys: ['Ctrl+KeyP'], run: () => void printDocument() },
);

/** Схема многоуровневого списка у выделения (самого внешнего списка). */
function currentScheme(): Scheme | null {
  const { $from } = active().state.selection;
  for (let d = 1; d <= $from.depth; d++) {
    const n = $from.node(d);
    if (n.type === schema.nodes.ordered_list || n.type === schema.nodes.bullet_list) return (n.attrs.scheme as Scheme | null) ?? null;
  }
  return null;
}

function startAtDialog(): void {
  const cur = currentStart(active().state);
  if (cur == null) { toast('Поставьте курсор в нумерованный список'); return; }
  dialog((box, close) => {
    const n = el('input', { type: 'number', value: String(cur), min: '1' });
    const go = () => { close(); setListStart(+n.value || 1)(active().state, active().dispatch); active().focus(); };
    box.addEventListener('keydown', (e) => { if (e.key === 'Enter') go(); });
    box.append(el('h2', { textContent: 'Начать нумерацию с' }), el('div', { class: 'dlg-row' }, el('label', { textContent: 'Номер' }), n),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), el('button', { class: 'primary', textContent: 'Применить', onclick: go })));
  });
}

const inspectorSource = () => ({ view: active(), defaults: meta.defaults });
const inspectorActions = {
  resetPara: () => { dispatchCmd(setParaAttrs({ align: null, indent: null, first: null, before: null, after: null, line: null })); },
  resetChars: () => { dispatchCmd(clearFormatting); },
};

/** Помощник проверки оформления: действия над документом, страницей и колонтитулами. */
const checkHost: CheckHost = {
  state: () => view.state,
  dispatch: (tr) => view.dispatch(tr),
  meta: () => meta,
  changePage: (p) => changePage(p),
  setDefaults: (d) => { Object.assign(meta.defaults, d); applyPage(); replan(view); updateUi(); },
  pageNumbers: (where, align, skipFirst) => {
    closeHf();
    const hasField = (j: HfDoc) => !!j && JSON.stringify(j).includes('"type":"field"');
    const para = schema.nodes.paragraph.create({ align: align === 'left' ? null : align }, schema.nodes.field.create({ kind: 'page' }));
    const mk = (): HfDoc => schema.nodes.doc.create(null, [para]).toJSON() as HfDoc;
    const other = where === 'header' ? 'footer' : 'header';
    if (!hasField(where === 'header' ? meta.hf.header : meta.hf.footer)) setHf(meta.hf, where, false, mk());
    void other;
    if (skipFirst) {
      meta.hf.differentFirst = true;
      const first = where === 'header' ? meta.hf.firstHeader : meta.hf.firstFooter;
      if (hasField(first)) setHf(meta.hf, where, true, null);
    }
    refreshHf();
    updateUi();
    scheduleRecovery();
  },
};

function flagsDialog(): void {
  dialog((box, close) => {
    const flags = listFlags(view.state.doc);
    box.append(el('h2', { textContent: 'Метки в документе' }),
      el('p', { textContent: 'Красная — проверить, зелёная — готово. Метки ставит двойной щелчок на левом поле (Shift — зелёная) или Ctrl+Alt+R / Ctrl+Alt+G. В файл .docx и в печать метки не попадают.' }));
    if (!flags.length) box.append(el('p', { textContent: 'Меток нет.' }));
    for (const f of flags) {
      box.append(el('button', {
        class: 'flag-row', textContent: `${f.flag === 'red' ? '🔴' : '🟢'}  ${f.text || '(пустой абзац)'}`,
        onclick: () => { close(); view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, f.pos + 1)).scrollIntoView()); view.focus(); },
      }));
    }
    box.append(el('div', { class: 'dlg-actions' }, el('button', { class: 'primary', textContent: 'Закрыть', onclick: close })));
  });
}

/** Поиск команд в заголовке окна (Alt+Q): по-русски и по-английски, с терпимостью к опечаткам — общий поиск с ENS-Table. */
function initCommandSearch(): void {
  const input = $('search') as HTMLInputElement;
  const box = document.createElement('div');
  box.className = 'search-results';
  box.hidden = true;
  document.body.append(box);
  let hits: { id: string; label: string; keys?: string }[] = [];
  let on = 0;
  const render = () => {
    box.replaceChildren(...hits.map((h, i) => {
      const b = document.createElement('button');
      b.className = i === on ? 'on' : '';
      const l = document.createElement('span');
      l.textContent = h.label;
      const k = document.createElement('span');
      k.className = 'k';
      k.textContent = h.keys ?? '';
      b.append(l, k);
      b.addEventListener('mousedown', (e) => { e.preventDefault(); choose(i); });
      return b;
    }));
    box.hidden = !hits.length;
    const r = input.getBoundingClientRect();
    box.style.left = `${r.left}px`;
    box.style.top = `${r.bottom + 4}px`;
  };
  const choose = (i: number) => {
    const h = hits[i];
    box.hidden = true;
    input.value = '';
    input.blur();
    if (h) setTimeout(() => run(h.id), 0);
  };
  input.addEventListener('input', () => {
    const q = input.value.trim();
    hits = q ? allCommands()
      .map((c) => ({ c, score: Math.max(matchScore(q, c.label), ...(c.keywords ?? []).map((k) => matchScore(q, k) * 0.9)) }))
      .filter((x) => x.score > 0 && !x.c.id.includes(':menu'))
      .sort((a, b) => b.score - a.score)
      .slice(0, 9)
      .map(({ c }) => ({ id: c.id, label: c.label, keys: c.keys?.[0]?.replace(/Key|Digit/g, '').replaceAll('+', '+') })) : [];
    on = 0;
    render();
  });
  input.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown') { on = Math.min(hits.length - 1, on + 1); render(); e.preventDefault(); }
    else if (e.key === 'ArrowUp') { on = Math.max(0, on - 1); render(); e.preventDefault(); }
    else if (e.key === 'Enter') { choose(on); e.preventDefault(); }
    else if (e.key === 'Escape') { box.hidden = true; input.value = ''; input.blur(); view.focus(); }
  });
  input.addEventListener('blur', () => setTimeout(() => { box.hidden = true; }, 150));
  register({ id: 'search.focus', label: 'Поиск команд', keys: ['Alt+KeyQ'], run: () => { input.focus(); input.select(); } });
}

// ---------------------------------------------------------------- вкладки ленты

const PAGE_SIZES = [
  { value: 'a4', label: 'A4' }, { value: 'a5', label: 'A5' }, { value: 'a3', label: 'A3' },
  { value: 'letter', label: 'Letter' }, { value: 'legal', label: 'Legal' },
];
const MARGINS = [
  { value: 'normal', label: 'Поля: обычные' }, { value: 'narrow', label: 'Поля: узкие' },
  { value: 'moderate', label: 'Поля: средние' }, { value: 'wide', label: 'Поля: широкие' },
];
const LINES = [
  { value: '1', label: 'Интервал 1,0' }, { value: '1.15', label: 'Интервал 1,15' }, { value: '1.5', label: 'Интервал 1,5' }, { value: '2', label: 'Интервал 2,0' },
];
const STYLES = [
  { value: '0', label: 'Обычный' }, { value: '1', label: 'Заголовок 1' }, { value: '2', label: 'Заголовок 2' }, { value: '3', label: 'Заголовок 3' },
  { value: '4', label: 'Заголовок 4' }, { value: '5', label: 'Заголовок 5' }, { value: '6', label: 'Заголовок 6' },
];

const PAGE_NUMBER_POSITIONS = [
  { value: '', label: 'Номер страницы…' },
  { value: 'footer-center', label: 'Внизу, по центру' }, { value: 'footer-right', label: 'Внизу, справа' }, { value: 'footer-left', label: 'Внизу, слева' },
  { value: 'header-center', label: 'Вверху, по центру' }, { value: 'header-right', label: 'Вверху, справа' }, { value: 'header-left', label: 'Вверху, слева' },
];

const TABS: RibbonTab[] = [
  {
    id: 'home', label: 'Главная',
    groups: [
      { name: 'Отмена', items: [{ type: 'medium', command: 'edit.undo', icon: 'arrow_undo' }, { type: 'medium', command: 'edit.redo', icon: 'arrow_redo' }] },
      {
        name: 'Буфер обмена',
        items: [
          { type: 'large', command: 'clipboard.paste', icon: 'clipboard_paste', label: 'Вставить' },
          { type: 'small', command: 'clipboard.cut', icon: 'cut' },
          { type: 'small', command: 'clipboard.copy', icon: 'copy' },
        ],
      },
      {
        name: 'Шрифт',
        items: [
          { type: 'select', command: 'format.fontName', options: FONTS, width: 150 },
          { type: 'select', command: 'format.fontSize', options: FONT_SIZES, width: 56 },
          { type: 'small', command: 'format.bold', icon: 'text_bold' },
          { type: 'small', command: 'format.italic', icon: 'text_italic' },
          { type: 'small', command: 'format.underline', icon: 'text_underline' },
          { type: 'small', command: 'format.strike', icon: 'text_strike' },
          { type: 'small', command: 'format.subscript', icon: 'text_sub' },
          { type: 'small', command: 'format.superscript', icon: 'text_sup' },
          { type: 'color', command: 'format.highlight', icon: 'highlight', kind: 'fill' },
          { type: 'color', command: 'format.fontColor', icon: 'text_color', kind: 'font' },
          { type: 'small', command: 'format.clear', icon: 'clear_format' },
        ],
      },
      {
        name: 'Абзац',
        items: [
          { type: 'small', command: 'list.bullet', icon: 'text_bullet_list' },
          { type: 'small', command: 'list.ordered', icon: 'text_number_list' },
          { type: 'small', command: 'indent.dec', icon: 'text_indent_less' },
          { type: 'small', command: 'indent.inc', icon: 'text_indent_more' },
          { type: 'select', command: 'para.line', options: LINES, width: 120 },
          { type: 'small', command: 'align.left', icon: 'text_align_left' },
          { type: 'small', command: 'align.center', icon: 'text_align_center' },
          { type: 'small', command: 'align.right', icon: 'text_align_right' },
          { type: 'small', command: 'align.justify', icon: 'text_align_justify' },
        ],
      },
      { name: 'Стили', items: [{ type: 'select', command: 'para.style', options: STYLES, width: 130 }] },
      {
        name: 'Нумерация',
        items: [
          { type: 'select', command: 'list.scheme', width: 190, options: [{ value: '', label: 'Схема нумерации…' }, ...PRESETS.map((p) => ({ value: p.id, label: p.name })), { value: 'custom', label: 'Настроить…' }] },
          { type: 'medium', command: 'list.restart', icon: 'arrow_counterclockwise' },
          { type: 'medium', command: 'list.startAt', icon: 'text_number_list' },
          { type: 'medium', command: 'list.continue', icon: 'text_number_list_ltr_rotate_270' },
          { type: 'medium', command: 'list.fromManual', icon: 'wand' },
        ],
      },
      {
        name: 'Редактирование',
        items: [
          { type: 'medium', command: 'edit.find', icon: 'search' },
          { type: 'medium', command: 'edit.replace', icon: 'arrow_sort' },
        ],
      },
    ],
  },
  {
    id: 'insert', label: 'Вставка',
    groups: [
      {
        name: 'Страницы',
        items: [{ type: 'large', command: 'insert.pageBreak', icon: 'page_break', label: 'Разрыв страницы' }],
      },
      {
        name: 'Таблицы и рисунки',
        items: [
          { type: 'large', command: 'insert.table', icon: 'table', label: 'Таблица' },
          { type: 'large', command: 'insert.image', icon: 'image', label: 'Рисунок' },
        ],
      },
      { name: 'Ссылки', items: [{ type: 'large', command: 'insert.link', icon: 'link', label: 'Ссылка' }] },
      {
        name: 'Колонтитулы',
        items: [
          { type: 'large', command: 'hf.editHeader', icon: 'header', label: 'Верхний колонтитул' },
          { type: 'large', command: 'hf.editFooter', icon: 'footer', label: 'Нижний колонтитул' },
          { type: 'select', command: 'insert.pageNumber', width: 150, options: PAGE_NUMBER_POSITIONS },
        ],
      },
    ],
  },
  {
    id: 'layout', label: 'Макет',
    groups: [
      {
        name: 'Параметры страницы',
        items: [
          { type: 'select', command: 'page.size', options: PAGE_SIZES, width: 90 },
          { type: 'select', command: 'page.orientation', options: [{ value: 'portrait', label: 'Книжная' }, { value: 'landscape', label: 'Альбомная' }], width: 100 },
          { type: 'select', command: 'page.margins', options: MARGINS, width: 150 },
        ],
      },
      {
        name: 'Абзац',
        items: [{ type: 'select', command: 'para.after', options: [{ value: '', label: 'После абзаца: как в документе' }, { value: '0', label: 'После: 0 пт' }, { value: '6', label: 'После: 6 пт' }, { value: '8', label: 'После: 8 пт' }, { value: '12', label: 'После: 12 пт' }], width: 210 }],
      },
    ],
  },
  {
    id: 'hf', label: 'Колонтитулы',
    groups: [
      {
        name: 'Колонтитулы',
        items: [
          { type: 'large', command: 'hf.editHeader', icon: 'header', label: 'Верхний колонтитул' },
          { type: 'large', command: 'hf.editFooter', icon: 'footer', label: 'Нижний колонтитул' },
        ],
      },
      {
        name: 'Вставка',
        items: [
          { type: 'large', command: 'hf.page', icon: 'page_number', label: 'Номер страницы' },
          { type: 'large', command: 'hf.pages', icon: 'page_count', label: 'Число страниц' },
          { type: 'select', command: 'insert.pageNumber', width: 150, options: PAGE_NUMBER_POSITIONS },
        ],
      },
      {
        name: 'Параметры',
        items: [{ type: 'large', command: 'hf.differentFirst', icon: 'document_text', label: 'Особый для первой страницы' }],
      },
      {
        name: 'Закрытие',
        items: [{ type: 'large', command: 'hf.close', icon: 'close_circle', label: 'Закрыть колонтитулы' }],
      },
    ],
  },
  {
    id: 'review', label: 'Проверка',
    groups: [
      {
        name: 'Типограф',
        items: [
          { type: 'large', command: 'typo.apply', icon: 'wand', label: 'Типограф' },
          { type: 'large', command: 'typo.live', icon: 'text_font', label: 'Кавычки при наборе' },
        ],
      },
      {
        name: 'Метки',
        items: [
          { type: 'large', command: 'mark.red', icon: 'info', label: 'Проверить' },
          { type: 'large', command: 'mark.green', icon: 'checkmark', label: 'Готово' },
          { type: 'large', command: 'mark.next', icon: 'arrow_redo', label: 'Следующая' },
          { type: 'large', command: 'mark.list', icon: 'book_open', label: 'Список меток' },
        ],
      },
      {
        name: 'Оформление',
        items: [
          { type: 'large', command: 'review.check', icon: 'checkmark', label: 'Проверка по ГОСТ' },
          { type: 'large', command: 'review.inspector', icon: 'search', label: 'Рентген формата' },
        ],
      },
      { name: 'Чтение', items: [{ type: 'large', command: 'view.focusPara', icon: 'target', label: 'Фокус абзаца' }] },
    ],
  },
  {
    id: 'versions', label: 'Версии',
    groups: [
      {
        name: 'Версии',
        items: [
          { type: 'large', command: 'ver.history', icon: 'history', label: 'История версий' },
          { type: 'large', command: 'ver.named', icon: 'save_edit', label: 'Записать версию' },
        ],
      },
      { name: 'Копии на диске', items: [{ type: 'large', command: 'ver.disk', icon: 'folder_open', label: 'Копии на диске' }] },
      {
        name: 'Защита от потерь',
        items: [{
          type: 'select', command: 'ver.reminder', width: 210,
          options: [{ value: '5', label: 'Напоминать о сохранении: 5 мин' }, { value: '10', label: 'Напоминать: 10 мин' }, { value: '20', label: 'Напоминать: 20 мин' }, { value: '0', label: 'Не напоминать' }],
        }],
      },
    ],
  },
  {
    id: 'table', label: 'Таблица',
    groups: [
      {
        name: 'Строки и столбцы',
        items: [
          { type: 'large', command: 'table.rowBefore', icon: 'row_add', label: 'Строка выше' },
          { type: 'large', command: 'table.rowAfter', icon: 'row_add', label: 'Строка ниже' },
          { type: 'large', command: 'table.colBefore', icon: 'col_add', label: 'Столбец слева' },
          { type: 'large', command: 'table.colAfter', icon: 'col_add', label: 'Столбец справа' },
          { type: 'large', command: 'table.delRow', icon: 'row_del', label: 'Удалить строку' },
          { type: 'large', command: 'table.delCol', icon: 'col_del', label: 'Удалить столбец' },
        ],
      },
      {
        name: 'Ячейки',
        items: [
          { type: 'large', command: 'table.merge', icon: 'merge', label: 'Объединить' },
          { type: 'large', command: 'table.split', icon: 'cell_split', label: 'Разделить' },
          { type: 'large', command: 'table.borders', icon: 'border_all', label: 'Границы' },
          { type: 'large', command: 'table.delete', icon: 'eraser', label: 'Удалить таблицу' },
        ],
      },
    ],
  },
  {
    id: 'view', label: 'Вид',
    groups: [
      {
        name: 'Масштаб',
        items: [
          { type: 'large', command: 'view.zoomIn', icon: 'zoom_in', label: 'Увеличить' },
          { type: 'large', command: 'view.zoomOut', icon: 'zoom_in', label: 'Уменьшить' },
          { type: 'large', command: 'view.zoomReset', icon: 'document_text', label: '100%' },
        ],
      },
      { name: 'Оформление', items: [{ type: 'large', command: 'view.theme', icon: 'paint_brush', label: 'Тема' }] },
    ],
  },
];

const refreshRibbon = renderRibbon($('tabs'), $('ribbon'), () => fileMenu(), TABS);

function setSelect(cmd: string, value: string): void {
  const sel = document.querySelector<HTMLSelectElement>(`select[data-command="${cmd}"]`);
  if (!sel || sel.value === value) return;
  if (![...sel.options].some((o) => o.value === value)) sel.add(new Option(value, value));
  sel.value = value;
}

const TABLE_CMDS = new Set(['table.rowBefore', 'table.rowAfter', 'table.colBefore', 'table.colAfter', 'table.delRow', 'table.delCol', 'table.merge', 'table.split', 'table.borders', 'table.delete']);

function updateUi(): void {
  refreshRibbon();
  const dirty = isDirty();
  if (dirty && !dirtySince) dirtySince = Date.now();
  if (!dirty) dirtySince = 0;
  updateSavedStatus();
  $('doc-name').textContent = `${docName}${dirty ? ' •' : ''} — ENS-Text`;
  document.title = `${docName}${dirty ? '*' : ''} — ENS-Text`;
  const canUndo = undoDepth(active().state) > 0;
  const canRedo = redoDepth(active().state) > 0;
  ($('qa-undo') as HTMLButtonElement).disabled = !canUndo;
  ($('qa-redo') as HTMLButtonElement).disabled = !canRedo;
  document.querySelector('.btn[data-command="edit.undo"]')?.classList.toggle('disabled', !canUndo);
  document.querySelector('.btn[data-command="edit.redo"]')?.classList.toggle('disabled', !canRedo);
  const inTable = isInTable(active().state);
  document.querySelectorAll<HTMLElement>('.btn[data-command]').forEach((b) => {
    if (TABLE_CMDS.has(b.dataset.command!)) b.classList.toggle('disabled', !inTable);
  });
  const st = currentStyle(active().state);
  setSelect('format.fontName', st.font ?? meta.defaults.font);
  setSelect('format.fontSize', String(st.size ?? meta.defaults.size));
  setSelect('para.style', String(currentLevel(active().state)));
  setSelect('para.line', String(paraAttr<number>(active().state, 'line') ?? meta.defaults.line));
  setSelect('page.orientation', meta.page.w > meta.page.h ? 'landscape' : 'portrait');
  setSelect('ver.reminder', String(settings.saveReminderMinutes));
  setSelect('list.scheme', presetOf(currentScheme()) === 'custom' ? '' : presetOf(currentScheme()) === 'simple' ? '' : presetOf(currentScheme()));
  setSelect('page.size', sizeKey());
  setSelect('page.margins', marginsKey());
}

// --- Страница: размеры и поля

const SIZES: Record<string, [number, number]> = { a4: [595.3, 841.9], a5: [419.5, 595.3], a3: [841.9, 1190.6], letter: [612, 792], legal: [612, 1008] };
const cm = (v: number) => Math.round(v * 28.3465 * 10) / 10;
const MARGIN_SETS: Record<string, [number, number, number, number]> = {
  normal: [cm(2), cm(2), cm(3), cm(1.5)], narrow: [cm(1.27), cm(1.27), cm(1.27), cm(1.27)],
  moderate: [cm(2.54), cm(2.54), cm(1.91), cm(1.91)], wide: [cm(2.54), cm(2.54), cm(5.08), cm(5.08)],
};

function sizeKey(): string {
  const w = Math.min(meta.page.w, meta.page.h);
  const h = Math.max(meta.page.w, meta.page.h);
  for (const [k, [sw, sh]] of Object.entries(SIZES)) if (Math.abs(sw - w) < 2 && Math.abs(sh - h) < 2) return k;
  return 'a4';
}

function marginsKey(): string {
  const p = meta.page;
  for (const [k, [t, b, l, r]] of Object.entries(MARGIN_SETS)) {
    if ([p.top - t, p.bottom - b, p.left - l, p.right - r].every((d) => Math.abs(d) < 1.5)) return k;
  }
  return 'normal';
}

function changePage(next: Partial<PageSetup>): void {
  meta.page = { ...meta.page, ...next };
  hfPx = { header: measureHf(meta.hf, meta.page, 'header'), footer: measureHf(meta.hf, meta.page, 'footer') };
  applyPage();
  replan(view);
  updateUi();
  scheduleRecovery();
}

function setPageSize(key: string): void {
  const [w, h] = SIZES[key] ?? SIZES.a4;
  const land = meta.page.w > meta.page.h;
  changePage(land ? { w: h, h: w } : { w, h });
}

function setOrientation(landscape: boolean): void {
  const lo = Math.min(meta.page.w, meta.page.h);
  const hi = Math.max(meta.page.w, meta.page.h);
  changePage(landscape ? { w: hi, h: lo } : { w: lo, h: hi });
}

function setMargins(key: string): void {
  const [top, bottom, left, right] = MARGIN_SETS[key] ?? MARGIN_SETS.normal;
  changePage({ top, bottom, left, right });
}

// ---------------------------------------------------------------- окна

function tableDialog(): void {
  dialog((box, close) => {
    const rows = el('input', { type: 'number', value: '3', min: '1', max: '100' });
    const cols = el('input', { type: 'number', value: '3', min: '1', max: '30' });
    const borders = el('input', { type: 'checkbox', checked: true });
    const ok = el('button', { class: 'primary', textContent: 'Вставить' });
    const go = () => {
      const r = Math.min(100, Math.max(1, +rows.value || 1));
      const c = Math.min(30, Math.max(1, +cols.value || 1));
      close();
      dispatchCmd(insertTable(r, c, borders.checked));
    };
    ok.addEventListener('click', go);
    box.addEventListener('keydown', (e) => { if (e.key === 'Enter') go(); });
    box.append(
      el('h2', { textContent: 'Вставка таблицы' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Строк' }), rows),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Столбцов' }), cols),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'С границами' }), borders),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), ok),
    );
  });
}

function linkDialog(): void {
  dialog((box, close) => {
    const url = el('input', { type: 'text', value: linkAt(active().state) ?? 'https://', placeholder: 'Адрес ссылки' });
    const ok = el('button', { class: 'primary', textContent: 'Применить' });
    const go = () => {
      const v = url.value.trim();
      close();
      dispatchCmd(setLink(v === 'https://' ? '' : v));
    };
    ok.addEventListener('click', go);
    box.addEventListener('keydown', (e) => { if (e.key === 'Enter') go(); });
    box.append(
      el('h2', { textContent: 'Гиперссылка' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Адрес' }), url),
      el('p', { textContent: 'Пустой адрес убирает ссылку с выделенного текста. Открыть ссылку: Ctrl+щелчок.' }),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), ok),
    );
  });
}

function toggleBorders(): void {
  const { $from } = active().state.selection;
  for (let d = $from.depth; d > 0; d--) {
    const n = $from.node(d);
    if (n.type === schema.nodes.table) {
      active().dispatch(active().state.tr.setNodeMarkup($from.before(d), undefined, { ...n.attrs, borders: !n.attrs.borders }));
      active().focus();
      return;
    }
  }
}

// ---------------------------------------------------------------- рисунки

async function insertImageFile(file: File): Promise<void> {
  const src = await new Promise<string>((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result as string);
    r.onerror = () => rej(r.error);
    r.readAsDataURL(file);
  });
  const dim = await new Promise<{ w: number; h: number }>((res) => {
    const img = new Image();
    img.onload = () => res({ w: img.naturalWidth || 200, h: img.naturalHeight || 150 });
    img.onerror = () => res({ w: 200, h: 150 });
    img.src = src;
  });
  // Не шире рабочей области страницы (96 dpi: 1 pt = 4/3 px).
  const maxW = ((meta.page.w - meta.page.left - meta.page.right) * 4) / 3;
  const k = dim.w > maxW ? maxW / dim.w : 1;
  dispatchCmd(insertImage(src, Math.round(dim.w * k), Math.round(dim.h * k), file.name.replace(/\.[^.]+$/, '')));
}

$('image-input').addEventListener('change', (e) => {
  const input = e.target as HTMLInputElement;
  const f = input.files?.[0];
  input.value = '';
  if (f) void insertImageFile(f);
});

// ---------------------------------------------------------------- колонтитулы

function openHf(kind: HfKind, pageIndex = currentPage() - 1): void {
  closeHf();
  const first = meta.hf.differentFirst && pageIndex === 0;
  const e = effPage();
  hfSession = openHfEditor({
    sheet: $('sheet'), kind, first, pageIndex, pageStep: geometry(e).P, page: e, hf: meta.hf,
    onChange: () => { refreshHf(); updateUi(); scheduleRecovery(); },
  });
  $('sheet').classList.add('hf-mode');
  view.setProps({ editable: () => !hfSession });
  paintHf();
  document.querySelectorAll<HTMLButtonElement>('#tabs .tab').forEach((t) => { if (t.textContent === 'Колонтитулы') t.click(); });
  updateUi();
}

function closeHf(): void {
  if (!hfSession) return;
  hfSession.close();
  hfSession = null;
  $('sheet').classList.remove('hf-mode');
  view.setProps({ editable: () => !hfSession });
  refreshHf();
  document.querySelectorAll<HTMLButtonElement>('#tabs .tab').forEach((t) => { if (t.textContent === 'Главная') t.click(); });
  view.focus();
  updateUi();
}

/** «Вставка → Номер страницы»: абзац с полем в нужном месте листа. */
function insertPageNumber(pos: string): void {
  closeHf();
  const [where, align] = pos.split('-');
  const kind: HfKind = where === 'header' ? 'header' : 'footer';
  const para = schema.nodes.paragraph.create({ align: align === 'left' ? null : align }, schema.nodes.field.create({ kind: 'page' }));
  const cur = kind === 'header' ? meta.hf.header : meta.hf.footer;
  const old = cur ? schema.nodeFromJSON(cur) : null;
  const blank = !old || (old.childCount === 1 && old.firstChild!.content.size === 0);
  const doc = schema.nodes.doc.create(null, blank ? [para] : [...old.content.content, para]);
  setHf(meta.hf, kind, false, doc.toJSON() as HfDoc);
  refreshHf();
  updateUi();
  scheduleRecovery();
  toast(`Номер страницы добавлен ${kind === 'header' ? 'в верхний' : 'в нижний'} колонтитул`);
}

/** Двойной щелчок по верхнему или нижнему полю листа открывает колонтитул; по тексту — закрывает. */
$('sheet').addEventListener('dblclick', (e) => {
  const rect = $('sheet').getBoundingClientRect();
  const scale = rect.width / ($('sheet') as HTMLElement).offsetWidth || 1;
  const eff = effPage();
  const g = geometry(eff);
  const y = (e.clientY - rect.top) / scale;
  const i = Math.floor(y / g.P);
  const inPage = y - i * g.P;
  if (i < 0 || i >= pageCount(view) || inPage > g.H) return;
  if (hfSession && (e.target as HTMLElement).closest('.hf-edit')) return;
  if (inPage < eff.top * PX - 4) { e.preventDefault(); openHf('header', i); }
  else if (inPage > g.H - eff.bottom * PX + 4) { e.preventDefault(); openHf('footer', i); }
  else if (hfSession) closeHf();
  else if (e.clientX < rect.left + eff.left * PX * scale - 2 && e.clientX > rect.left + i * 0) {
    // Двойной щелчок на левом поле листа — метка у этого абзаца (Shift — зелёная, как в таблицах).
    const at = view.posAtCoords({ left: rect.left + (eff.left * PX + 12) * scale, top: e.clientY });
    if (!at) return;
    e.preventDefault();
    view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(at.pos))));
    dispatchCmd(toggleFlag(e.shiftKey ? 'green' : 'red'));
  }
});
document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && hfSession && document.getElementById('dialog')!.hidden) closeHf(); });

// ---------------------------------------------------------------- печать

/** Печать по нашей вёрстке: страницы вплотную, поля внутри страницы, колонтитулы на каждом листе — как на экране. */
async function printPrepare(): Promise<void> {
  closeHf();
  layout.gap = 0;
  replan(view);
  applyPage();
  await new Promise((r) => setTimeout(r, pagStats.items > 600 ? 400 : 120));
}

function printRestore(): void {
  if (layout.gap === 16) return;
  layout.gap = 16;
  replan(view);
  applyPage();
}

async function printDocument(): Promise<void> {
  await printPrepare();
  window.addEventListener('afterprint', printRestore, { once: true });
  try {
    window.print();
  } finally {
    // Chromium (в том числе WebView2) возвращается из print() после закрытия окна печати; остальные — по afterprint.
    if (navigator.userAgent.includes('Chrome')) printRestore();
  }
}

// ---------------------------------------------------------------- поиск и замена

const findbar = $('findbar');

function openFind(withReplace: boolean): void {
  findbar.hidden = false;
  findbar.replaceChildren();
  if (!document.getElementById('fmt-fonts')) document.body.append(el('datalist', { id: 'fmt-fonts' }, ...OFFICE_FONTS.map((f) => el('option', { value: f }))));
  const q = el('input', { type: 'text', placeholder: 'Найти', value: findKey.getState(view.state)?.query ?? '' });
  const r = el('input', { type: 'text', placeholder: 'Заменить на' });
  const mc = el('input', { type: 'checkbox' });
  const count = el('span', { class: 'count' });
  // Условие по оформлению («Формат»): шрифт, кегль, Ж/К/Ч/зачёркнутый; и оформление для замены.
  const fmtRow = (title: string, withStrike: boolean) => {
    const font = el('input', { type: 'text', placeholder: 'Шрифт', title: 'Шрифт (любой, если пусто)'});
    font.setAttribute('list', 'fmt-fonts');
    const size = el('input', { type: 'number', placeholder: 'Кегль', min: '1', max: '400', step: '0.5' });
    const flags = (['strong', 'em', 'underline', ...(withStrike ? ['strike'] : [])] as const).map((k) => {
      const c = el('input', { type: 'checkbox' });
      return { k, c, label: el('label', { class: 'fmtflag' }, c, ({ strong: 'Ж', em: 'К', underline: 'Ч', strike: 'З' } as Record<string, string>)[k]) };
    });
    const get = (): Record<string, string | number | boolean> => {
      const o: Record<string, string | number | boolean> = {};
      if (font.value.trim()) o.font = font.value.trim();
      if (parseFloat(size.value) > 0) o.size = parseFloat(size.value);
      for (const f of flags) if (f.c.checked) o[f.k] = true;
      return o;
    };
    const row = el('div', { class: 'fmtrow' }, el('span', { class: 'fmttitle', textContent: title }), font, size, ...flags.map((f) => f.label));
    return { row, get, inputs: [font, size, ...flags.map((f) => f.c)] };
  };
  const sf = fmtRow('Искать с оформлением:', true);
  const rf = fmtRow('Заменить оформлением:', false);
  const fmtBox = el('div', { class: 'fmtbox' }, sf.row, ...(withReplace ? [rf.row] : []));
  fmtBox.hidden = true;
  const sync = () => {
    view.dispatch(view.state.tr.setMeta(findKey, { query: q.value, matchCase: mc.checked, fmt: fmtActive(sf.get()) ? sf.get() : null, current: -1 }));
    const s = findKey.getState(view.state)!;
    count.textContent = q.value || s.fmt ? `Найдено: ${s.matches.length}` : '';
  };
  q.addEventListener('input', sync);
  mc.addEventListener('change', sync);
  for (const i of sf.inputs) { i.addEventListener('input', sync); i.addEventListener('change', sync); }
  q.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); findStep(view, e.shiftKey ? -1 : 1); }
    if (e.key === 'Escape') closeFind();
  });
  r.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeFind(); });
  const replaceOne = () => {
    const s = findKey.getState(view.state)!;
    let idx = s.current;
    const sel = view.state.selection;
    if (idx < 0 || s.matches[idx]?.from !== sel.from || s.matches[idx]?.to !== sel.to) { findStep(view, 1); return; }
    replaceMatch(view, idx, r.value, rf.get());
    sync();
    findStep(view, 1);
    idx = -1;
  };
  const buttons = [
    el('button', { textContent: '↑', title: 'Предыдущее (Shift+Enter)', onclick: () => findStep(view, -1) }),
    el('button', { textContent: '↓', title: 'Следующее (Enter)', onclick: () => findStep(view, 1) }),
  ];
  const fmtBtn = el('button', { textContent: 'Формат', title: 'Искать и заменять по оформлению: шрифт, кегль, Ж/К/Ч', onclick: () => { fmtBox.hidden = !fmtBox.hidden; fmtBtn.classList.toggle('on', !fmtBox.hidden); } });
  const kids: (Node | string)[] = [q, ...buttons, el('label', {}, mc, ' Учитывать регистр'), fmtBtn, count];
  if (withReplace) {
    kids.push(r, el('button', { textContent: 'Заменить', onclick: replaceOne }), el('button', {
      textContent: 'Заменить все',
      onclick: () => { if (findKey.getState(view.state)?.matches.length) snap('before', `Перед «Заменить все» (${q.value} → ${r.value})`); const n = replaceAll(view, r.value, rf.get()); sync(); toast(`Заменено: ${n}`); },
    }));
  }
  kids.push(el('button', { textContent: '✕', title: 'Закрыть (Esc)', onclick: closeFind }));
  findbar.append(...kids, fmtBox);
  const sel = view.state.selection;
  if (!sel.empty && sel.$from.sameParent(sel.$to)) q.value = view.state.doc.textBetween(sel.from, sel.to);
  sync();
  q.focus();
  q.select();
}

function closeFind(): void {
  findbar.hidden = true;
  view.dispatch(view.state.tr.setMeta(findKey, { query: '', fmt: null, current: -1 }));
  view.focus();
}

// ---------------------------------------------------------------- файлы

function showNotice(text: string, actions: { label: string; run: () => void }[] = []): void {
  const n = $('notice');
  n.replaceChildren(el('span', { textContent: text }), ...actions.map((a) => el('button', { textContent: a.label, onclick: a.run })));
  n.hidden = false;
}

function hideNotice(): void {
  $('notice').hidden = true;
}

function loadDoc(doc: PMNode, newMetaValue: DocMeta, name: string, dirty = false, hinfo?: { raw: string; key: string; mtime?: number }): void {
  if (hfSession) { hfSession.close(); hfSession = null; $('sheet').classList.remove('hf-mode'); }
  meta = newMetaValue;
  meta.hf ??= newMeta().hf;
  hfPx = { header: measureHf(meta.hf, meta.page, 'header'), footer: measureHf(meta.hf, meta.page, 'footer') };
  docName = name;
  applyPage();
  view.updateState(createState(doc, [paginationPlugin, quotesPlugin]));
  savedDoc = dirty ? emptyDoc() : doc;
  savedMeta = dirty ? '' : metaKey();
  hideNotice();
  dirtySince = 0;
  lastSavedAt = 0;
  const hn = hinfo ? adoptHistory(hinfo.raw, hinfo) : (() => { hist = V.emptyHistory(); histRev++; lastSnapDoc = view.state.doc; lastVersionAt = Date.now(); return null; })();
  const notes: string[] = [];
  const buttons: { label: string; run: () => void }[] = [{ label: 'Понятно', run: hideNotice }];
  if (meta.lossy.length) {
    const tail = desktop.isDesktop
      ? 'Поверх исходного файла сохранять нельзя — «Сохранить» предложит сохранить копию, в ней этих элементов не будет.'
      : 'Оригинал не изменится: «Сохранить» скачивает новую копию, в ней этих элементов не будет.';
    notes.push(`В этой версии не перенесено: ${meta.lossy.join('; ')}. ${tail}`);
  }
  if (hn) { notes.push(hn.text); buttons.unshift(...hn.buttons); }
  if (notes.length) showNotice(notes.join(' '), buttons);
  replan(view);
  updateStats();
  updateUi();
  view.focus();
}

function newDocument(): void {
  void clearRecovery(docId);
  hist = V.emptyHistory();
  loadDoc(emptyDoc(), newMeta(), 'Документ1');
}

/** Если есть несохранённые правки — спрашивает; продолжает, когда можно. */
async function guardUnsaved(next: () => void): Promise<void> {
  if (!isDirty()) { next(); return; }
  dialog((box, close) => {
    box.append(
      el('h2', { textContent: 'Сохранить изменения?' }),
      el('p', { textContent: `В документе «${docName}» есть несохранённые изменения.` }),
      el('div', { class: 'dlg-actions' },
        el('button', { class: 'primary', textContent: 'Сохранить', onclick: async () => { close(); await save(); next(); } }),
        el('button', { class: 'danger', textContent: 'Не сохранять', onclick: () => { close(); next(); } }),
        el('button', { textContent: 'Отмена', onclick: close })),
    );
  });
}

$('file-input').addEventListener('change', async (e) => {
  const input = e.target as HTMLInputElement;
  const f = input.files?.[0];
  input.value = '';
  if (!f) return;
  try {
    const { doc, meta: m, history } = importDocx(new Uint8Array(await f.arrayBuffer()));
    void clearRecovery(docId);
    loadDoc(doc, m, f.name.replace(/\.docx$/i, ''), false, { raw: history, key: f.name, mtime: f.lastModified });
  } catch (err) {
    dialog((box, close) => {
      box.append(
        el('h2', { textContent: 'Не удалось открыть файл' }),
        el('p', { textContent: `${f.name}: ${(err as Error).message}` }),
        el('div', { class: 'dlg-actions' }, el('button', { class: 'primary', textContent: 'Закрыть', onclick: close })),
      );
    });
  }
});

/** Сохранение: .docx собирается целиком и проверяется чтением, прежде чем отдаётся пользователю. */
async function save(asNew = false): Promise<void> {
  if (desktop.isDesktop) { await desktop.save(asNew); return; }
  try {
    let name = docName;
    if (asNew || docName === 'Документ1') {
      const entered = await askName(docName);
      if (!entered) return;
      name = entered;
    }
    const bytes = buildDocx(name);
    const blob = new Blob([bytes as Uint8Array<ArrayBuffer>], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${name}.docx`;
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 10000);
    docName = name;
    savedDoc = view.state.doc;
    savedMeta = metaKey();
    void clearRecovery(docId);
    afterSaved(`${name}.docx`);
    updateUi();
    toast(`Сохранено: ${name}.docx`);
  } catch (err) {
    rollbackSave();
    dialog((box, close) => {
      box.append(
        el('h2', { textContent: 'Не удалось сохранить' }),
        el('p', { textContent: `${(err as Error).message}. Документ остался открытым и не потерян.` }),
        el('div', { class: 'dlg-actions' }, el('button', { class: 'primary', textContent: 'Закрыть', onclick: close })),
      );
    });
  }
}

function askName(current: string): Promise<string | null> {
  return new Promise((resolve) => {
    dialog((box, close) => {
      const name = el('input', { type: 'text', value: current });
      const done = (v: string | null) => { close(); resolve(v); };
      const ok = el('button', { class: 'primary', textContent: 'Сохранить', onclick: () => done(name.value.trim().replace(/[\\/:*?"<>|]/g, '_') || 'Документ') });
      box.addEventListener('keydown', (e) => { if (e.key === 'Enter') ok.click(); if (e.key === 'Escape') resolve(null); });
      box.append(
        el('h2', { textContent: 'Сохранить как' }),
        el('div', { class: 'dlg-row' }, el('label', { textContent: 'Имя файла' }), name, el('span', { textContent: '.docx' })),
        el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: () => done(null) }), ok),
      );
    });
  });
}

function fileMenu(): void {
  dialog((box, close) => {
    const item = (label: string, id: string) => el('button', { textContent: label, onclick: () => { close(); run(id); } });
    box.append(
      el('h2', { textContent: 'Файл' }),
      el('div', { class: 'menu-list' },
        item('Создать', 'file.new'), item('Открыть…', 'file.open'), item('Сохранить', 'file.save'), item('Сохранить как…', 'file.saveAs'), item('История версий…', 'ver.history'), item('Печать…', 'file.print'),
        el('button', { textContent: 'О программе', onclick: () => { close(); about(); } })),
    );
  });
}

function about(): void {
  dialog((box, close) => {
    box.append(
      el('h2', { textContent: `ENS-Text ${APP_VERSION}` }),
      el('p', { textContent: 'Текстовый редактор ENS-Office. Работает без интернета. Формат файлов — .docx.' }),
      el('p', { textContent: 'Версия 0.1 — первый этап: абзацы, заголовки, списки, таблицы, изображения, ссылки, страницы, печать. Колонтитулы, сноски и совместная работа — в планах.' }),
      el('div', { class: 'dlg-actions' }, el('button', { class: 'primary', textContent: 'Закрыть', onclick: close })),
    );
  });
}

// ---------------------------------------------------------------- восстановление

function scheduleRecovery(): void {
  clearTimeout(recoveryTimer);
  recoveryTimer = window.setTimeout(persistRecovery, 1500);
}

function persistRecovery(): void {
  if (desktop.isDesktop) { void desktop.persist(); return; }
  if (!isDirty()) { void clearRecovery(docId); return; }
  void saveRecovery({ id: docId, name: docName, savedAt: Date.now(), doc: view.state.doc.toJSON(), meta, history: histText() });
}

document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') persistRecovery(); });
// В настольной версии вопрос о сохранении и копию для восстановления ведёт desktop.ts.
if (!desktop.isDesktop) {
  window.addEventListener('beforeunload', (e) => {
    if (isDirty()) { persistRecovery(); e.preventDefault(); e.returnValue = ''; }
  });
}

async function offerRecovery(): Promise<void> {
  let list: RecoveryRecord[] = [];
  try { list = await listRecovery(); } catch { return; }
  list = list.filter((r) => r.id !== docId).sort((a, b) => b.savedAt - a.savedAt);
  if (!list.length) return;
  dialog((box, close) => {
    box.append(el('h2', { textContent: 'Восстановление документов' }), el('p', { textContent: 'Найдены документы, закрытые без сохранения. Восстановленная версия откроется в этом окне; исходные файлы не затронуты.' }));
    for (const r of list) {
      const when = new Date(r.savedAt).toLocaleString('ru-RU');
      box.append(el('div', { class: 'rec-item' },
        el('div', { class: 'rec-title', textContent: r.name }),
        el('div', { class: 'rec-meta', textContent: `Последняя копия: ${when}` }),
        el('div', { class: 'rec-buttons' },
          el('button', {
            class: 'primary', textContent: 'Восстановить',
            onclick: () => {
              try {
                const doc = schema.nodeFromJSON(r.doc);
                doc.check();
                close();
                loadDoc(doc, r.meta, `${r.name} (восстановлено)`, true);
                if (r.history) { hist = V.parse(r.history); histRev++; }
                void clearRecovery(r.id);
                persistRecovery();
              } catch {
                toast('Копию прочитать не удалось');
              }
            },
          }),
          el('button', { class: 'danger', textContent: 'Удалить копию', onclick: async () => { await clearRecovery(r.id); close(); void offerRecovery(); } }),
        )));
    }
    box.append(el('div', { class: 'rec-actions' }, el('button', { textContent: 'Решить позже', onclick: close })));
  }, false);
}

// ---------------------------------------------------------------- запуск

$('qa-save').innerHTML = ICONS.save_20 ?? '';
$('qa-print').innerHTML = ICONS.print_20 ?? '';
$('qa-undo').innerHTML = ICONS.arrow_undo_20 ?? '';
$('qa-redo').innerHTML = ICONS.arrow_redo_20 ?? '';
$('qa-save').addEventListener('click', () => run('file.save'));
$('qa-print').addEventListener('click', () => run('file.print'));
$('qa-undo').addEventListener('click', () => run('edit.undo'));
$('qa-redo').addEventListener('click', () => run('edit.redo'));
$('zoom-in').addEventListener('click', () => setZoom(zoom + 0.1));
$('zoom-out').addEventListener('click', () => setZoom(zoom - 0.1));
$('zoom-val').addEventListener('click', () => setZoom(1));

document.addEventListener('keydown', (e) => {
  if (e.isComposing) return;
  handleKey(e);
}, true);
// Масштаб страницы браузера запрещён: Ctrl+колесо меняет масштаб документа.
$('desk').addEventListener('wheel', (e) => {
  if (!e.ctrlKey) return;
  e.preventDefault();
  setZoom(zoom + (e.deltaY < 0 ? 0.1 : -0.1));
}, { passive: false });
window.addEventListener('wheel', (e) => { if (e.ctrlKey) e.preventDefault(); }, { passive: false });

// Выпадающие списки ленты должны отдавать фокус обратно в текст.
$('ribbon').addEventListener('change', () => setTimeout(() => view.focus(), 0));

initThemes();
initCommandSearch();
applyPage();
setZoom(1);
updateStats();
updateUi();

await loadFonts();
applyPage();
setRestoreFocus(() => view.focus());
if (desktop.isDesktop) {
  await desktop.start({
    name: () => docName,
    isDirty,
    isPristine: () => !isDirty() && view.state.doc.childCount === 1 && view.state.doc.firstChild!.content.size === 0,
    docxBytes: buildDocx,
    rollbackSave,
    afterSaved: (key) => afterSaved(key),
    copyBytes: () => { snap('auto'); return exportDocx(view.state.doc, meta, docName, histText()); },
    snapshotJson: () => JSON.stringify({ doc: view.state.doc.toJSON(), meta, name: docName, history: histText() }),
    openBytes: (bytes, name, info) => {
      const { doc, meta: m, history } = importDocx(bytes);
      loadDoc(doc, m, name, false, info ? { raw: history, key: info.key, mtime: info.mtime } : undefined);
      return m.lossy;
    },
    restoreJson: (json, title) => {
      const r = JSON.parse(json) as { doc: Record<string, unknown>; meta: DocMeta; history?: string };
      const doc = schema.nodeFromJSON(r.doc);
      doc.check();
      loadDoc(doc, r.meta, title, true);
      if (r.history) { hist = V.parse(r.history); histRev++; }
    },
    markSaved: (name) => {
      docName = name;
      savedDoc = view.state.doc;
      savedMeta = metaKey();
      updateUi();
      toast(`Сохранено: ${name}.docx`);
    },
  });
} else {
  void offerRecovery();
}

// Для автоматических проверок (Playwright): доступ к редактору из консоли.
if (import.meta.env.DEV || import.meta.env.MODE === 'test') {
  (window as unknown as Record<string, unknown>).ens = {
    view, schema, findState: () => findKey.getState(view.state)!, pageCount: () => pageCount(view), pagStats, geometry: () => geometry(effPage()), PRESETS, hist: () => hist, snap, V, buildDocx, rollbackSave, historyHost, effPage, printPrepare, printRestore, hfSession: () => hfSession, PX, importDocx, exportDocx, loadDoc, getMeta: () => meta, isDirty, run, A4, lastColor,
  };
}
