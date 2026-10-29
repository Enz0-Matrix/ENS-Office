// Окно «История версий»: список версий, сравнение с текущим документом (или с предыдущей версией), возврат всего документа
// или отдельного фрагмента, именные версии, копии на диске (предыдущие сохранения и закрытое без сохранения).

import type { Node as PMNode } from 'prosemirror-model';
import { historyList, isDesktop, type HistoryEntry } from '../../table/src/platform';
import { dialog, el, toast } from './ui';
import {
  blockText, diffBlocks, emptyHistory, nodeFromBlock, storeBlocks, wordDiff,
  type History, type Hunk, type Version, type VersionKind,
} from './versions';
import type { DocMeta } from './docx';

export const KIND_LABEL: Record<VersionKind, string> = {
  save: 'Сохранение', auto: 'Автоверсия', named: 'Именная версия', open: 'При открытии', before: 'Перед действием', outside: 'Файл изменён другой программой',
};

export const fmtWhen = (t: number) => new Date(t).toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', year: '2-digit', hour: '2-digit', minute: '2-digit' });

export interface HistoryHost {
  history(): History;
  doc(): PMNode;
  meta(): DocMeta;
  user(): string;
  docName(): string;
  /** Записать именную версию (с подписью). */
  saveNamed(label: string): void;
  /** Заменить весь документ документом версии (с отменой и страховочной версией). */
  restore(v: Version): void;
  /** Заменить блоки [from,to) текущего документа блоками версии. */
  revert(range: [number, number], nodes: PMNode[]): void;
  /** Открыть копию с диска в новом окне. */
  openCopy(path: string): void;
  /** Показать сразу сравнение с этой версией. */
  focus?: string;
}

const wordsLabel = (n: number) => `${n} сл.`;

export function openHistory(host: HistoryHost, tab: 'versions' | 'disk' = 'versions', select?: string, against: 'current' | 'previous' = 'current'): void {
  dialog((box, close) => {
    box.classList.add('history-modal');
    const body = el('div', { class: 'hist-body' });
    const tabs = el('div', { class: 'hist-tabs' });
    const mk = (id: 'versions' | 'disk', label: string) => el('button', { class: id === tab ? 'on' : '', textContent: label, onclick: () => { close(); openHistory(host, id); } });
    tabs.append(mk('versions', 'Версии'));
    if (isDesktop) tabs.append(mk('disk', 'Копии на диске'));
    box.append(el('h2', { textContent: `История версий — ${host.docName()}` }), tabs, body,
      el('div', { class: 'dlg-actions' }, el('button', { class: 'primary', textContent: 'Закрыть', onclick: close })));
    if (tab === 'disk') void renderDisk(host, body);
    else renderVersions(host, body, select, against);
  }, false);
}

function renderVersions(host: HistoryHost, body: HTMLElement, select?: string, initial: 'current' | 'previous' = 'current'): void {
  const h = host.history();
  // Текущее состояние (в том числе несохранённое) — для сравнения: блоки во временном хранилище.
  const cur = emptyHistory();
  const curIds = storeBlocks(cur, host.doc());
  const latest = h.versions[h.versions.length - 1];
  const list = el('div', { class: 'hist-list' });
  const detail = el('div', { class: 'hist-detail' });
  body.append(list, detail);
  let selected: Version | null = select ? h.versions.find((v) => v.id === select) ?? null : null;
  selected ??= latest ?? null;
  let against: 'current' | 'previous' = initial;

  const addRow = (v: Version) => {
    const idx = h.versions.indexOf(v);
    const prev = h.versions[idx - 1];
    const delta = prev ? v.words - prev.words : 0;
    const row = el('button', { class: `hist-row${v === selected ? ' on' : ''}`, onclick: () => { selected = v; render(); } },
      el('div', { class: 'hist-kind', textContent: v.kind === 'named' && v.label ? `«${v.label}»` : KIND_LABEL[v.kind] }),
      el('div', { class: 'hist-meta', textContent: `${fmtWhen(v.t)} · ${v.u} · ${wordsLabel(v.words)}${prev && delta ? ` (${delta > 0 ? '+' : ''}${delta})` : ''}` }),
      v.kind !== 'named' && v.label ? el('div', { class: 'hist-note', textContent: v.label }) : '');
    list.append(row);
  };

  function render(): void {
    list.replaceChildren();
    for (const v of [...h.versions].reverse()) addRow(v);
    if (!h.versions.length) list.append(el('p', { class: 'hist-empty', textContent: 'Версий пока нет. Они появляются при каждом сохранении, раз в несколько минут во время работы и перед массовыми заменами. Именную версию можно записать кнопкой справа.' }));
    detail.replaceChildren();
    const bar = el('div', { class: 'hist-bar' });
    bar.append(el('button', { textContent: 'Записать именную версию…', onclick: () => nameDialog(host, () => { close2(); }) }));
    detail.append(bar);
    if (!selected) return;
    const v = selected;
    const idx = h.versions.indexOf(v);
    const prev = h.versions[idx - 1];
    const sel = el('select', {}, new Option('Сравнить с текущим документом', 'current'), ...(prev ? [new Option('Сравнить с предыдущей версией', 'previous')] : []));
    sel.value = against === 'previous' && prev ? 'previous' : 'current';
    sel.addEventListener('change', () => { against = sel.value as typeof against; render(); });
    bar.prepend(sel);
    if (against === 'current') bar.append(el('button', { class: 'danger', textContent: 'Восстановить документ целиком', onclick: () => { host.restore(v); close2(); } }));
    detail.append(el('div', { class: 'hist-title', textContent: `${v.kind === 'named' && v.label ? `«${v.label}»` : KIND_LABEL[v.kind]} · ${fmtWhen(v.t)} · ${v.u}` }));
    // Что показываем: старая сторона — выбранная версия; новая — текущий документ или предыдущая версия.
    const oldIds = against === 'previous' && prev ? prev.blocks : v.blocks;
    const newIds = against === 'previous' ? v.blocks : curIds;
    const oldStore = h.blocks;
    const newStore = against === 'previous' ? h.blocks : cur.blocks;
    detail.append(renderDiff(host, h, oldIds, newIds, oldStore, newStore, against === 'current', () => { close2(); }));
  }
  // После действий окно пересоздаётся целиком (состояние документа изменилось).
  const close2 = () => { document.getElementById('dialog')!.hidden = true; document.getElementById('dialog')!.replaceChildren(); };
  render();
}

function renderDiff(
  host: HistoryHost, h: History, oldIds: string[], newIds: string[],
  oldStore: Record<string, string>, newStore: Record<string, string>, canRevert: boolean, after: () => void,
): HTMLElement {
  const out = el('div', { class: 'diff' });
  out.append(el('p', { class: 'diff-legend' }, el('span', { class: 'd-ins', textContent: 'зелёное — добавлено' }), ' ', el('span', { class: 'd-del', textContent: 'красное — убрано' }), canRevert ? ' (относительно выбранной версии)' : ' (относительно предыдущей версии)'));
  const hunks = diffBlocks(oldIds, newIds);
  if (hunks.every((x) => x.type === 'same')) {
    out.append(el('p', { class: 'diff-same', textContent: 'Содержимое одинаково.' }));
    return out;
  }
  const text = (id: string, store: Record<string, string>) => blockText(store[id] ?? '');
  const label = canRevert ? 'выбранной версии' : 'старой версии';
  for (const x of hunks) {
    if (x.type === 'same') {
      const n = x.a[1] - x.a[0];
      out.append(el('div', { class: 'diff-skip', textContent: `… без изменений: ${n} ${n === 1 ? 'блок' : n < 5 ? 'блока' : 'блоков'}` }));
      continue;
    }
    const hunk = el('div', { class: 'diff-hunk' });
    const aN = x.a[1] - x.a[0];
    const bN = x.b[1] - x.b[0];
    if (aN === bN) {
      for (let k = 0; k < aN; k++) {
        const spans = wordDiff(text(oldIds[x.a[0] + k], oldStore), text(newIds[x.b[0] + k], newStore));
        hunk.append(el('p', { class: 'diff-p' }, ...spans.map((s) => el('span', { class: `d-${s.t}`, textContent: s.text }))));
      }
    } else {
      for (let k = x.a[0]; k < x.a[1]; k++) hunk.append(el('p', { class: 'diff-p d-del', textContent: text(oldIds[k], oldStore) || '(пустой блок)' }));
      for (let k = x.b[0]; k < x.b[1]; k++) hunk.append(el('p', { class: 'diff-p d-ins', textContent: text(newIds[k], newStore) || '(пустой блок)' }));
    }
    if (canRevert) {
      hunk.append(el('button', {
        class: 'diff-revert', textContent: `Вернуть как в ${label}`,
        onclick: () => {
          const nodes = oldIds.slice(x.a[0], x.a[1]).map((id) => nodeFromBlock(h, id));
          host.revert([x.b[0], x.b[1]], nodes);
          after();
          toast('Фрагмент возвращён. Отменить можно Ctrl+Z.');
        },
      }));
    }
    out.append(hunk);
  }
  return out;
}

/** Окно «Именная версия» (из ленты). */
export function askNamedVersion(host: HistoryHost): void {
  nameDialog(host, () => {});
}

function nameDialog(host: HistoryHost, done: () => void): void {
  dialog((box, close) => {
    const name = el('input', { type: 'text', placeholder: 'Например: «Отправлено на согласование»' });
    const go = () => { const v = name.value.trim(); close(); host.saveNamed(v || 'Версия'); done(); toast('Версия записана'); };
    box.addEventListener('keydown', (e) => { if (e.key === 'Enter') go(); });
    box.append(
      el('h2', { textContent: 'Именная версия' }),
      el('p', { textContent: 'Запишется точка, к которой можно вернуться. Именные версии хранятся дольше обычных.' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Название' }), name),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), el('button', { class: 'primary', textContent: 'Записать', onclick: go })),
    );
  });
}

const KIND_DISK: Record<HistoryEntry['kind'], string> = { backup: 'Перед перезаписью файла', discarded: 'Закрыт без сохранения' };
const fmtSize = (n: number) => (n > 1048576 ? `${(n / 1048576).toFixed(1)} МБ` : `${Math.max(1, Math.round(n / 1024))} КБ`);

async function renderDisk(host: HistoryHost, body: HTMLElement): Promise<void> {
  body.append(el('p', { class: 'hist-empty', textContent: 'Копии на диске лежат в профиле пользователя: предыдущее содержимое файла перед каждой перезаписью (10 последних на файл, 60 дней) и документы, закрытые кнопкой «Не сохранять» (20 последних, 14 дней).' }));
  let list: HistoryEntry[] = [];
  try { list = await historyList(); } catch (e) { body.append(el('p', { textContent: `Не удалось прочитать список: ${(e as Error).message ?? e}` })); return; }
  if (!list.length) { body.append(el('p', { class: 'hist-empty', textContent: 'Копий пока нет.' })); return; }
  const box = el('div', { class: 'disk-list' });
  for (const e of list) {
    box.append(el('div', { class: 'rec-item' },
      el('div', { class: 'rec-title', textContent: e.name.replace(/\.docx$/i, '') }),
      el('div', { class: 'rec-meta', textContent: `${KIND_DISK[e.kind]} · ${fmtWhen(e.ms)} · ${fmtSize(e.size)}` }),
      el('div', { class: 'rec-buttons' }, el('button', { class: 'primary', textContent: 'Открыть в новом окне', onclick: () => host.openCopy(e.path) }))));
  }
  body.append(box);
}
