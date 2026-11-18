// SPDX-License-Identifier: AGPL-3.0-or-later
// «Рецензирование → Сравнить»: два файла DOCX, текущий документ или версия из истории.
// Результат — новый документ с исправлениями; исходные файлы и текущий документ не меняются.

import { closeHistory } from 'prosemirror-history';
import type { Node as PMNode } from 'prosemirror-model';
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { compareDocs, listRevisions, resolveRevisions, type RevRange } from './compare';
import { importDocx, type DocMeta } from './docx';
import { dialog, el, toast } from './ui';
import { versionDoc, type History } from './versions';
import { fmtWhen, KIND_LABEL } from './history-ui';

export interface CompareHost {
  doc(): PMNode;
  meta(): DocMeta;
  docName(): string;
  history(): History;
  user(): string;
  /** Открыть результат новым документом (с вопросом о несохранённом). */
  create(doc: PMNode, meta: DocMeta, name: string): void;
}

interface Side { doc: PMNode; meta: DocMeta; name: string }

/** Сравнить две редакции: старая — «было», новая — «стало». Автор правок — автор новой редакции. */
export function compareSides(older: Side, newer: Side, author: string, date = new Date()): { doc: PMNode; meta: DocMeta; name: string; inserted: number; deleted: number } {
  const res = compareDocs(older.doc, newer.doc, { author, date: date.toISOString().replace(/\.\d+Z$/, 'Z') });
  const strip = (n: string) => n.replace(/\.docx$/i, '');
  return { doc: res.doc, meta: structuredClone(newer.meta), name: `Сравнение — ${strip(older.name)} и ${strip(newer.name)}`, inserted: res.inserted, deleted: res.deleted };
}

export function compareDialog(host: CompareHost, preset?: { olderVersion?: string }): void {
  dialog((box, close) => {
    box.classList.add('compare-dialog');
    const status = el('p', { role: 'status' });
    const versions = [...host.history().versions].reverse();
    const pick = (label: string, initial: string) => {
      const select = el('select', { ariaLabel: label },
        new Option('Файл DOCX…', 'file'), new Option(`Текущий документ («${host.docName()}»)`, 'current'),
        ...versions.map((v) => new Option(`Версия: ${v.kind === 'named' && v.label ? `«${v.label}»` : KIND_LABEL[v.kind]} · ${fmtWhen(v.t)} · ${v.u}`, 'v:' + v.id)));
      select.value = initial;
      const file = el('input', { type: 'file', accept: '.docx', ariaLabel: label + ': файл' });
      const sync = () => { file.hidden = select.value !== 'file'; };
      select.addEventListener('change', () => { sync(); status.textContent = ''; });
      sync();
      const row = el('div', { class: 'dlg-row' }, el('label', { textContent: label }), select, file);
      const read = async (): Promise<Side> => {
        if (select.value === 'current') return { doc: host.doc(), meta: host.meta(), name: host.docName() };
        if (select.value.startsWith('v:')) {
          const v = host.history().versions.find((x) => x.id === select.value.slice(2));
          if (!v) throw new Error('Версия не найдена.');
          return { doc: versionDoc(host.history(), v), meta: { ...host.meta(), ...structuredClone(v.meta) }, name: `${host.docName()} (${fmtWhen(v.t)})` };
        }
        const f = file.files?.[0];
        if (!f) throw new Error(`${label}: выберите файл DOCX.`);
        if (!/\.docx$/i.test(f.name)) throw new Error('Нужен файл DOCX.');
        if (f.size > 32 * 1024 * 1024) throw new Error('Файл превышает 32 МБ.');
        const imported = importDocx(new Uint8Array(await f.arrayBuffer()));
        return { doc: imported.doc, meta: imported.meta, name: f.name };
      };
      return { row, read, select, file };
    };
    const older = pick('Исходный документ (было)', preset?.olderVersion ? 'v:' + preset.olderVersion : versions.length ? 'v:' + versions[0].id : 'file');
    const newer = pick('Изменённый документ (стало)', 'current');
    older.select.dataset.compare = 'older';
    newer.select.dataset.compare = 'newer';
    older.file.dataset.compare = 'older-file';
    newer.file.dataset.compare = 'newer-file';
    const authorIn = el('input', { type: 'text', value: host.user(), ariaLabel: 'Автор исправлений' });
    const go = el('button', { class: 'primary', textContent: 'Сравнить', onclick: async () => {
      go.disabled = true;
      status.textContent = 'Сравнение…';
      try {
        const [a, b] = await Promise.all([older.read(), newer.read()]);
        if (!box.isConnected) return;
        const res = compareSides(a, b, authorIn.value.trim() || host.user());
        const lossy = [...new Set([...a.meta.lossy, ...b.meta.lossy])];
        if (!res.inserted && !res.deleted) { status.textContent = 'Различий в тексте нет.'; go.disabled = false; return; }
        close();
        host.create(res.doc, { ...res.meta, lossy }, res.name);
        toast(`Исправлений: ${res.inserted + res.deleted} (вставок ${res.inserted}, удалений ${res.deleted}). Принимайте их по одной на вкладке «Проверка».`);
      } catch (e) {
        status.textContent = 'Не удалось сравнить: ' + (e as Error).message;
        go.disabled = false;
      }
    } });
    box.append(el('h2', { textContent: 'Сравнить документы' }),
      el('p', { textContent: 'Выберите две редакции: файлы DOCX, текущий документ или версию из истории. Различия покажутся по словам как исправления: вставленное подчёркнуто, удалённое зачёркнуто. Результат откроется новым документом; исходные файлы не меняются.' }),
      older.row, newer.row, el('div', { class: 'dlg-row' }, el('label', { textContent: 'Автор исправлений' }), authorIn), status,
      el('div', { class: 'dlg-actions' }, go, el('button', { textContent: 'Отмена', onclick: close })));
  }, false);
}

// --- Работа с исправлениями в редакторе

/** Исправление под курсором (или пересекающее выделение), иначе ближайшее следующее. */
export function currentRevision(state: EditorState, dir: 1 | -1 | 0 = 0): RevRange | null {
  const list = listRevisions(state.doc);
  if (!list.length) return null;
  const { from, to } = state.selection;
  if (dir === 0) {
    const here = list.find((r) => r.from <= to && r.to >= from && !(r.to === from && from !== to));
    if (here) return here;
    return list.find((r) => r.from >= from) ?? list[0];
  }
  if (dir === 1) return list.find((r) => r.from > from || (r.from === from && r.to > to)) ?? list[0];
  return [...list].reverse().find((r) => r.from < from) ?? list[list.length - 1];
}

const select = (tr: Transaction, r: RevRange | null): Transaction => {
  if (!r) return tr;
  const max = tr.doc.content.size;
  const $from = tr.doc.resolve(Math.min(r.from, max)), $to = tr.doc.resolve(Math.min(r.to, max));
  return tr.setSelection(TextSelection.between($from, $to)).scrollIntoView();
};

export function goRevision(view: EditorView, dir: 1 | -1): boolean {
  const r = currentRevision(view.state, dir);
  if (!r) { toast('Исправлений нет.'); return false; }
  view.dispatch(select(view.state.tr, r));
  return true;
}

/** Принять или отклонить исправление под курсором и перейти к следующему. */
export function resolveOne(view: EditorView, accept: boolean): boolean {
  const r = currentRevision(view.state);
  if (!r) { toast('Исправлений нет.'); return false; }
  const tr = resolveRevisions(view.state, new Set([r.id]), accept);
  if (!tr) return false;
  const after = tr.mapping.map(r.from, -1);
  tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(after, tr.doc.content.size))));
  const next = listRevisions(tr.doc).find((x) => x.from >= after) ?? listRevisions(tr.doc)[0] ?? null;
  // Каждое решение — отдельный шаг отмены, даже если кнопку нажимают подряд.
  view.dispatch(closeHistory(select(tr, next)));
  if (!next) toast('Все исправления обработаны.');
  return true;
}

export function resolveAll(view: EditorView, accept: boolean): boolean {
  const n = listRevisions(view.state.doc).length;
  const tr = resolveRevisions(view.state, 'all', accept);
  if (!tr) { toast('Исправлений нет.'); return false; }
  view.dispatch(closeHistory(tr.scrollIntoView()));
  toast(`${accept ? 'Принято' : 'Отклонено'} исправлений: ${n}.`);
  return true;
}

export type RevView = 'markup' | 'original' | 'final';
export const REV_VIEW_LABEL: Record<RevView, string> = { markup: 'Исправления', original: 'Было (исходный)', final: 'Стало (без пометок)' };
