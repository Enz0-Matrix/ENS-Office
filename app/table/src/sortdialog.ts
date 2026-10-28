// Окна сортировки, как в Excel, с пожеланиями пользователя:
//  * «Обнаружены данные вне выделения»: расширить (строки целиком) или только выделенное;
//  * «Настраиваемая сортировка»: несколько уровней, заголовки, что делать с пустыми.

import { columnName } from './engine';

export type BlankMode = 'last' | 'first' | 'zero';

export interface SortLevel { column: number; descending: boolean }

export interface SortRequest {
  levels: SortLevel[];
  header: boolean;
  blanks: BlankMode;
  /** true — строки целиком (вся таблица), false — только выделенные столбцы. */
  expand: boolean;
}

const BLANKS: [BlankMode, string][] = [
  ['last', 'в конце (как в Excel)'],
  ['first', 'в начале'],
  ['zero', 'как ноль — строки не разъезжаются'],
];

const overlay = () => document.getElementById('recovery')!;
const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);

function modal(html: string): HTMLElement {
  const o = overlay();
  o.innerHTML = `<div class="modal sort-modal">${html}</div>`;
  o.hidden = false;
  const m = o.querySelector<HTMLElement>('.modal')!;
  m.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Escape') (m.querySelector('[data-act="cancel"]') as HTMLButtonElement)?.click();
  });
  return m;
}

/** Excel: «Обнаружены данные вне указанного диапазона». */
export function askExpand(): Promise<'expand' | 'selection' | 'custom' | null> {
  const m = modal(`
    <h2>Сортировка</h2>
    <p>Рядом с выделенным есть другие данные. Если отсортировать только выделенное,
       значения в соседних столбцах останутся на своих местах и строки «разъедутся».</p>
    <label class="sort-choice"><input type="radio" name="exp" value="expand" checked> Автоматически расширить выделенный диапазон — строки целиком</label>
    <label class="sort-choice"><input type="radio" name="exp" value="selection"> Сортировать в пределах указанного выделения</label>
    <div class="rec-actions">
      <button data-act="custom">Настраиваемая сортировка…</button>
      <button data-act="ok" class="primary">Сортировка</button>
      <button data-act="cancel">Отмена</button>
    </div>`);
  return new Promise((resolve) => {
    m.addEventListener('click', (e) => {
      const act = (e.target as HTMLElement).closest<HTMLElement>('[data-act]')?.dataset.act;
      if (!act) return;
      overlay().hidden = true;
      if (act === 'cancel') resolve(null);
      else if (act === 'custom') resolve('custom');
      else resolve(m.querySelector<HTMLInputElement>('input[name="exp"]:checked')!.value as 'expand' | 'selection');
    });
    m.querySelector<HTMLButtonElement>('[data-act="ok"]')!.focus();
  });
}

/** «Настраиваемая сортировка»: уровни, заголовки, пустые ячейки, область. */
export function customSort(columns: { column: number; title: string }[], initial: SortRequest, canExpand: boolean): Promise<SortRequest | null> {
  const levels = initial.levels.length ? [...initial.levels] : [{ column: columns[0]?.column ?? 1, descending: false }];
  const m = modal(`
    <h2>Настраиваемая сортировка</h2>
    <div class="sort-levels"></div>
    <button data-act="add" class="sort-add">+ Добавить уровень</button>
    <label class="sort-choice"><input type="checkbox" data-f="header" ${initial.header ? 'checked' : ''}> Мои данные содержат заголовки</label>
    <label class="sort-row">Пустые ячейки
      <select data-f="blanks">${BLANKS.map(([v, t]) => `<option value="${v}" ${v === initial.blanks ? 'selected' : ''}>${t}</option>`).join('')}</select></label>
    ${canExpand ? `<label class="sort-row">Что переставлять
      <select data-f="expand"><option value="1" ${initial.expand ? 'selected' : ''}>строки целиком (вся таблица)</option>
      <option value="0" ${initial.expand ? '' : 'selected'}>только выделенные столбцы</option></select></label>` : ''}
    <div class="rec-actions"><button data-act="ok" class="primary">ОК</button><button data-act="cancel">Отмена</button></div>`);
  const box = m.querySelector<HTMLElement>('.sort-levels')!;
  const render = () => {
    box.innerHTML = levels.map((l, i) => `
      <div class="sort-level" data-i="${i}">
        <span>${i === 0 ? 'Сортировать по' : 'Затем по'}</span>
        <select data-k="column">${columns.map((c) => `<option value="${c.column}" ${c.column === l.column ? 'selected' : ''}>${esc(c.title)}</option>`).join('')}</select>
        <select data-k="order"><option value="0" ${l.descending ? '' : 'selected'}>От А до Я / по возрастанию</option>
          <option value="1" ${l.descending ? 'selected' : ''}>От Я до А / по убыванию</option></select>
        ${levels.length > 1 ? `<button data-act="del" title="Удалить уровень">✕</button>` : ''}
      </div>`).join('');
  };
  render();
  box.addEventListener('change', (e) => {
    const el = e.target as HTMLSelectElement;
    const i = Number(el.closest<HTMLElement>('.sort-level')!.dataset.i);
    if (el.dataset.k === 'column') levels[i].column = Number(el.value);
    else levels[i].descending = el.value === '1';
  });
  return new Promise((resolve) => {
    m.addEventListener('click', (e) => {
      const t = e.target as HTMLElement;
      const act = t.closest<HTMLElement>('[data-act]')?.dataset.act;
      if (act === 'add') {
        const used = new Set(levels.map((l) => l.column));
        levels.push({ column: columns.find((c) => !used.has(c.column))?.column ?? columns[0].column, descending: false });
        render();
      } else if (act === 'del') {
        levels.splice(Number(t.closest<HTMLElement>('.sort-level')!.dataset.i), 1);
        render();
      } else if (act === 'cancel') {
        overlay().hidden = true;
        resolve(null);
      } else if (act === 'ok') {
        overlay().hidden = true;
        resolve({
          levels,
          header: m.querySelector<HTMLInputElement>('[data-f="header"]')!.checked,
          blanks: m.querySelector<HTMLSelectElement>('[data-f="blanks"]')!.value as BlankMode,
          expand: canExpand ? m.querySelector<HTMLSelectElement>('[data-f="expand"]')!.value === '1' : initial.expand,
        });
      }
    });
  });
}

/** Подпись столбца в окне: заголовок таблицы или «Столбец B». */
export function columnTitle(column: number, header: string): string {
  return header ? `${header} (${columnName(column)})` : `Столбец ${columnName(column)}`;
}
