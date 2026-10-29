// «Диспетчер имён», как в Excel: список имён книги и листов, создание, изменение, удаление.
// Имя вводят и в поле имени (слева от строки формул): выделенный диапазон получает это имя.

import type { Model } from './engine';

export interface NamesHost {
  model(): Model;
  sheet(): number;
  /** Выделение как ссылка с именем листа: «Лист1!$A$1:$B$2». */
  selectionRef(): string;
  /** Перейти к ссылке из определения имени. */
  goTo(formula: string): void;
  changed(): void;
  toast(text: string): void;
}

import { addRangePicker } from './rangepick';

let host: NamesHost;
export function initNames(h: NamesHost): void { host = h; }

/** Имена самой программы и Excel (фильтры, закрепление, область печати) в списке не нужны. */
const hidden = (name: string) => name.startsWith('_ENS_') || name.startsWith('_xlnm.');

const NAME_RE = /^[\p{L}_\\][\p{L}\p{N}_.\\]{0,254}$/u;

/** Похоже на адрес ячейки (A1, XFD1048576, R1C1): так назвать нельзя. */
const looksLikeCell = (name: string) => /^[A-Za-z]{1,3}\d+$/.test(name) || /^[Rr]\d*[Cc]\d*$/.test(name) || /^[RrCc]$/.test(name);

export function checkName(name: string): string | null {
  if (!NAME_RE.test(name)) return 'Имя начинается с буквы или «_» и состоит из букв, цифр, точек и «_», без пробелов.';
  if (looksLikeCell(name)) return 'Имя не должно быть похоже на адрес ячейки (A1, R1C1).';
  return null;
}

const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);

/** Определение имени по тексту (без «=»); null — такого имени нет. Для перехода из поля имени. */
export function nameFormula(name: string): string | null {
  const found = host.model().getDefinedNameList().find((d) =>
    d.name.toLowerCase() === name.toLowerCase() && (d.scope == null || d.scope === host.sheet()));
  return found ? found.formula.replace(/^=/, '') : null;
}

/** Дать имя текущему выделению (из поля имени); true — имя создано. */
export function defineFromSelection(name: string): boolean {
  const problem = checkName(name);
  if (problem) { host.toast(problem); return false; }
  const m = host.model();
  if (m.getDefinedNameList().some((d) => d.scope == null && d.name.toLowerCase() === name.toLowerCase())) {
    host.toast(`Имя «${name}» уже есть. Измените его в диспетчере имён (Формулы → Диспетчер имён).`);
    return false;
  }
  try {
    m.newDefinedName(name, null, host.selectionRef());
    host.changed();
    host.toast(`Имя «${name}» присвоено выделенному диапазону.`);
    return true;
  } catch (e) {
    host.toast(`Не удалось создать имя: ${e}`);
    return false;
  }
}

export function openNameManager(): void {
  const overlay = document.getElementById('recovery')!;
  const m = host.model();
  const sheets = m.getWorksheetsProperties().map((p) => p.name);
  let selected = -1;
  const list = () => m.getDefinedNameList().filter((d) => !hidden(d.name));

  const draw = () => {
    const rows = list();
    overlay.innerHTML = `<div class="modal names-modal">
      <h2>Диспетчер имён</h2>
      <div class="names-list">${rows.length ? `<table>
        <thead><tr><th>Имя</th><th>Диапазон</th><th>Область</th></tr></thead>
        <tbody>${rows.map((d, i) => `<tr data-i="${i}" class="${i === selected ? 'on' : ''}"><td>${esc(d.name)}</td><td>${esc(d.formula.replace(/^=/, ''))}</td>
          <td>${d.scope == null ? 'Книга' : esc(sheets[d.scope] ?? '')}</td></tr>`).join('')}</tbody></table>` : '<p>Имён пока нет. Выделите диапазон, введите имя в поле слева от строки формул и нажмите Enter — или нажмите «Создать».</p>'}</div>
      <div class="names-form" hidden>
        <label>Имя <input data-f="name" autocomplete="off" spellcheck="false"></label>
        <label>Область <select data-f="scope"><option value="">Книга</option>${sheets.map((s, i) => `<option value="${i}">${esc(s)}</option>`).join('')}</select></label>
        <label>Диапазон <input data-f="formula" autocomplete="off" spellcheck="false"></label>
        <div class="names-note"></div>
        <div class="rec-actions"><button data-act="useSel">Взять из выделения</button><button data-act="save" class="primary">Сохранить</button><button data-act="back">Отмена</button></div>
      </div>
      <div class="rec-actions names-actions">
        <button data-act="new">Создать…</button><button data-act="edit" ${selected < 0 ? 'disabled' : ''}>Изменить…</button>
        <button data-act="goto" ${selected < 0 ? 'disabled' : ''}>Перейти</button>
        <button data-act="delete" class="danger" ${selected < 0 ? 'disabled' : ''}>Удалить</button>
        <button data-act="close" class="primary">Закрыть</button>
      </div></div>`;
    overlay.hidden = false;
    const box = overlay.querySelector<HTMLElement>('.modal')!;
    const form = box.querySelector<HTMLElement>('.names-form')!;
    const f = (k: string) => box.querySelector<HTMLInputElement & HTMLSelectElement>(`[data-f="${k}"]`)!;
    let editing: { name: string; scope: number | null } | null = null;
    addRangePicker(f('formula'), 'sheet');
    const openForm = (init?: { name: string; scope: number | null; formula: string }) => {
      editing = init ? { name: init.name, scope: init.scope } : null;
      f('name').value = init?.name ?? '';
      f('scope').value = init?.scope == null ? '' : String(init.scope);
      f('formula').value = init?.formula.replace(/^=/, '') ?? host.selectionRef();
      form.hidden = false;
      f('name').focus();
    };
    box.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Escape') { if (!form.hidden) form.hidden = true; else overlay.hidden = true; }
    });
    box.addEventListener('click', (e) => {
      const t = e.target as HTMLElement;
      const row = t.closest<HTMLElement>('tr[data-i]');
      if (row) { selected = Number(row.dataset.i); draw(); return; }
      const act = t.closest<HTMLElement>('[data-act]')?.dataset.act;
      const cur = selected >= 0 ? list()[selected] : undefined;
      if (act === 'close') { overlay.hidden = true; return; }
      if (act === 'back') { form.hidden = true; return; }
      if (act === 'new') openForm();
      else if (act === 'edit' && cur) openForm({ name: cur.name, scope: cur.scope ?? null, formula: cur.formula });
      else if (act === 'goto' && cur) { overlay.hidden = true; host.goTo(cur.formula.replace(/^=/, '')); }
      else if (act === 'delete' && cur) {
        if (!confirm(`Удалить имя «${cur.name}»? Формулы, которые его используют, перестанут считаться.`)) return;
        try { m.deleteDefinedName(cur.name, cur.scope ?? null); selected = -1; host.changed(); } catch (err) { host.toast(String(err)); }
        draw();
      } else if (act === 'useSel') f('formula').value = host.selectionRef();
      else if (act === 'save') {
        const name = f('name').value.trim(), formula = f('formula').value.trim().replace(/^=/, '');
        const scope = f('scope').value === '' ? null : Number(f('scope').value);
        const note = box.querySelector<HTMLElement>('.names-note')!;
        const problem = checkName(name) ?? (formula ? null : 'Укажите диапазон или формулу.');
        if (problem) { note.textContent = problem; return; }
        const clash = list().some((d) => d.name.toLowerCase() === name.toLowerCase() && (d.scope ?? null) === scope
          && !(editing && editing.name === d.name && editing.scope === scope));
        if (clash) { note.textContent = `Имя «${name}» в этой области уже есть.`; return; }
        try {
          if (editing) m.updateDefinedName(editing.name, editing.scope, name, scope, formula);
          else m.newDefinedName(name, scope, formula);
          host.changed();
        } catch (err) { note.textContent = `Не получилось: ${err}`; return; }
        selected = -1;
        draw();
      }
    });
  };
  draw();
}
