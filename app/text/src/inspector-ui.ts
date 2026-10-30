// «Рентген» форматирования (панель справа: что и откуда задано в этом месте) и окно «Проверка оформления».

import type { EditorState } from 'prosemirror-state';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { DocDefaults } from './docx';
import { allProfiles, CUSTOM_ID, defaultCustomProfile, loadCustomProfile, runChecks, saveCustomProfile, type CheckHost, type Issue, type Profile } from './checks';
import { dialog, el, toast } from './ui';

// --- Панель «Рентген»

let panel: HTMLElement | null = null;

export const inspectorOpen = () => !!panel;

interface Row { k: string; v: string; manual: boolean }

export function describeHere(state: EditorState, d: DocDefaults): Row[] {
  const sel = state.selection;
  // Выделен весь документ или несколько ячеек: показываем форматирование у начала выделения.
  const $from = sel.$from.parent.isTextblock ? sel.$from : TextSelection.near(sel.$from, 1).$from;
  const { from, to, empty } = sel;
  const rows: Row[] = [];
  const block = $from.parent;
  if (!block.isTextblock) return rows;
  const isHead = block.type.name === 'heading';
  rows.push({ k: 'Стиль', v: isHead ? `Заголовок ${block.attrs.level}` : 'Обычный', manual: isHead });
  for (let depth = $from.depth; depth > 0; depth--) {
    const t = $from.node(depth).type.name;
    if (t === 'list_item') { rows.push({ k: 'Список', v: $from.node(depth - 1).type.name === 'ordered_list' ? 'нумерованный' : 'маркированный', manual: true }); break; }
  }
  const a = block.attrs;
  const pt = (n: number) => `${Math.round(n * 10) / 10} пт`;
  rows.push({ k: 'Выравнивание', v: ({ left: 'по левому краю', center: 'по центру', right: 'по правому краю', justify: 'по ширине' } as Record<string, string>)[a.align ?? 'left'], manual: !!a.align });
  rows.push({ k: 'Отступ слева', v: a.indent ? pt(a.indent) : '0', manual: !!a.indent });
  rows.push({ k: 'Первая строка', v: a.first ? pt(a.first) : 'нет', manual: !!a.first });
  rows.push({ k: 'Интервал перед / после', v: `${a.before != null ? pt(a.before) : 'по умолчанию'} / ${a.after != null ? pt(a.after) : 'по умолчанию'}`, manual: a.before != null || a.after != null });
  rows.push({ k: 'Межстрочный', v: a.line ? String(a.line) : `${d.line} (по умолчанию)`, manual: !!a.line });
  if (a.pageBreakBefore) rows.push({ k: 'Разрыв страницы', v: 'перед абзацем', manual: true });
  const marks = empty ? state.storedMarks || $from.marks() : $from.marksAcross(state.doc.resolve(to)) ?? [];
  const ts = marks.find((m) => m.type.name === 'textStyle');
  rows.push({ k: 'Шрифт', v: (ts?.attrs.font as string) || `${d.font} (по умолчанию)`, manual: !!ts?.attrs.font });
  rows.push({ k: 'Размер', v: ts?.attrs.size ? pt(ts.attrs.size as number) : `${d.size} пт (по умолчанию)`, manual: !!ts?.attrs.size });
  const has = (n: string) => marks.some((m) => m.type.name === n);
  const style = [has('strong') && 'жирный', has('em') && 'курсив', has('underline') && 'подчёркнутый', has('strike') && 'зачёркнутый', has('sup') && 'надстрочный', has('sub') && 'подстрочный'].filter(Boolean);
  rows.push({ k: 'Начертание', v: style.join(', ') || 'обычное', manual: style.length > 0 });
  if (ts?.attrs.color) rows.push({ k: 'Цвет текста', v: String(ts.attrs.color), manual: true });
  if (ts?.attrs.bg) rows.push({ k: 'Заливка', v: String(ts.attrs.bg), manual: true });
  const link = marks.find((m) => m.type.name === 'link');
  if (link) rows.push({ k: 'Ссылка', v: String(link.attrs.href), manual: true });
  void from;
  return rows;
}

export function toggleInspector(get: () => { view: EditorView; defaults: DocDefaults }, actions: { resetPara(): void; resetChars(): void }): void {
  if (panel) { panel.remove(); panel = null; return; }
  panel = el('aside', { class: 'inspector' });
  document.body.append(panel);
  renderInspector(get, actions);
}

export function renderInspector(get: () => { view: EditorView; defaults: DocDefaults }, actions: { resetPara(): void; resetChars(): void }): void {
  if (!panel) return;
  const { view, defaults } = get();
  const rows = describeHere(view.state, defaults);
  panel.replaceChildren(
    el('div', { class: 'insp-head' }, el('b', { textContent: 'Форматирование здесь' }), el('button', { textContent: '✕', title: 'Закрыть', onclick: () => { panel?.remove(); panel = null; } })),
    el('p', { class: 'insp-note', textContent: 'Жирным — то, что задано вручную; обычным — из настроек документа.' }),
    ...rows.map((r) => el('div', { class: 'insp-row' }, el('span', { textContent: r.k }), el('span', { class: r.manual ? 'man' : '', textContent: r.v }))),
    el('div', { class: 'insp-actions' },
      el('button', { textContent: 'Сбросить формат абзаца', onclick: actions.resetPara }),
      el('button', { textContent: 'Сбросить формат текста', onclick: actions.resetChars })),
  );
}

// --- Окно «Проверка оформления»

export function checkDialog(host: CheckHost, safety: (label: string) => void, view: () => EditorView): void {
  let profileId = (() => { try { return localStorage.getItem('ens-text.profile') ?? 'gost7097'; } catch { return 'gost7097'; } })();
  dialog((box, close) => {
    box.classList.add('check-modal');
    const profile = (): Profile | null => allProfiles().find((p) => p.id === profileId) ?? null;
    const sel = el('select', {});
    const fillSel = () => {
      sel.replaceChildren(new Option('Только замечания по документу', ''), ...allProfiles().map((p) => new Option(p.name, p.id)));
      sel.value = allProfiles().some((p) => p.id === profileId) ? profileId : '';
    };
    fillSel();
    const editBtn = el('button', { textContent: 'Изменить свой профиль…', title: 'Поля, шрифты, интервал, отступ и номера страниц по регламенту вашей организации', onclick: () => profileEditor(() => { profileId = CUSTOM_ID; try { localStorage.setItem('ens-text.profile', profileId); } catch { /* без памяти */ } fillSel(); render(); }) });
    const list = el('div', { class: 'check-list' });
    const note = el('p', { class: 'check-note' });
    const render = () => {
      const issues = runChecks(host, profile());
      list.replaceChildren(...issues.map((i) => issueRow(i)));
      const bad = issues.filter((i) => !i.ok);
      note.textContent = profile() ? profile()!.note : 'Выберите профиль оформления (ГОСТ), чтобы проверить поля, шрифт, интервал и номера страниц. Значения — рабочие: сверьте с регламентом вашей организации.';
      fixAll.disabled = !bad.some((i) => i.fix);
      fixAll.textContent = bad.length ? `Исправить всё (${bad.length})` : 'Всё в порядке';
    };
    const issueRow = (i: Issue) => el('div', { class: `check-row ${i.ok ? 'ok' : 'bad'}` },
      el('span', { class: 'check-mark', textContent: i.ok ? '✓' : '✗' }),
      el('div', { class: 'check-text' }, el('div', { textContent: i.title }), i.detail && !i.ok ? el('div', { class: 'check-detail', textContent: i.detail }) : ''),
      i.at?.length && !i.ok ? el('button', { textContent: 'Показать', onclick: () => { const v = view(); v.dispatch(v.state.tr.setSelection(TextSelection.near(v.state.doc.resolve(i.at![0] + 1))).scrollIntoView()); } }) : '',
      i.fix && !i.ok ? el('button', { class: 'primary', textContent: 'Исправить', onclick: () => { safety(`Перед исправлением: ${i.title}`); i.fix!(); render(); } }) : '');
    const fixAll = el('button', {
      class: 'primary', onclick: () => {
        safety('Перед исправлением оформления');
        let guard = 0;
        // Исправление одного замечания может открыть другое (например, пробелы → типографика): повторяем, но не бесконечно.
        while (guard++ < 3) {
          const bad = runChecks(host, profile()).filter((i) => !i.ok && i.fix);
          if (!bad.length) break;
          for (const i of bad) i.fix!();
        }
        render();
        toast('Оформление исправлено. Отменить каждый шаг можно Ctrl+Z, всё сразу — вернуть версию «Перед исправлением оформления».');
      },
    });
    sel.addEventListener('change', () => { profileId = sel.value; try { localStorage.setItem('ens-text.profile', profileId); } catch { /* без памяти */ } render(); });
    box.append(el('h2', { textContent: 'Проверка оформления' }), el('div', { class: 'dlg-row' }, el('label', { textContent: 'Профиль' }), sel, editBtn), note, list,
      el('div', { class: 'dlg-actions' }, fixAll, el('button', { textContent: 'Закрыть', onclick: close })));
    render();
  }, false);
}

// --- Редактор своего профиля оформления

const PN: [string, string][] = [['', 'Номера страниц не проверять'], ['header-center', 'Вверху по центру'], ['header-right', 'Вверху справа'], ['footer-center', 'Внизу по центру'], ['footer-right', 'Внизу справа']];

function profileEditor(done: () => void): void {
  dialog((box, close) => {
    const p = loadCustomProfile();
    const num = (v: number) => el('input', { type: 'number', value: String(v), min: '0', step: '0.5' });
    const name = el('input', { type: 'text', value: p.name });
    const m = { left: num(p.margins.left), right: num(p.margins.right), top: num(p.margins.top), bottom: num(p.margins.bottom) };
    const fonts = el('textarea', { rows: 3, value: Object.entries(p.fonts).map(([f, [a, b]]) => `${f}: ${a}${a === b ? '' : `-${b}`}`).join('\n'), placeholder: 'Times New Roman: 13-14' });
    const lines = el('input', { type: 'text', value: p.lines.join('; '), placeholder: '1; 1.5' });
    const indent = num(p.indentCm);
    const just = el('input', { type: 'checkbox', checked: p.justify });
    const pn = el('select', {}, ...PN.map(([v, t]) => new Option(t, v)));
    pn.value = p.pageNumbers ? `${p.pageNumbers.where}-${p.pageNumbers.align}` : '';
    const skip = el('input', { type: 'checkbox', checked: p.pageNumbers?.skipFirst ?? true });
    const note = el('textarea', { rows: 2, value: p.note === defaultCustomProfile().note ? '' : p.note, placeholder: 'Пометка: на что опирается профиль, например «Положение об документообороте, п. 4.2»' });
    const err = el('div', { class: 'ru-out bad' });
    err.hidden = true;
    const row = (label: string, ...k: Node[]) => el('div', { class: 'dlg-row' }, el('label', { textContent: label }), ...k);
    const save = () => {
      const f: Record<string, [number, number]> = {};
      for (const l of fonts.value.split('\n').map((x) => x.trim()).filter(Boolean)) {
        const mm = /^(.+?)\s*:\s*(\d+(?:[.,]\d+)?)(?:\s*[-–]\s*(\d+(?:[.,]\d+)?))?\s*(?:пт)?$/.exec(l);
        if (!mm) { err.textContent = `Не понял строку шрифтов: «${l}». Нужно так: Times New Roman: 13-14`; err.hidden = false; return; }
        const a = parseFloat(mm[2].replace(',', '.')), b = mm[3] ? parseFloat(mm[3].replace(',', '.')) : a;
        f[mm[1].trim()] = [Math.min(a, b), Math.max(a, b)];
      }
      const ls = lines.value.split(/[;,\s]+/).map((x) => parseFloat(x.replace(',', '.'))).filter((x) => x > 0);
      if (!Object.keys(f).length || !ls.length) { err.textContent = 'Нужен хотя бы один шрифт и один межстрочный интервал.'; err.hidden = false; return; }
      const [where, align] = pn.value.split('-');
      saveCustomProfile({
        ...p, name: name.value.trim() || defaultCustomProfile().name, note: note.value.trim() || defaultCustomProfile().note,
        margins: { left: +m.left.value, right: +m.right.value, top: +m.top.value, bottom: +m.bottom.value },
        fonts: f, lines: ls, indentCm: +indent.value || 0, justify: just.checked,
        pageNumbers: pn.value ? { where: where as 'header' | 'footer', align: align as 'center' | 'right', skipFirst: skip.checked } : null,
      });
      close();
      done();
      toast('Свой профиль сохранён');
    };
    box.append(
      el('h2', { textContent: 'Свой профиль оформления' }),
      row('Название', name),
      row('Поля, мм (не меньше)', el('span', { class: 'prof-margins' }, 'лев. ', m.left, ' прав. ', m.right, ' верх. ', m.top, ' ниж. ', m.bottom)),
      row('Шрифты и кегль, пт', fonts),
      row('Межстрочный интервал', lines),
      row('Абзацный отступ, см', indent),
      row('Выравнивание', el('label', {}, just, ' по ширине')),
      row('Номера страниц', pn, el('label', {}, skip, ' кроме первой')),
      row('Пометка', note),
      err,
      el('p', { textContent: 'Профиль хранится на этом компьютере и появляется в списке проверки. Шрифты — по одному на строке; первый считается предпочтительным.' }),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), el('button', { class: 'primary', textContent: 'Сохранить', onclick: save })),
    );
  }, false);
}
