// SPDX-License-Identifier: AGPL-3.0-or-later
// Окно «Закладка» (Вставка → Ссылки → Закладка, Ctrl+Shift+F5): добавить, перейти, переименовать, удалить.
// Внизу — ссылки без цели (на закладки, которых нет), с переходом к каждой.

import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { addBookmark, brokenLinks, gotoBookmark, isHidden, linksTo, listBookmarks, removeBookmark, renameBookmark, validateName } from './bookmarks';
import { choose, dialog, el, toast } from './ui';

/** Имя по умолчанию: первые слова выделенного текста, как в Word. */
function suggest(view: EditorView): string {
  const { from, to } = view.state.selection;
  const words = view.state.doc.textBetween(from, to, ' ', ' ').match(/[\p{L}\p{N}]+/gu) ?? [];
  const name = words.slice(0, 4).join('_').replace(/^[^\p{L}]+/u, '').slice(0, 40);
  return name;
}

export function bookmarkDialog(view: EditorView): void {
  dialog((box, close) => {
    box.classList.add('bookmark-dialog');
    const name = el('input', { type: 'text', value: suggest(view), ariaLabel: 'Имя закладки', maxLength: 40 });
    const status = el('p', { role: 'status', class: 'bm-status' });
    const list = el('select', { size: 8, ariaLabel: 'Закладки', class: 'bm-list' });
    const order = el('select', { ariaLabel: 'Порядок' }, new Option('По имени', 'name'), new Option('По положению', 'pos'));
    const hidden = el('input', { type: 'checkbox', ariaLabel: 'Скрытые закладки' });
    const broken = el('div', { class: 'bm-broken' });
    const selected = () => list.value || name.value.trim();

    const draw = () => {
      const all = listBookmarks(view.state.doc).filter((b) => hidden.checked || !isHidden(b.name));
      if (order.value === 'name') all.sort((a, b) => a.name.localeCompare(b.name, 'ru'));
      const keep = list.value;
      list.replaceChildren(...all.map((b) => {
        const refs = linksTo(view.state.doc, b.name).length;
        return new Option(`${b.name}${b.context ? ' — ' + b.context : ''}${refs ? ` (ссылок: ${refs})` : ''}`, b.name);
      }));
      if (keep && all.some((b) => b.name === keep)) list.value = keep;
      const bad = brokenLinks(view.state.doc);
      broken.replaceChildren();
      if (bad.length) {
        broken.append(el('b', { textContent: `Ссылки без цели: ${bad.length}` }), ...bad.map((l) => el('button', {
          class: 'bm-broken-item', textContent: `«${l.text.slice(0, 40)}» → ${l.href}`,
          onclick: () => { close(); view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, l.from, l.to)).scrollIntoView()); view.focus(); },
        })));
      }
    };
    list.addEventListener('change', () => { name.value = list.value; status.textContent = ''; });
    list.addEventListener('dblclick', () => { if (list.value) { close(); gotoBookmark(view, list.value); } });
    order.addEventListener('change', draw);
    hidden.addEventListener('change', draw);

    const add = el('button', { class: 'primary', textContent: 'Добавить', onclick: () => {
      const n = name.value.trim();
      const exists = listBookmarks(view.state.doc).some((b) => b.name.toLowerCase() === n.toLowerCase());
      // Существующее имя — закладка переносится сюда, как в Word.
      const err = exists ? null : validateName(view.state.doc, n);
      if (err) { status.textContent = err; return; }
      view.dispatch(addBookmark(view.state, n));
      close();
      toast(exists ? `Закладка «${n}» перенесена сюда.` : `Закладка «${n}» добавлена.`);
      view.focus();
    } });
    const go = el('button', { textContent: 'Перейти', onclick: () => {
      const n = selected();
      close();
      if (!gotoBookmark(view, n)) toast(`Закладки «${n}» нет.`);
    } });
    const rename = el('button', { textContent: 'Переименовать', onclick: () => {
      const from = list.value, to = name.value.trim();
      if (!from) { status.textContent = 'Выберите закладку в списке и введите новое имя.'; return; }
      if (from === to) return;
      const err = validateName(view.state.doc, to, from);
      if (err) { status.textContent = err; return; }
      const refs = linksTo(view.state.doc, from).length;
      const tr = renameBookmark(view.state, from, to);
      if (tr) view.dispatch(tr);
      status.textContent = `Переименовано${refs ? `; ссылок обновлено: ${refs}` : ''}.`;
      list.value = '';
      draw();
      list.value = to;
    } });
    const del = el('button', { class: 'danger', textContent: 'Удалить', onclick: async () => {
      const n = list.value;
      if (!n) { status.textContent = 'Выберите закладку в списке.'; return; }
      const refs = linksTo(view.state.doc, n).length;
      const remove = () => { const tr = removeBookmark(view.state, n); if (tr) view.dispatch(tr); };
      if (refs) {
        // Окно подтверждения заменяет это окно; после ответа открываем «Закладку» заново.
        const ans = await choose('Удалить закладку?', `На закладку «${n}» ведут ссылки: ${refs}. После удаления они станут ссылками без цели.`, [
          { id: 'del', label: 'Удалить', danger: true }, { id: 'cancel', label: 'Отмена', primary: true }]);
        if (ans === 'del') remove();
        bookmarkDialog(view);
        return;
      }
      remove();
      status.textContent = `Закладка «${n}» удалена (отмена — Ctrl+Z).`;
      draw();
    } });
    box.append(el('h2', { textContent: 'Закладка' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Имя' }), name),
      list,
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Порядок' }), order, el('label', { class: 'bm-hidden' }, hidden, ' скрытые закладки')),
      status, broken,
      el('div', { class: 'dlg-actions' }, add, go, rename, del, el('button', { textContent: 'Закрыть', onclick: close })));
    draw();
    name.focus();
    name.select();
    box.addEventListener('keydown', (e) => { if (e.key === 'Enter' && document.activeElement === name) add.click(); });
  }, false);
}
