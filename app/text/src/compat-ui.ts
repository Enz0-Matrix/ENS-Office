// Окно «Совместимость с Word».

import { Selection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { DocMeta } from './docx';
import { analyze, verdict } from './compat';
import { dialog, el } from './ui';

export interface CompatHost { view: () => EditorView; meta: () => DocMeta; pages: () => number }

export function compatDialog(host: CompatHost): void {
  dialog((box, close) => {
    const findings = analyze(host.view().state.doc, host.meta(), host.pages());
    const warns = findings.filter((f) => f.level === 'warn');
    const row = (f: (typeof findings)[number]) => el('div', { class: `cmp-row ${f.level}` },
      el('span', { class: 'cmp-dot', textContent: f.level === 'warn' ? '!' : 'i' }),
      el('div', { class: 'cmp-text' },
        el('div', { class: 'cmp-title', textContent: f.count > 1 && f.level === 'warn' ? `${f.title} (${f.count})` : f.title }),
        el('div', { class: 'cmp-detail', textContent: f.detail })),
      ...(f.pos != null ? [el('button', { textContent: 'Показать', onclick: () => {
        const v = host.view();
        const pos = Math.min(f.pos!, v.state.doc.content.size);
        v.dispatch(v.state.tr.setSelection(Selection.near(v.state.doc.resolve(pos))).scrollIntoView());
        close();
        v.focus();
      } })] : []));
    box.append(
      el('h2', { textContent: 'Совместимость с Word' }),
      el('p', { class: `cmp-verdict ${warns.length ? 'warn' : 'ok'}`, textContent: verdict(findings) }),
      el('div', { class: 'cmp-list' }, ...findings.map(row)),
      el('p', { textContent: 'Это проверка по правилам, а не запуск Word: она называет типичные причины «поплыла вёрстка». Живой Word, Р7-Офис и LibreOffice на ваших файлах стоит проверить глазами.' }),
      el('div', { class: 'dlg-actions' }, el('button', { class: 'primary', textContent: 'Закрыть', onclick: close })),
    );
  }, false);
}
