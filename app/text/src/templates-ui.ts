// Окно «Создать из шаблона» и сохранение своего шаблона.

import type { Node as PMNode } from 'prosemirror-model';
import { importDocx, type DocMeta } from './docx';
import { builtInTemplates, getMine, listMine, removeMine, saveMine } from './templates';
import { dialog, el, toast } from './ui';

export interface TemplateHost {
  create: (doc: PMNode, meta: DocMeta, name: string) => void;
  /** Текущий документ в виде docx (без истории версий) — для сохранения своего шаблона. */
  current: () => { bytes: Uint8Array; name: string };
}

export function templatesDialog(host: TemplateHost, startMine = false): void {
  dialog((box, close) => {
    const list = el('div', { class: 'tpl-list' });
    const tabs = el('div', { class: 'rd-tabs' });
    let mine = startMine;

    const row = (title: string, hint: string, ...btns: HTMLElement[]) => el('div', { class: 'tpl-row' },
      el('div', { class: 'tpl-text' }, el('div', { class: 'tpl-title', textContent: title }), el('div', { class: 'tpl-hint', textContent: hint })), ...btns);

    const render = async () => {
      list.replaceChildren();
      for (const b of tabs.children) b.classList.toggle('on', (b as HTMLElement).dataset.tab === (mine ? 'mine' : 'std'));
      if (!mine) {
        for (const t of builtInTemplates()) list.append(row(t.title, t.hint, el('button', { class: 'primary', textContent: 'Создать', onclick: () => { const b = t.build(); close(); host.create(b.doc, b.meta, b.name); } })));
        return;
      }
      const mineList = await listMine();
      list.append(el('p', { textContent: 'Мои шаблоны лежат только на этом компьютере, внутри программы. Они не отправляются в интернет, не попадают в репозиторий и выпуски. Другой пользователь или другой компьютер их не увидит.' }));
      if (!mineList.length) list.append(el('p', { textContent: 'Пока пусто. Откройте документ и нажмите «Сохранить текущий документ как мой шаблон».' }));
      for (const t of mineList) {
        list.append(row(t.name, new Date(t.at).toLocaleString('ru-RU'),
          el('button', { class: 'primary', textContent: 'Создать', onclick: async () => {
            const rec = await getMine(t.id);
            if (!rec) return;
            try { const d = importDocx(rec.bytes); close(); host.create(d.doc, d.meta, `${t.name}`); } catch (e) { toast(`Шаблон не открылся: ${(e as Error).message}`); }
          } }),
          el('button', { textContent: 'Удалить', onclick: async () => { await removeMine(t.id); void render(); } })));
      }
    };

    const tabBtn = (label: string, key: string, isMine: boolean) => {
      const b = el('button', { textContent: label, onclick: () => { mine = isMine; void render(); } });
      b.dataset.tab = key;
      return b;
    };
    tabs.append(tabBtn('Готовые шаблоны', 'std', false), tabBtn('Мои шаблоны', 'mine', true));

    box.append(
      el('h2', { textContent: 'Создать из шаблона' }), tabs, list,
      el('div', { class: 'dlg-actions' },
        el('button', { textContent: 'Сохранить текущий документ как мой шаблон…', onclick: () => { close(); saveTemplateDialog(host); } }),
        el('button', { class: 'primary', textContent: 'Закрыть', onclick: close })),
    );
    void render();
  }, false);
}

export function saveTemplateDialog(host: TemplateHost): void {
  dialog((box, close) => {
    const cur = host.current();
    const name = el('input', { type: 'text', value: cur.name, placeholder: 'Название шаблона' });
    const go = async () => {
      const v = name.value.trim();
      if (!v) return;
      try { await saveMine(v, cur.bytes); close(); toast(`Шаблон «${v}» сохранён — только на этом компьютере`); }
      catch (e) { toast(`Не удалось сохранить шаблон: ${(e as Error).message}`); }
    };
    box.addEventListener('keydown', (e) => { if (e.key === 'Enter') void go(); });
    box.append(
      el('h2', { textContent: 'Сохранить как мой шаблон' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Название' }), name),
      el('p', { textContent: 'Шаблон хранится только на этом компьютере, внутри программы. Никуда не отправляется. Чтобы перенести его на другой компьютер, сохраните документ обычным файлом .docx.' }),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), el('button', { class: 'primary', textContent: 'Сохранить', onclick: () => void go() })),
    );
    name.select();
  });
}
