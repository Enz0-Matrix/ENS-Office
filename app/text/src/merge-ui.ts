// Три локальных файла, предварительный просмотр конфликтов, затем новый документ.
import type { Node as PMNode } from 'prosemirror-model';
import { importDocx, type DocMeta } from './docx';
import { baseOverlap, describeMeta, mergeKey, planMerge, resolveMerge, type MergeChoice, type MergePart } from './merge';
import { dialog, el } from './ui';

interface Copy { doc: PMNode; meta: DocMeta; name: string }
interface Host { create(doc: PMNode, meta: DocMeta, name: string): void }

export function mergeDialog(host: Host): void {
  dialog((box, close) => {
    box.classList.add('merge-dialog');
    const status = el('p', { role: 'status' });
    const results = el('div', { class: 'merge-results' });
    const inputs = ['Общая основа', 'Моя копия', 'Копия коллеги'].map((label, i) => {
      const input = el('input', { type: 'file', accept: '.docx', ariaLabel: label });
      input.dataset.mergeFile = String(i);
      input.addEventListener('change', () => { generation++; results.replaceChildren(); status.textContent = ''; });
      return { input, row: el('div', { class: 'dlg-row' }, el('label', { textContent: label }), input) };
    });
    let generation = 0;
    const compare = el('button', { textContent: 'Сравнить копии', onclick: async () => {
      const token = ++generation;
      results.replaceChildren();
      const files = inputs.map(({ input }) => input.files?.[0]);
      if (files.some((f) => !f)) { status.textContent = 'Выберите все три файла DOCX.'; return; }
      status.textContent = 'Чтение локальных файлов…';
      try {
        const copies: Copy[] = await Promise.all(files.map(async (file) => {
          const f = file!;
          if (!/\.docx$/i.test(f.name)) throw new Error('Нужны файлы DOCX.');
          if (f.size > 32 * 1024 * 1024) throw new Error('Файл превышает 32 МБ.');
          const imported = importDocx(new Uint8Array(await f.arrayBuffer()));
          if (imported.meta.lossy.length) throw new Error(`${f.name}: есть неподдерживаемые элементы (${imported.meta.lossy.join('; ')}). Слияние остановлено, файлы не изменены.`);
          return { doc: imported.doc, meta: imported.meta, name: f.name };
        }));
        if (token !== generation || !box.isConnected) return;
        const [base, mine, theirs] = copies;
        const plan = planMerge(base.doc, mine.doc, theirs.doc);
        for (const [who, copy] of [['Моя копия', mine], ['Копия коллеги', theirs]] as const) {
          if (base.doc.childCount >= 4 && baseOverlap(base.doc, copy.doc) < 0.3) results.append(el('p', { role: 'alert',
            textContent: `«${who}» почти не совпадает с общей основой. Проверьте, что файлы выбраны в правильных полях.` }));
        }
        const choices = new Map<number, MergeChoice>();
        const labels: Record<MergeChoice, string> = { base: 'Основа', mine: 'Моя копия', theirs: 'Копия коллеги' };
        let resultMeta = mine.meta;
        const b = mergeKey(base.meta), m = mergeKey(mine.meta), t = mergeKey(theirs.meta);
        const metaConflict = m !== t && b !== m && b !== t;
        if (m === b) resultMeta = theirs.meta;
        let metaChoice: MergeChoice | undefined;
        const apply = el('button', { class: 'primary', textContent: 'Создать объединённый документ', disabled: true, onclick: () => {
          try {
            const doc = resolveMerge(plan, choices);
            const meta = metaConflict ? copies[{ base: 0, mine: 1, theirs: 2 }[metaChoice!]].meta : resultMeta;
            close();
            host.create(doc, structuredClone(meta), mine.name.replace(/\.docx$/i, '') + ' — объединённый');
          } catch (e) { status.textContent = (e as Error).message; }
        } });
        const sync = () => {
          const remaining = plan.conflicts - choices.size + (metaConflict && !metaChoice ? 1 : 0);
          status.textContent = `Конфликтов: ${plan.conflicts + Number(metaConflict)}. Осталось выбрать: ${remaining}.`;
          apply.disabled = remaining !== 0;
        };
        const select = (label: string, choose: (value: MergeChoice) => void) => {
          const control = el('select', { ariaLabel: label }, el('option', { value: '', textContent: 'Выберите вариант…' }),
            ...(['base', 'mine', 'theirs'] as const).map((key) => el('option', { value: key, textContent: labels[key] })));
          // Пустой пункт служит подсказкой и недоступен после выбора. Выбран он явно: иначе браузер
          // может показать первый доступный вариант, и решение будет выглядеть принятым, хотя его нет.
          control.options[0].disabled = true;
          control.selectedIndex = 0;
          control.addEventListener('change', () => { if (control.value) choose(control.value as MergeChoice); sync(); });
          return control;
        };
        const render = (parts: MergePart[]) => {
          for (const part of parts) {
            if (part.kind === 'container') { render(part.children); continue; }
            if (part.kind !== 'conflict') continue;
            const card = el('section', { class: 'merge-conflict' }, el('h3', { textContent: `Конфликт ${part.id + 1}` }));
            for (const key of ['base', 'mine', 'theirs'] as const) card.append(el('b', { textContent: labels[key] }),
              el('pre', { textContent: part[key].length ? part[key].map((n) => n.textContent || `[${n.type.name}]`).join('\n') : '(удалено / нет вставки)' }));
            card.append(el('details', {}, el('summary', { textContent: 'Форматирование и структура (JSON)' }),
              el('pre', { textContent: describeMeta({ base: part.base.map((n) => n.toJSON()), mine: part.mine.map((n) => n.toJSON()), theirs: part.theirs.map((n) => n.toJSON()) }) })),
              select(`Решение конфликта ${part.id + 1}`, (value) => choices.set(part.id, value)));
            results.append(card);
          }
        };
        render(plan.parts);
        if (metaConflict) results.append(el('section', { class: 'merge-conflict' },
          el('h3', { textContent: 'Параметры страницы, шрифты по умолчанию и колонтитулы' }),
          ...copies.map((copy, i) => el('details', {}, el('summary', { textContent: labels[(['base', 'mine', 'theirs'] as const)[i]] }), el('pre', { textContent: describeMeta(copy.meta) }))),
          select('Параметры объединённого документа', (value) => { metaChoice = value; })));
        results.append(el('p', { textContent: 'Независимые изменения объединяются автоматически. При конфликте выбирается целый фрагмент; таблицы и списки сравниваются по вложенным абзацам. Результат открывается новым документом; исходные файлы не перезаписываются.' }), apply);
        sync();
      } catch (e) { if (token === generation) status.textContent = 'Не удалось объединить: ' + (e as Error).message; }
    } });
    box.append(el('h2', { textContent: 'Объединить копии без сети' }),
      el('p', { textContent: 'Выберите общую основу до правок и две её изменённые копии. Сравнение учитывает текст, форматирование и структуру. Файлы обрабатываются локально.' }),
      ...inputs.map(({ row }) => row), compare, status, results,
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close })));
  }, false);
}
