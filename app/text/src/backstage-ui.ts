// SPDX-License-Identifier: AGPL-3.0-or-later
// Меню «Файл» ENS-Text (Backstage), как в ENS-Table и Office 2021–2024: слева разделы, справа содержимое.
// Вид и классы общие с ENS-Table (styles.css: .backstage, .bs-*), содержимое своё: документы, шаблоны, параметры
// текстового редактора. Общий код ENS-Table не меняется.

import { command, run } from '../../table/src/commands';
import { icon } from '../../table/src/ribbon';
import { saveSettings, settings } from '../../table/src/settings';
import logoUrl from '../../table/src/assets/logo-128.png?inline';
import { builtInTemplates } from './templates';
import { APP_VERSION } from './version';

export interface FileMenuHost {
  info(): { title: string; path: string; dirty: boolean; words: number; pages: number; versions: number };
  isDesktop: boolean;
  /** Открыть файл с диска по пути (настольная версия). */
  openPath(path: string): void;
  /** Создать документ: пустой или из встроенного шаблона (id), с вопросом о несохранённом. */
  create(templateId: string | null): void;
  /** Типограф при наборе. */
  liveQuotes(): boolean;
  setLiveQuotes(on: boolean): void;
  /** Настройки поменялись: перерисовать то, что от них зависит. */
  applySettings(): void;
}

type Section = 'info' | 'new' | 'open' | 'print' | 'options';

interface FileRef { path: string; title: string; at?: number }
const RECENT = 'ens-text.recent';
const PINNED = 'ens-text.pinned';
const load = (k: string): FileRef[] => { try { const v = JSON.parse(localStorage.getItem(k) ?? '[]'); return Array.isArray(v) ? v : []; } catch { return []; } };
const store = (k: string, v: FileRef[]) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* без хранилища — до закрытия окна */ } };
const titleOf = (path: string) => path.split(/[\\/]/).pop()!.replace(/\.docx$/i, '');

/** Файл открыт или сохранён: в начало списка «Последние» (не больше 20). */
export function rememberFile(path: string): void {
  if (!/[\\/]/.test(path)) return; // в браузере пути нет — список не ведём
  store(RECENT, [{ path, title: titleOf(path), at: Date.now() }, ...load(RECENT).filter((r) => r.path !== path)].slice(0, 20));
}

const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);

let host: FileMenuHost;
const root = document.createElement('div');
root.className = 'backstage text-backstage';
root.hidden = true;
root.setAttribute('role', 'dialog');
root.setAttribute('aria-label', 'Меню «Файл»');

export function initFileMenu(h: FileMenuHost): void {
  host = h;
  document.body.append(root);
  document.addEventListener('keydown', (e) => {
    if (root.hidden || e.key !== 'Escape') return;
    e.preventDefault();
    e.stopPropagation();
    hideFileMenu();
  }, true);
}

export const fileMenuOpen = (): boolean => !root.hidden;

export function hideFileMenu(): void {
  if (root.hidden) return;
  root.hidden = true;
  document.dispatchEvent(new CustomEvent('ens-file-menu', { detail: false }));
}

export function showFileMenu(section: Section = 'info'): void {
  root.innerHTML = `
    <nav class="bs-nav">
      <button class="bs-back" data-go="back" title="Назад (Esc)" aria-label="Закрыть меню">${icon('arrow_left', 24)}</button>
      <button data-go="info">${icon('info', 20)}Сведения</button>
      <button data-go="new">${icon('document_add', 20)}Создать</button>
      <button data-go="open">${icon('folder_open', 20)}Открыть</button>
      <hr>
      <button data-go="save">${icon('save', 20)}Сохранить</button>
      <button data-go="saveAs">${icon('save_edit', 20)}Сохранить как</button>
      <button data-go="print">${icon('print', 20)}Печать</button>
      <button data-go="close">${icon('dismiss', 20)}Закрыть</button>
      <div class="bs-spacer"></div>
      <button data-go="options">${icon('settings', 20)}Параметры</button>
      <div class="bs-about">
        <img class="bs-logo" src="${logoUrl}" alt="ENS-Office" width="64" height="64"><br>
        ENS-Office · ENS-Text ${esc(APP_VERSION)}<br>
        Автор: Н. Едренников<br>
        Лицензия: GNU AGPL-3.0-or-later<br>
        или коммерческая лицензия<br>
        Обратная связь: <a href="mailto:n.edrennikov@yandex.ru">n.edrennikov@yandex.ru</a><br>
        Новости: <a href="https://t.me/ENS_Office" target="_blank" rel="noopener">t.me/ENS_Office</a>
      </div>
    </nav>
    <section class="bs-body"></section>`;
  root.querySelectorAll<HTMLButtonElement>('[data-go]').forEach((b) => b.addEventListener('click', () => go(b.dataset.go!)));
  root.hidden = false;
  document.dispatchEvent(new CustomEvent('ens-file-menu', { detail: true }));
  render(section);
  root.querySelector<HTMLElement>(`[data-go="${section}"]`)?.focus();
}

function go(target: string): void {
  const now = (id: string) => { hideFileMenu(); run(id); };
  switch (target) {
    case 'back': hideFileMenu(); return;
    case 'save': now('file.save'); return;
    case 'saveAs': now('file.saveAs'); return;
    case 'close': now('file.close'); return;
    default: render(target as Section);
  }
}

/** Кнопка-строка, запускающая команду (если она есть в программе). */
function wide(iconName: string, label: string, id: string): HTMLButtonElement {
  const b = document.createElement('button');
  b.className = 'bs-wide';
  b.innerHTML = `${icon(iconName, 24)}<span>${esc(label)}</span>`;
  b.disabled = !command(id);
  b.addEventListener('click', () => { hideFileMenu(); run(id); });
  return b;
}

function render(section: Section): void {
  root.querySelectorAll('[data-go]').forEach((b) => b.classList.toggle('active', (b as HTMLElement).dataset.go === section));
  const body = root.querySelector<HTMLElement>('.bs-body')!;
  body.replaceChildren();
  const h1 = (t: string) => { const h = document.createElement('h1'); h.textContent = t; body.append(h); };
  const h2 = (t: string) => { const h = document.createElement('h2'); h.textContent = t; body.append(h); };
  const p = (t: string, cls = 'bs-muted') => { const e = document.createElement('p'); e.className = cls; e.textContent = t; body.append(e); return e; };
  const actions = (...b: HTMLElement[]) => { const d = document.createElement('div'); d.className = 'bs-actions'; d.append(...b); body.append(d); };

  if (section === 'info') {
    const d = host.info();
    h1('Сведения');
    h2(`${d.title}${d.dirty ? ' — есть несохранённые изменения' : ''}`);
    p(d.path || (host.isDesktop ? 'Документ ещё не сохранялся' : 'Документ открыт в браузере: «Сохранить» скачивает файл .docx'));
    p(`Страниц: ${d.pages} · слов: ${d.words} · версий в файле: ${d.versions}`);
    actions(
      wide('history', 'История версий: вернуть любой момент, сравнить', 'ver.history'),
      wide('document', 'Совместимость с Word: что может выглядеть иначе', 'review.compat'),
      wide('checkmark', 'Проверка оформления по ГОСТ', 'review.check'),
      wide('document', 'Журнал регистрации документов', 'file.journal'),
    );
  } else if (section === 'new') {
    h1('Создать');
    const grid = document.createElement('div');
    grid.className = 'bs-tiles';
    const tile = (title: string, hint: string, id: string | null) => {
      const b = document.createElement('button');
      b.className = 'bs-tile';
      b.title = hint;
      b.innerHTML = `<span class="bs-page"></span><span>${esc(title)}</span>`;
      b.addEventListener('click', () => { hideFileMenu(); host.create(id); });
      grid.append(b);
    };
    tile('Новый документ', 'Пустой документ A4', null);
    for (const t of builtInTemplates()) tile(t.title, t.hint, t.id);
    body.append(grid);
    actions(wide('document', 'Мои шаблоны…', 'file.templates'), wide('document', 'Служебный документ (мастер)…', 'file.wizard'), wide('save', 'Сохранить текущий как мой шаблон…', 'file.saveTemplate'));
  } else if (section === 'open') {
    h1('Открыть');
    actions(wide('folder_open', 'Обзор…', 'file.open'));
    if (!host.isDesktop) { p('В браузере список последних файлов не ведётся: у страницы нет доступа к путям на диске. В настольной версии здесь будут закреплённые и последние документы.'); return; }
    const pinned = load(PINNED);
    const row = (r: FileRef, isPinned: boolean) => {
      const wrap = document.createElement('div');
      wrap.className = 'bs-recent-row';
      const b = document.createElement('button');
      b.className = 'bs-recent-item';
      b.innerHTML = `${icon('document', 24)}<span><b>${esc(r.title)}</b><br><small>${esc(r.path)}</small></span>`;
      b.addEventListener('click', () => { hideFileMenu(); host.openPath(r.path); });
      const star = document.createElement('button');
      star.className = 'bs-pin';
      star.textContent = isPinned ? '★' : '☆';
      star.title = isPinned ? 'Открепить' : 'Закрепить';
      star.addEventListener('click', () => {
        store(PINNED, isPinned ? pinned.filter((x) => x.path !== r.path) : [{ path: r.path, title: r.title }, ...pinned]);
        render('open');
      });
      wrap.append(b, star);
      return wrap;
    };
    h2('Закреплённые');
    const pinBox = document.createElement('div');
    pinBox.className = 'bs-recent';
    if (!pinned.length) pinBox.append(Object.assign(document.createElement('p'), { className: 'bs-muted', textContent: 'Звёздочка у документа в «Последних» закрепляет его здесь.' }));
    for (const r of pinned) pinBox.append(row(r, true));
    body.append(pinBox);
    h2('Последние');
    const box = document.createElement('div');
    box.className = 'bs-recent';
    const recent = load(RECENT);
    if (!recent.length) box.append(Object.assign(document.createElement('p'), { className: 'bs-muted', textContent: 'Здесь появятся документы, которые вы открывали или сохраняли.' }));
    for (const r of recent) box.append(row(r, pinned.some((x) => x.path === r.path)));
    body.append(box);
  } else if (section === 'print') {
    h1('Печать');
    p('Документ печатается так, как выглядит на экране: листы, поля, колонтитулы. Чтобы получить PDF, в окне печати выберите «Сохранить как PDF» или «Печать в файл».', 'bs-muted');
    actions(wide('print', 'Печать… (Ctrl+P)', 'file.print'));
    h2('Страница');
    actions(wide('document', 'Поля, размер и ориентация — вкладка «Макет»', 'page.margins'));
  } else if (section === 'options') {
    h1('Параметры');
    const row = (label: string, input: HTMLElement, after = '') => {
      const l = document.createElement('label');
      l.className = 'bs-row';
      if (input instanceof HTMLInputElement && input.type === 'checkbox') l.append(input, ` ${label}`);
      else l.append(`${label} `, input, after);
      body.append(l);
    };
    const check = (key: string, on: boolean, set: (v: boolean) => void) => {
      const c = document.createElement('input');
      c.type = 'checkbox';
      c.checked = on;
      c.dataset.set = key;
      c.addEventListener('change', () => { set(c.checked); host.applySettings(); });
      return c;
    };
    const num = (key: string, value: number, min: number, max: number, set: (v: number) => void) => {
      const n = document.createElement('input');
      n.type = 'number';
      n.min = String(min);
      n.max = String(max);
      n.value = String(value);
      n.dataset.set = key;
      n.style.width = '64px';
      n.addEventListener('change', () => { const v = Math.min(max, Math.max(min, Math.round(Number(n.value) || 0))); n.value = String(v); set(v); saveSettings(); host.applySettings(); });
      return n;
    };
    h2('Автор');
    const name = document.createElement('input');
    name.dataset.set = 'userName';
    name.placeholder = 'как в системе';
    name.style.width = '240px';
    name.value = settings.userName;
    name.addEventListener('change', () => { settings.userName = name.value.trim(); saveSettings(); host.applySettings(); });
    row('Ваше имя (журнал изменений и версии документа)', name);
    h2('Сохранение');
    row('Напоминать о сохранении через', num('saveReminderMinutes', settings.saveReminderMinutes, 0, 240, (v) => { settings.saveReminderMinutes = v; }), ' мин. (0 — не напоминать)');
    row('Версия в истории документа не реже чем раз в', num('autosaveMinutes', settings.autosaveMinutes, 1, 120, (v) => { settings.autosaveMinutes = v; }), ' мин.');
    h2('Правка');
    row('Типограф при наборе: «ёлочки», тире', check('liveQuotes', host.liveQuotes(), (v) => host.setLiveQuotes(v)));
    const toggles: [string, string][] = [['view.marks', 'Показывать скрытые знаки (¶ · →)'], ['view.ruler', 'Показывать линейку']];
    for (const [id, label] of toggles) {
      const c = command(id);
      if (!c) continue;
      row(label, check(id, !!c.checked?.(), () => run(id)));
    }
    actions(wide('document', 'Словарь: слова, которые не подчёркиваются…', 'review.dictionary'), wide('paint_brush', 'Тема оформления…', 'view.theme'));
    p('Параметры хранятся на этом компьютере.');
  }
}
