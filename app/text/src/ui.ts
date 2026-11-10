// Общие мелочи интерфейса: окна-диалоги, сборка элементов, всплывающее сообщение.

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

/** Куда вернуть фокус после закрытия окна (ставит main.ts: в текст документа). */
let restoreFocus: () => void = () => {};
export function setRestoreFocus(f: () => void): void {
  restoreFocus = f;
}

export function toast(text: string): void {
  const t = $('toast');
  t.textContent = text;
  t.hidden = false;
  clearTimeout((toast as unknown as { h?: number }).h);
  (toast as unknown as { h?: number }).h = window.setTimeout(() => { t.hidden = true; }, 3500);
}

/** Открыто ли окно (оверлей #dialog). Нужно, чтобы не пустить перезапуск за спиной окна. */
function isDialogOpen(): boolean {
  const ov = document.getElementById('dialog');
  return !!ov && !ov.hidden;
}

/** Перезапуск страницы: F5, Ctrl+R и Ctrl+Shift+R. */
const isReloadKey = (e: KeyboardEvent): boolean => {
  const k = e.key.toLowerCase();
  if (k === 'f5') return true;
  return (e.ctrlKey || e.metaKey) && k === 'r';
};

// Пока открыто окно, меню WebView по правой кнопке не показывается: в нём есть «Обновить»,
// которое перезапускает программу и уносит несохранённый документ. Заодно Ctrl+R и F5.
// В полях ввода меню остаётся — иначе пропадёт «Вставить» (правило как в ENS-Table).
// Слушатели навешиваются один раз на весь модуль, состояние берётся из разметки —
// так вложенные окна не оставляют за собой висящих обработчиков.
document.addEventListener('contextmenu', (e) => {
  if (!isDialogOpen()) return;
  if ((e.target as HTMLElement | null)?.closest('input, textarea')) return;
  e.preventDefault();
}, true);
document.addEventListener('keydown', (e) => {
  if (!isDialogOpen() || !isReloadKey(e)) return;
  e.preventDefault();
  e.stopPropagation();
}, true);

export function dialog(build: (box: HTMLElement, close: () => void) => void, small = true): void {
  const ov = $('dialog');
  const box = document.createElement('div');
  box.className = `modal${small ? ' small' : ''}`;
  const close = () => { ov.hidden = true; ov.replaceChildren(); document.removeEventListener('keydown', onKey, true); restoreFocus(); };
  const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') { e.stopPropagation(); close(); } };
  document.addEventListener('keydown', onKey, true);
  build(box, close);
  ov.replaceChildren(box);
  ov.hidden = false;
  box.querySelector<HTMLElement>('input,button.primary,.menu-close')?.focus();
}

export function el<K extends keyof HTMLElementTagNameMap>(tag: K, props: Partial<HTMLElementTagNameMap[K]> & { class?: string } = {}, ...kids: (Node | string)[]): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  const { class: cls, ...rest } = props as Record<string, unknown>;
  if (cls) e.className = cls as string;
  Object.assign(e, rest);
  e.append(...kids);
  return e;
}

/** Вопрос с несколькими кнопками; возвращает id выбранной (Esc — 'cancel'). */
export function choose<T extends string>(title: string, text: string, buttons: { id: T; label: string; primary?: boolean; danger?: boolean }[]): Promise<T | 'cancel'> {
  return new Promise((resolve) => {
    let done = false;
    const finish = (v: T | 'cancel', close?: () => void) => { if (done) return; done = true; close?.(); resolve(v); };
    dialog((box, close) => {
      const esc = (e: KeyboardEvent) => { if (e.key === 'Escape') { document.removeEventListener('keydown', esc, true); finish('cancel'); } };
      document.addEventListener('keydown', esc, true);
      box.append(
        el('h2', { textContent: title }),
        el('p', { textContent: text }),
        el('div', { class: 'dlg-actions' }, ...buttons.map((b) => el('button', {
          class: b.primary ? 'primary' : b.danger ? 'danger' : '', textContent: b.label,
          onclick: () => { document.removeEventListener('keydown', esc, true); finish(b.id, close); },
        }))),
      );
    });
  });
}
