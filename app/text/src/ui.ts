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
  box.querySelector<HTMLElement>('input,button.primary')?.focus();
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
