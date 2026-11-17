// SPDX-License-Identifier: AGPL-3.0-or-later
// Выпадающее меню у кнопки ленты (регистр, межстрочный интервал): тот же вид, что у контекстного меню.

export interface PopItem { label: string; run: () => void; checked?: boolean; keys?: string; sample?: string }

let box: HTMLDivElement | null = null;

export function closePopMenu(): void {
  box?.remove();
  box = null;
}

/** Показать меню под элементом anchor (или в точке x,y). '-' — разделитель. */
export function popMenu(anchor: HTMLElement | { x: number; y: number }, items: (PopItem | '-')[], cls = ''): void {
  closePopMenu();
  box = document.createElement('div');
  box.className = `context-menu ctx-text pop-menu ${cls}`;
  box.setAttribute('role', 'menu');
  for (const it of items) {
    if (it === '-') { box.append(Object.assign(document.createElement('hr'), { className: 'ctx-sep' })); continue; }
    const b = document.createElement('button');
    b.type = 'button';
    b.setAttribute('role', 'menuitemradio');
    b.setAttribute('aria-checked', String(!!it.checked));
    const mark = document.createElement('i');
    mark.className = 'pm-check';
    mark.textContent = it.checked ? '✓' : '';
    const span = document.createElement('span');
    span.textContent = it.label;
    b.append(mark, span);
    if (it.keys) { const k = document.createElement('small'); k.textContent = it.keys; b.append(k); }
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => { closePopMenu(); it.run(); });
    box.append(b);
  }
  document.body.append(box);
  const r = anchor instanceof HTMLElement ? anchor.getBoundingClientRect() : { left: anchor.x, bottom: anchor.y };
  const { width, height } = box.getBoundingClientRect();
  box.style.left = `${Math.max(4, Math.min(r.left, window.innerWidth - width - 4))}px`;
  box.style.top = `${Math.max(4, Math.min(r.bottom + 2, window.innerHeight - height - 4))}px`;
  box.addEventListener('keydown', (e) => {
    const bs = [...box!.querySelectorAll<HTMLButtonElement>('button')];
    const i = bs.indexOf(document.activeElement as HTMLButtonElement);
    if (e.key === 'Escape') { e.preventDefault(); closePopMenu(); }
    else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); bs[(i + (e.key === 'ArrowDown' ? 1 : bs.length - 1)) % bs.length]?.focus(); }
  });
}

document.addEventListener('mousedown', (e) => { if (box && !box.contains(e.target as Node)) closePopMenu(); }, true);
window.addEventListener('resize', closePopMenu);
