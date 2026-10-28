// Кнопка выбора диапазона мышью, как в Excel: окно сворачивается в плавающую строку,
// диапазон выделяют прямо в таблице; Enter или кнопка «Готово» вставляют его в поле.

export interface RangePickHost {
  /** Текущее выделение в таблице в виде ссылки; format решает, как её записать в поле. */
  selection(format: 'plain' | 'sheet'): string;
  focusGrid(): void;
}

let host: RangePickHost | null = null;
export function initRangePicker(h: RangePickHost): void { host = h; }

/** Добавляет к полю кнопку «Выбрать диапазон в таблице». format: plain — «A1:C5», sheet — «Лист1!$A$1:$C$5». */
export function addRangePicker(input: HTMLInputElement, format: 'plain' | 'sheet' = 'plain'): void {
  const btn = Object.assign(document.createElement('button'), {
    type: 'button', className: 'range-pick', title: 'Выбрать диапазон в таблице', textContent: '▦',
  });
  btn.dataset.act = 'range-pick';
  input.insertAdjacentElement('afterend', btn);
  btn.addEventListener('click', (e) => { e.stopPropagation(); pick(input, format); });
}

function pick(input: HTMLInputElement, format: 'plain' | 'sheet'): void {
  if (!host) return;
  const overlay = document.getElementById('recovery')!;
  const bar = document.createElement('div');
  bar.className = 'range-bar';
  bar.innerHTML = '<span>Выделите диапазон в таблице:</span><input readonly><button class="primary">Готово</button><button data-cancel>Отмена</button>';
  document.body.appendChild(bar);
  const shown = bar.querySelector<HTMLInputElement>('input')!;
  overlay.classList.add('picking');
  const refresh = () => { try { shown.value = host!.selection(format); } catch { /* нет выделения */ } };
  refresh();
  const timer = window.setInterval(refresh, 100);
  const finish = (apply: boolean) => {
    window.clearInterval(timer);
    document.removeEventListener('keydown', onKey, true);
    if (apply) { refresh(); input.value = shown.value; input.dispatchEvent(new Event('change', { bubbles: true })); }
    bar.remove();
    overlay.classList.remove('picking');
    input.focus();
  };
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Enter' || e.key === 'Escape') {
      e.preventDefault(); e.stopPropagation();
      finish(e.key === 'Enter');
    }
  };
  document.addEventListener('keydown', onKey, true);
  bar.addEventListener('click', (e) => {
    const t = e.target as HTMLElement;
    if (t.tagName === 'BUTTON') finish(!t.hasAttribute('data-cancel'));
  });
}
