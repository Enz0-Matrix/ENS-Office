// Меню «Файл» (Backstage), как в Office 2021–2024: слева разделы, справа содержимое.

import { FONTS, icon } from './ribbon';
import { settings, saveSettings } from './settings';
import { isDesktop } from './platform';
import { DONATE_URL } from './modules';
import { DEFAULT_PRINT, FOOTER_PRESETS, HEADER_PRESETS, marginsOf, resetPreviewPage, type PrintOptions } from './print';

let printOpts: PrintOptions = { ...DEFAULT_PRINT };
const isWindows = navigator.userAgent.includes('Windows');
const hfOptions = (presets: [string, string][], v: string) => {
  const known = presets.some(([c]) => c === v);
  return presets.map(([c, t]) => `<option value="${esc(c)}" ${c === v ? 'selected' : ''}>${esc(t)}</option>`).join('') +
    `<option value="*" ${known ? '' : 'selected'}>Свой текст…</option>`;
};

export interface BackstageHost {
  docInfo(): { title: string; path: string; dirty: boolean };
  recent(): { path: string; title: string }[];
  newBook(): void;
  open(): void;
  openPath(path: string): void;
  save(): void;
  saveAs(): void;
  close(): void;
  applySettings(): void;
  printPreview(el: HTMLElement, opts: PrintOptions): void;
  print(opts: PrintOptions): void;
  /** Parameters of the active sheet (saved in the book). */
  printOptions(): PrintOptions;
  showChangeLog(): void;
  /** Область печати: задать по выделению / убрать; текст для показа или ''. */
  printArea(): string;
  setPrintArea(clear: boolean): void;
  /** Сквозные строки и столбцы: тексты «1:3» и «A:B» ('' — нет); set возвращает текст ошибки или ''. */
  printTitles(): { rows: string; cols: string };
  setPrintTitles(rows: string, cols: string): string;
  savePrintOptions(opts: PrintOptions): void;
  recoveryFolder(): Promise<string>;
  chooseRecoveryFolder(): Promise<void>;
}

type Section = 'info' | 'new' | 'open' | 'print' | 'options';

import { APP_VERSION } from './version';

const root = document.createElement('div');
root.className = 'backstage';
root.hidden = true;
document.body.append(root);

let host: BackstageHost;

export function initBackstage(h: BackstageHost): void {
  host = h;
  document.addEventListener('keydown', (e) => {
    if (!root.hidden && e.key === 'Escape') hideBackstage();
  });
}

export function hideBackstage(): void {
  root.hidden = true;
}

export function showBackstage(section: Section = 'info'): void {
  root.innerHTML = `
    <nav class="bs-nav">
      <button class="bs-back" data-go="back" title="Назад (Esc)">${icon('arrow_left', 24)}</button>
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
      ${DONATE_URL ? `<a class="bs-donate" href="${DONATE_URL}" target="_blank" rel="noopener">♥ Поддержать проект</a>` : ''}
      <div class="bs-about">
        ENS-Office · ENS-Table ${APP_VERSION}<br>
        Автор: Н. Едренников<br>
        Обратная связь: <a href="mailto:n.edrennikov@yandex.ru">n.edrennikov@yandex.ru</a>
      </div>
    </nav>
    <section class="bs-body"></section>`;
  root.querySelectorAll<HTMLButtonElement>('[data-go]').forEach((b) =>
    b.addEventListener('click', () => go(b.dataset.go!)));
  root.hidden = false;
  render(section);
}

function go(target: string): void {
  switch (target) {
    case 'back': hideBackstage(); return;
    case 'save': hideBackstage(); host.save(); return;
    case 'saveAs': hideBackstage(); host.saveAs(); return;
    case 'close': hideBackstage(); host.close(); return;
    default: render(target as Section);
  }
}

function render(section: Section): void {
  root.querySelectorAll('[data-go]').forEach((b) => b.classList.toggle('active', (b as HTMLElement).dataset.go === section));
  const body = root.querySelector<HTMLElement>('.bs-body')!;
  const doc = host.docInfo();
  if (section === 'info') {
    body.innerHTML = `
      <h1>Сведения</h1>
      <h2>${esc(doc.title)}${doc.dirty ? ' — есть несохранённые изменения' : ''}</h2>
      <p class="bs-muted">${doc.path ? esc(doc.path) : 'Книга ещё не сохранялась'}</p>
      ${isDesktop ? `<p class="bs-muted">Временные файлы для восстановления после сбоя: <span class="bs-rec">…</span></p>` : ''}
      <button class="bs-wide" data-act="changelog">${icon('history', 24) || icon('document', 24)}Журнал изменений: кто, что и когда менял</button>`;
    body.querySelector('[data-act="changelog"]')!.addEventListener('click', () => { hideBackstage(); host.showChangeLog(); });
    if (isDesktop) void host.recoveryFolder().then((p) => { const el = body.querySelector('.bs-rec'); if (el) el.textContent = p; });
  } else if (section === 'new') {
    body.innerHTML = `<h1>Создать</h1>
      <button class="bs-tile" data-act="blank"><span class="bs-blank"></span>Пустая книга</button>`;
    body.querySelector('[data-act="blank"]')!.addEventListener('click', () => { hideBackstage(); host.newBook(); });
  } else if (section === 'open') {
    const recent = host.recent();
    body.innerHTML = `<h1>Открыть</h1>
      <button class="bs-wide" data-act="browse">${icon('folder_open', 24)}Обзор…</button>
      <h2>Последние</h2>
      <div class="bs-recent">${recent.length ? '' : '<p class="bs-muted">Здесь появятся книги, которые вы открывали.</p>'}</div>`;
    body.querySelector('[data-act="browse"]')!.addEventListener('click', () => { hideBackstage(); host.open(); });
    const list = body.querySelector('.bs-recent')!;
    for (const r of recent) {
      const b = document.createElement('button');
      b.className = 'bs-recent-item';
      b.innerHTML = `${icon('document', 24)}<span><b>${esc(r.title)}</b><br><small>${esc(r.path)}</small></span>`;
      b.addEventListener('click', () => { hideBackstage(); host.openPath(r.path); });
      list.append(b);
    }
  } else if (section === 'print') {
    body.innerHTML = `<h1>Печать</h1>
      <div class="print-layout">
        <div class="print-options">
          <div class="print-btns">
            <button class="print-go" data-act="print">${icon('print', 20)}Печать…</button>
            <button class="print-go" data-act="pdf">${icon('save', 20)}Сохранить в PDF…</button>
          </div>
          <p class="bs-muted">PDF: в окне печати выберите принтер «${isWindows ? 'Microsoft Print to PDF' : 'Печать в файл (PDF)'}» или «Сохранить как PDF».</p>
          <label>Что печатать
            <select data-p="what"><option value="sheet">Активный лист</option><option value="selection">Выделенный диапазон</option></select></label>
          <label>Ориентация
            <select data-p="landscape"><option value="false">Книжная</option><option value="true">Альбомная</option></select></label>
          <label>Размер бумаги
            <select data-p="paper"><option value="A4">A4 (210 × 297 мм)</option><option value="A3">A3 (297 × 420 мм)</option>
            <option value="A5">A5 (148 × 210 мм)</option><option value="Letter">Letter (216 × 279 мм)</option></select></label>
          <label>Масштаб
            <select data-p="fit"><option value="none">Текущий (100%)</option><option value="width">Вписать все столбцы на одну страницу</option>
            <option value="page">Вписать лист на одну страницу</option></select></label>
          <label>Поля
            <select data-p="margins"><option value="normal">Обычные</option><option value="narrow">Узкие</option><option value="wide">Широкие</option>
            <option value="custom">Настраиваемые…</option></select></label>
          <div class="print-margins" hidden>
            ${['Верхнее', 'Правое', 'Нижнее', 'Левое'].map((t, i) => `<label>${t}, мм <input type="number" min="0" max="100" step="1" data-m="${i}"></label>`).join('')}
          </div>
          <label>Верхний колонтитул
            <select data-hf="header">${hfOptions(HEADER_PRESETS, printOpts.header)}</select></label>
          <input class="print-hf-custom" data-hfc="header" placeholder="&L слева &C по центру &R справа; &P — стр., &N — всего" hidden>
          <label>Нижний колонтитул
            <select data-hf="footer">${hfOptions(FOOTER_PRESETS, printOpts.footer)}</select></label>
          <input class="print-hf-custom" data-hfc="footer" placeholder="Например: &CСтраница &P из &N" hidden>
          <div class="print-area"><span class="bs-muted">Область печати: <b class="pa-text"></b></span>
            <button data-act="pa-set">Задать по выделению</button><button data-act="pa-clear">Убрать</button></div>
          <div class="print-titles">
            <label>Сквозные строки <input data-t="rows" placeholder="например, 1:3" autocomplete="off"></label>
            <label>Сквозные столбцы <input data-t="cols" placeholder="например, A:B" autocomplete="off"></label>
            <span class="bs-muted pt-note"></span>
          </div>
          <label class="check"><input type="checkbox" data-p="centerH"> По центру страницы</label>
          <label class="check"><input type="checkbox" data-p="gridlines"> Печатать сетку</label>
          <label class="check"><input type="checkbox" data-p="repeatHeader"> Заголовок на каждой странице (первая или закреплённые строки)</label>
          <p class="bs-muted">Параметры страницы сохраняются в книге для каждого листа.</p>
        </div>
        <div class="print-preview"></div>
      </div>`;
    const preview = body.querySelector<HTMLElement>('.print-preview')!;
    printOpts = { ...host.printOptions(), what: printOpts.what };
    resetPreviewPage();
    // Колонтитулы: готовые варианты или свой текст в записи Excel.
    for (const k of ['header', 'footer'] as const) {
      const presets = k === 'header' ? HEADER_PRESETS : FOOTER_PRESETS;
      const sel = body.querySelector<HTMLSelectElement>(`[data-hf="${k}"]`)!;
      const inp = body.querySelector<HTMLInputElement>(`[data-hfc="${k}"]`)!;
      sel.innerHTML = hfOptions(presets, printOpts[k]);
      const custom = !presets.some(([c]) => c === printOpts[k]);
      inp.hidden = !custom;
      inp.value = printOpts[k];
      sel.addEventListener('change', () => {
        inp.hidden = sel.value !== '*';
        if (sel.value !== '*') { printOpts = { ...printOpts, [k]: sel.value }; changed(); } else inp.focus();
      });
      inp.addEventListener('change', () => { printOpts = { ...printOpts, [k]: inp.value }; changed(); });
    }
    const paText = body.querySelector<HTMLElement>('.pa-text')!;
    const showArea = () => { paText.textContent = host.printArea() || 'не задана (печатается весь лист)'; };
    showArea();
    body.querySelector('[data-act="pa-set"]')!.addEventListener('click', () => { host.setPrintArea(false); showArea(); refresh(); });
    body.querySelector('[data-act="pa-clear"]')!.addEventListener('click', () => { host.setPrintArea(true); showArea(); refresh(); });
    // Сквозные строки и столбцы: повторяются на каждой странице, хранятся в книге (и в xlsx).
    const tRows = body.querySelector<HTMLInputElement>('[data-t="rows"]')!;
    const tCols = body.querySelector<HTMLInputElement>('[data-t="cols"]')!;
    const tNote = body.querySelector<HTMLElement>('.pt-note')!;
    const showTitles = () => { const t = host.printTitles(); tRows.value = t.rows; tCols.value = t.cols; };
    showTitles();
    for (const inp of [tRows, tCols]) inp.addEventListener('change', () => {
      tNote.textContent = host.setPrintTitles(tRows.value, tCols.value);
      if (!tNote.textContent) refresh();
    });
    const marginBox = body.querySelector<HTMLElement>('.print-margins')!;
    const marginInputs = [...marginBox.querySelectorAll<HTMLInputElement>('[data-m]')];
    const showMargins = () => {
      marginBox.hidden = printOpts.margins !== 'custom';
      marginsOf(printOpts).forEach((v, i) => { marginInputs[i].value = String(Math.round(v)); });
    };
    const refresh = () => { showMargins(); host.printPreview(preview, printOpts); };
    const changed = () => { host.savePrintOptions(printOpts); refresh(); };
    body.querySelectorAll<HTMLInputElement | HTMLSelectElement>('[data-p]').forEach((el) => {
      const k = el.dataset.p as keyof PrintOptions;
      if (el instanceof HTMLInputElement) el.checked = !!printOpts[k];
      else el.value = String(printOpts[k]);
      el.addEventListener('change', () => {
        const v = el instanceof HTMLInputElement ? el.checked : el.value === 'true' ? true : el.value === 'false' ? false : el.value;
        // «Настраиваемые» начинаются с текущих полей.
        if (k === 'margins' && v === 'custom') printOpts = { ...printOpts, marginMm: marginsOf(printOpts) };
        printOpts = { ...printOpts, [k]: v } as PrintOptions;
        changed();
      });
    });
    marginInputs.forEach((inp, i) => inp.addEventListener('change', () => {
      const mm = [...printOpts.marginMm] as PrintOptions['marginMm'];
      mm[i] = Math.max(0, Math.min(100, Number(inp.value.replace(',', '.')) || 0));
      printOpts = { ...printOpts, marginMm: mm };
      changed();
    }));
    body.querySelector('[data-act="print"]')!.addEventListener('click', () => { hideBackstage(); host.print(printOpts); });
    body.querySelector('[data-act="pdf"]')!.addEventListener('click', () => { hideBackstage(); host.print(printOpts); });
    requestAnimationFrame(refresh);
  } else if (section === 'options') {
    body.innerHTML = `<h1>Параметры</h1>
      <h2>Ввод и числа</h2>
      <label class="bs-row">Разделитель дробной части
        <select data-set="decimal"><option value="comma">запятая: 0,5</option><option value="point">точка: 0.5</option></select></label>
      <label class="bs-row"><input type="checkbox" data-set="recognizeDates"> Распознавать даты при вводе в ячейки «Общий»</label>
      <label class="bs-row">Шрифт по умолчанию для новых книг
        <select data-set="defaultFont">${FONTS.map((f) => `<option>${f.label}</option>`).join('')}</select>
        <input type="number" min="6" max="72" data-set="defaultFontSize" style="width:64px"></label>
      <label class="bs-row"><input type="checkbox" data-set="expandEditor"> Расширять поле ввода вправо по мере набора</label>
      <h2>Журнал изменений</h2>
      <label class="bs-row">Ваше имя в журнале изменений книги
        <input data-set="userName" placeholder="как в системе" style="width:220px"></label>
      ${isDesktop ? `
      <h2>Сохранение и восстановление</h2>
      <label class="bs-row">Снимок книги для восстановления каждые
        <input type="number" min="1" max="120" data-set="autosaveMinutes" style="width:64px"> мин.</label>
      <div class="bs-row">Папка временных файлов: <span class="bs-rec">…</span>
        <button data-act="recfolder">Изменить…</button></div>
      <p class="bs-muted">Можно выбрать папку рядом с документами или сетевую папку. Если сетевая папка
      недоступна, временные файлы пишутся в профиль пользователя, чтобы ничего не потерять.</p>` : ''}`;
    const dec = body.querySelector<HTMLSelectElement>('[data-set="decimal"]')!;
    dec.value = settings.decimal;
    const un = body.querySelector<HTMLInputElement>('[data-set="userName"]')!;
    un.value = settings.userName;
    un.addEventListener('change', () => { settings.userName = un.value.trim(); saveSettings(); });
    dec.addEventListener('change', () => { settings.decimal = dec.value as 'comma' | 'point'; saveSettings(); host.applySettings(); });
    const exp = body.querySelector<HTMLInputElement>('[data-set="expandEditor"]')!;
    exp.checked = settings.expandEditor;
    exp.addEventListener('change', () => { settings.expandEditor = exp.checked; saveSettings(); host.applySettings(); });
    const fontSel = body.querySelector<HTMLSelectElement>('[data-set="defaultFont"]')!;
    const fontSize = body.querySelector<HTMLInputElement>('[data-set="defaultFontSize"]')!;
    fontSel.value = settings.defaultFont;
    fontSize.value = String(settings.defaultFontSize);
    const saveFont = () => {
      settings.defaultFont = fontSel.value;
      settings.defaultFontSize = Math.min(72, Math.max(6, Number(fontSize.value) || 12));
      saveSettings(); host.applySettings();
    };
    fontSel.addEventListener('change', saveFont);
    fontSize.addEventListener('change', saveFont);
    const dates = body.querySelector<HTMLInputElement>('[data-set="recognizeDates"]')!;
    dates.checked = settings.recognizeDates;
    dates.addEventListener('change', () => { settings.recognizeDates = dates.checked; saveSettings(); host.applySettings(); });
    if (isDesktop) {
      const mins = body.querySelector<HTMLInputElement>('[data-set="autosaveMinutes"]')!;
      mins.value = String(settings.autosaveMinutes);
      mins.addEventListener('change', () => {
        settings.autosaveMinutes = Math.min(120, Math.max(1, Number(mins.value) || 10));
        mins.value = String(settings.autosaveMinutes);
        saveSettings();
      });
      const show = () => void host.recoveryFolder().then((p) => { body.querySelector('.bs-rec')!.textContent = p; });
      show();
      body.querySelector('[data-act="recfolder"]')!.addEventListener('click', () => void host.chooseRecoveryFolder().then(show));
    }
  }
}

const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);
