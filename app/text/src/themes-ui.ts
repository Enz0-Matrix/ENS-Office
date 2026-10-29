// Темы оформления ENS-Tekst — те же, что в ENS-Table (CSS и список общие, настройки общие): меняют вид программы
// и цвет листа на экране; печать и файл остаются как есть.

import '../../table/src/themes.css';
import '../../table/src/themes2.css';
import {
  THEMES, allThemes, applyTheme, parseThemeFile, setCustomThemes, setHiddenThemes, systemReducesMotion, themeById, themeTemplate,
} from '../../table/src/themes';
import { copyText } from '../../table/src/platform';
import { saveSettings, settings } from '../../table/src/settings';
import { dialog, el, toast } from './ui';

const BUILTIN = new Set(THEMES.map((t) => t.id));

/** Включить тему из настроек: атрибуты для CSS, цвет листа. */
export function setTheme(): void {
  applyTheme(settings.theme, { theme: {} as never, render: () => {} }, settings.themeMotion, settings.themeMotionForced);
  // Лист на экране — цвет «листа» темы (светлый, чтобы читался чёрный текст из файлов); печать всегда белая.
  document.documentElement.style.setProperty('--paper', themeById(settings.theme).grid.sheetBg ?? '#ffffff');
}

export function initThemes(): void {
  setHiddenThemes(settings.hiddenThemes);
  setCustomThemes(settings.customThemes);
  setTheme();
}

function loadThemeFile(done: () => void): void {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.json,.enstheme,application/json';
  input.addEventListener('change', async () => {
    const file = input.files?.[0];
    if (!file) return;
    const t = parseThemeFile(await file.text());
    if (typeof t === 'string') { toast(t); return; }
    settings.customThemes = [...settings.customThemes.filter((x) => x.id !== t.id), t];
    settings.theme = t.id;
    saveSettings();
    setCustomThemes(settings.customThemes);
    setTheme();
    toast(`Тема «${t.name}» добавлена и включена.`);
    done();
  });
  input.click();
}

/** «Вид → Тема»: выбор темы оформления. */
export function themeDialog(): void {
  dialog((box, close) => {
    const reduce = systemReducesMotion();
    const moving = settings.themeMotion && (!reduce || settings.themeMotionForced);
    const reopen = () => { close(); themeDialog(); };
    const list = el('div', { class: 'theme-list' }, ...allThemes().map((t) => el('button', {
      class: `theme-row${t.id === settings.theme ? ' on' : ''}`,
      onclick: () => { settings.theme = t.id; saveSettings(); setTheme(); reopen(); },
    }, el('b', { textContent: `${t.id === settings.theme ? '✓ ' : ''}${t.name}` }), el('span', { textContent: ` — ${t.note}` }))));
    const extra = el('div', { class: 'theme-extra' },
      el('button', {
        textContent: moving ? 'Выключить анимацию темы' : reduce ? 'Включить анимацию темы (в системе она уменьшена)' : 'Включить анимацию темы',
        onclick: () => { settings.themeMotion = !moving; settings.themeMotionForced = !moving && reduce; saveSettings(); setTheme(); reopen(); },
      }),
      el('button', { textContent: 'Загрузить тему из файла…', onclick: () => loadThemeFile(reopen) }),
      el('button', {
        textContent: 'Сохранить образец темы (в буфер обмена)',
        onclick: () => { void copyText(themeTemplate(settings.theme)).then(() => toast('Образец темы скопирован. Вставьте в Блокнот, поменяйте цвета и сохраните как .json.')); },
      }),
      el('button', {
        textContent: 'Скрыть выбранную тему из списка', disabled: settings.theme === 'classic' || !BUILTIN.has(settings.theme),
        onclick: () => { settings.hiddenThemes = [...settings.hiddenThemes, settings.theme]; setHiddenThemes(settings.hiddenThemes); settings.theme = 'classic'; saveSettings(); setTheme(); reopen(); },
      }),
      ...(settings.hiddenThemes.length ? [el('button', { textContent: `Показать скрытые темы (${settings.hiddenThemes.length})`, onclick: () => { settings.hiddenThemes = []; setHiddenThemes([]); saveSettings(); reopen(); } })] : []),
      ...settings.customThemes.map((t) => el('button', {
        class: 'danger', textContent: `Удалить свою тему «${t.name}»`,
        onclick: () => { settings.customThemes = settings.customThemes.filter((x) => x.id !== t.id); if (settings.theme === t.id) settings.theme = 'classic'; saveSettings(); setCustomThemes(settings.customThemes); setTheme(); reopen(); },
      })),
    );
    box.append(el('h2', { textContent: 'Тема оформления' }), el('p', { textContent: 'Меняет вид программы и цвет листа на экране. Файл и печать остаются как есть.' }), list, extra,
      el('div', { class: 'dlg-actions' }, el('button', { class: 'primary', textContent: 'Готово', onclick: close })));
  }, false);
}
