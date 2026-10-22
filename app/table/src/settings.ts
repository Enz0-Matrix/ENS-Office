// Настройки пользователя. Хранятся в браузере (в оболочке Tauri — в профиле).

export interface Settings {
  /** Разделитель целой и дробной части: запятая (русская локаль) или точка. */
  decimal: 'comma' | 'point';
  /** Поле ввода расширяется вправо по мере набора, закрывая соседние ячейки. */
  expandEditor: boolean;
  /** Как часто сохранять снимок книги для восстановления после сбоя, минут. */
  autosaveMinutes: number;
  /** Распознавать даты при вводе в ячейки «Общий» («22.08» → дата). */
  recognizeDates: boolean;
  /** Сортировка: пустые ячейки в конце (как Excel), в начале или как ноль. */
  sortBlanks: 'last' | 'first' | 'zero';
  /** Кнопка «Границы»: что ставит по щелчку и какой линией. */
  borderType: string;
  borderStyle: string;
  /** Действие задевает скрытые строки/столбцы: спрашивать, как в Excel, всегда со скрытыми, только видимые. */
  hiddenPolicy: 'ask' | 'excel' | 'all' | 'visible';
  /** Свои темы, загруженные из файлов (themes.ts), и скрытые из списка встроенные. */
  customThemes: { id: string; name: string; note: string; tone?: 'light' | 'dark'; vars: Record<string, string>; grid: Record<string, string>; css?: string }[];
  hiddenThemes: string[];
  /** Высота строки формул (пиксели) и последняя развёрнутая высота — для кнопки ⌄. */
  formulaHeight: number;
  formulaExpanded: number;
  /** Тема оформления (themes.ts) и анимация в ней (мерцание и т. п.). */
  theme: string;
  themeMotion: boolean;
  /** Анимацию включили вручную, хотя в системе стоит «уменьшить движение». */
  themeMotionForced: boolean;
  /** Своя папка временных файлов восстановления ('' — в профиле пользователя). */
  recoveryRoot: string;
  /** Последние открытые книги (для «Файл → Открыть»). */
  recent: { path: string; title: string }[];
}

const KEY = 'ens-table.settings';
const defaults: Settings = {
  decimal: 'comma', expandEditor: false, autosaveMinutes: 10, recognizeDates: true, recoveryRoot: '', recent: [], sortBlanks: 'last', borderType: 'All', borderStyle: 'thin', theme: 'classic', themeMotion: true, themeMotionForced: false, hiddenPolicy: 'ask', customThemes: [], hiddenThemes: [], formulaHeight: 26, formulaExpanded: 80,
};

export const settings: Settings = (() => {
  try {
    return { ...defaults, ...JSON.parse(localStorage.getItem(KEY) ?? '{}') };
  } catch {
    return { ...defaults };
  }
})();

export function saveSettings(): void {
  try {
    localStorage.setItem(KEY, JSON.stringify(settings));
  } catch {
    /* хранилище недоступно — настройки действуют до закрытия */
  }
}

export function localeId(): string {
  return settings.decimal === 'comma' ? 'ru' : 'en';
}
