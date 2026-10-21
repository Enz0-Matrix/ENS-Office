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
  /** Своя папка временных файлов восстановления ('' — в профиле пользователя). */
  recoveryRoot: string;
  /** Последние открытые книги (для «Файл → Открыть»). */
  recent: { path: string; title: string }[];
}

const KEY = 'ens-table.settings';
const defaults: Settings = {
  decimal: 'comma', expandEditor: false, autosaveMinutes: 10, recognizeDates: true, recoveryRoot: '', recent: [], sortBlanks: 'last',
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
