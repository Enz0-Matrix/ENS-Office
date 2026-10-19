// Настройки пользователя. Хранятся в браузере (в оболочке Tauri — в профиле).

export interface Settings {
  /** Разделитель целой и дробной части: запятая (русская локаль) или точка. */
  decimal: 'comma' | 'point';
  /** Поле ввода расширяется вправо по мере набора, закрывая соседние ячейки. */
  expandEditor: boolean;
}

const KEY = 'ens-table.settings';
const defaults: Settings = { decimal: 'comma', expandEditor: false };

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
