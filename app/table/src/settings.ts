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
  /** Закреплённые в «Файл → Открыть» файлы и папки. */
  pinned: { path: string; title: string; kind: 'file' | 'folder' }[];
  /** Что отменяют стрелки: вся книга, только этот лист или все листы, кроме этого. */
  undoScope: 'book' | 'sheet' | 'others';
  /** Шрифт по умолчанию для новых книг. */
  defaultFont: string;
  defaultFontSize: number;
  /** Имя для журнала изменений книги ('' — имя пользователя в системе). */
  userName: string;
  /** Свои шаблоны диаграмм: цвета, легенда, подписи. */
  chartTemplates: { name: string; colors: string[]; legend: number; labels: boolean }[];
  /** Когда вы последний раз открывали или сохраняли файл (путь → мс): что изменилось с тех пор. */
  lastOpened: Record<string, number>;
  /** Помощник «У. Болт» (выключен по умолчанию). */
  helper: boolean;
  /** Пасхалка с конфетти уже показана. */
  bigBookCelebrated: boolean;
}

const KEY = 'ens-table.settings';
const defaults: Settings = {
  decimal: 'comma', undoScope: 'book', defaultFont: 'Times New Roman', defaultFontSize: 12, expandEditor: false, autosaveMinutes: 10, recognizeDates: true, recoveryRoot: '', recent: [], pinned: [], userName: '', chartTemplates: [], lastOpened: {}, helper: false, bigBookCelebrated: false, sortBlanks: 'last', borderType: 'All', borderStyle: 'thin', theme: 'classic', themeMotion: true, themeMotionForced: false, hiddenPolicy: 'ask', customThemes: [], hiddenThemes: [], formulaHeight: 26, formulaExpanded: 80,
};

export const settings: Settings = (() => {
  try {
    return { ...defaults, ...JSON.parse(localStorage.getItem(KEY) ?? '{}') };
  } catch {
    return { ...defaults };
  }
})();

/** Списки, общие для всех окон программы: файл, открытый в другом окне, сразу виден в «Последних». */
const SHARED = ['recent', 'pinned'] as const;

function readStored(): Partial<Settings> {
  try { return JSON.parse(localStorage.getItem(KEY) ?? '{}'); } catch { return {}; }
}

/** Подтянуть общие списки, которые изменили другие окна. */
export function refreshShared(): void {
  const stored = readStored();
  for (const k of SHARED) if (Array.isArray(stored[k])) (settings as unknown as Record<string, unknown>)[k] = stored[k];
}

export function saveSettings(): void {
  try {
    // Общие списки берём из хранилища: чужие окна могли их изменить, а у этого окна они устарели.
    const stored = readStored();
    const out: Record<string, unknown> = { ...settings };
    for (const k of SHARED) if (Array.isArray(stored[k])) out[k] = stored[k];
    localStorage.setItem(KEY, JSON.stringify(out));
    refreshShared();
  } catch {
    /* хранилище недоступно — настройки действуют до закрытия */
  }
}

/** Изменить общий список (последние, закреплённые) и сразу записать, не затирая правки других окон. */
export function updateShared<K extends (typeof SHARED)[number]>(key: K, change: (list: Settings[K]) => Settings[K]): void {
  refreshShared();
  settings[key] = change(settings[key]);
  try {
    const out = { ...readStored(), [key]: settings[key] };
    localStorage.setItem(KEY, JSON.stringify({ ...settings, ...out }));
  } catch { /* хранилище недоступно */ }
}

export function localeId(): string {
  return settings.decimal === 'comma' ? 'ru' : 'en';
}
