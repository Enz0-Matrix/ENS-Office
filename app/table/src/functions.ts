// Справочник функций для поиска (§6). Полный список и русские имена в движке —
// задание 6 в docs/01-tasks-for-assistants.md.

export interface FunctionInfo {
  ru: string;
  en: string;
  description: string;
}

export const FUNCTIONS: FunctionInfo[] = [
  { ru: 'СУММ', en: 'SUM', description: 'сумма чисел' },
  { ru: 'СРЗНАЧ', en: 'AVERAGE', description: 'среднее арифметическое' },
  { ru: 'СЧЁТ', en: 'COUNT', description: 'количество чисел' },
  { ru: 'СЧЁТЗ', en: 'COUNTA', description: 'количество непустых ячеек' },
  { ru: 'МАКС', en: 'MAX', description: 'наибольшее значение' },
  { ru: 'МИН', en: 'MIN', description: 'наименьшее значение' },
  { ru: 'ЕСЛИ', en: 'IF', description: 'условие: если … то … иначе' },
  { ru: 'ЕСЛИОШИБКА', en: 'IFERROR', description: 'значение при ошибке' },
  { ru: 'ВПР', en: 'VLOOKUP', description: 'поиск значения в первом столбце таблицы' },
  { ru: 'ГПР', en: 'HLOOKUP', description: 'поиск значения в первой строке таблицы' },
  { ru: 'ПРОСМОТРX', en: 'XLOOKUP', description: 'поиск значения в диапазоне' },
  { ru: 'ИНДЕКС', en: 'INDEX', description: 'значение по номеру строки и столбца' },
  { ru: 'ПОИСКПОЗ', en: 'MATCH', description: 'позиция значения в диапазоне' },
  { ru: 'СУММЕСЛИ', en: 'SUMIF', description: 'сумма по условию' },
  { ru: 'СУММЕСЛИМН', en: 'SUMIFS', description: 'сумма по нескольким условиям' },
  { ru: 'СЧЁТЕСЛИ', en: 'COUNTIF', description: 'количество по условию' },
  { ru: 'СЧЁТЕСЛИМН', en: 'COUNTIFS', description: 'количество по нескольким условиям' },
  { ru: 'ОКРУГЛ', en: 'ROUND', description: 'округление до указанного числа знаков' },
  { ru: 'ОКРУГЛВВЕРХ', en: 'ROUNDUP', description: 'округление вверх' },
  { ru: 'ОКРУГЛВНИЗ', en: 'ROUNDDOWN', description: 'округление вниз' },
  { ru: 'ПРОМЕЖУТОЧНЫЕ.ИТОГИ', en: 'SUBTOTAL', description: 'итог по видимым ячейкам' },
  { ru: 'СЦЕП', en: 'CONCAT', description: 'объединить текст' },
  { ru: 'ОБЪЕДИНИТЬ', en: 'TEXTJOIN', description: 'объединить текст с разделителем' },
  { ru: 'ЛЕВСИМВ', en: 'LEFT', description: 'первые символы текста' },
  { ru: 'ПРАВСИМВ', en: 'RIGHT', description: 'последние символы текста' },
  { ru: 'ПСТР', en: 'MID', description: 'символы из середины текста' },
  { ru: 'ДЛСТР', en: 'LEN', description: 'длина текста' },
  { ru: 'СЖПРОБЕЛЫ', en: 'TRIM', description: 'убрать лишние пробелы' },
  { ru: 'ТЕКСТ', en: 'TEXT', description: 'число как текст в заданном формате' },
  { ru: 'СЕГОДНЯ', en: 'TODAY', description: 'текущая дата' },
  { ru: 'ТДАТА', en: 'NOW', description: 'текущие дата и время' },
  { ru: 'ДАТА', en: 'DATE', description: 'дата из года, месяца и дня' },
  { ru: 'КОНМЕСЯЦА', en: 'EOMONTH', description: 'последний день месяца' },
  { ru: 'УНИК', en: 'UNIQUE', description: 'уникальные значения' },
  { ru: 'ФИЛЬТР', en: 'FILTER', description: 'отбор строк по условию' },
  { ru: 'СОРТ', en: 'SORT', description: 'сортировка диапазона' },
];
