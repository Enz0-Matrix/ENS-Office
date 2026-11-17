// SPDX-License-Identifier: AGPL-3.0-or-later
//! ENS-Office: ссылки на другие книги (внешние связи), в том числе на закрытые.
//!
//! Каждый лист книги-источника живёт в модели как скрытый служебный лист с именем
//! `[Источник.xlsx]Цены` (в Excel у настоящего листа таких скобок быть не может).
//! Поэтому ВПР, ИНДЕКС, СУММ и любые диапазоны по нему работают как по обычному листу.
//! Служебные листы всегда стоят после настоящих, не показываются и не сохраняются
//! как листы: в xlsx они уходят в `xl/externalLinks` (кэш значений), а ссылки
//! записываются в виде Excel `[1]Цены!A1`.
//!
//! Значения берутся из кэша файла или обновляются из файла-источника, который
//! программа читает с диска сама — запускать его не нужно.

use crate::cell::CellValue;
use crate::expressions::token::get_error_by_english_name;
use crate::model::Model;
use crate::types::{ExternalBook, SheetState};

/// Ссылка на лист другой книги, найденная в тексте формулы.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalRef {
    /// Папка перед `[`, как её написали (`C:\Отчёты\`), или пусто.
    pub path: String,
    /// То, что в квадратных скобках: имя файла или номер связи Excel (`1`).
    pub book: String,
    pub sheet: String,
}

/// Имя служебного листа для листа другой книги.
pub fn external_sheet_name(file: &str, sheet: &str) -> String {
    format!("[{file}]{sheet}")
}

/// `[Источник.xlsx]Цены` → (`Источник.xlsx`, `Цены`).
pub fn split_external_sheet_name(name: &str) -> Option<(&str, &str)> {
    let rest = name.strip_prefix('[')?;
    let end = rest.find(']')?;
    Some((&rest[..end], &rest[end + 1..]))
}

/// Служебный лист другой книги.
pub fn is_external_sheet_name(name: &str) -> bool {
    split_external_sheet_name(name).is_some()
}

/// Имя файла из пути: `C:\Папка\Книга.xlsx`, `file:///D:/Папка/Книга.xlsx` → `Книга.xlsx`.
pub fn file_name_of(target: &str) -> &str {
    target.rsplit(['/', '\\']).next().unwrap_or(target)
}

fn sheet_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_' || c == '.'
}

/// Находит ссылки на листы других книг вне строк в кавычках и заменяет их.
/// `replace` получает ссылку и возвращает текст на место `…` в `…!A1` (с кавычками,
/// если нужно) или `None` — оставить как было.
///
/// Понимает `[1]Цены!`, `'[1]Мои цены'!`, `[Книга.xlsx]Цены!`, `'C:\Папка\[Книга.xlsx]Цены'!`.
/// Структурные ссылки (`Таблица1[Сумма]`, `[@Цена]`) не трогает.
pub fn rewrite_external_refs(
    formula: &str,
    replace: &mut dyn FnMut(&ExternalRef) -> Option<String>,
) -> String {
    if !formula.contains('[') {
        return formula.to_string();
    }
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len() + 8);
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '"' {
            // строка: до закрывающей кавычки ("" — кавычка внутри)
            out.push(c);
            i += 1;
            while i < chars.len() {
                out.push(chars[i]);
                if chars[i] == '"' {
                    if chars.get(i + 1) == Some(&'"') {
                        out.push('"');
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            continue;
        }
        if c == '\'' {
            // имя листа в кавычках: '…'!  ('' — апостроф внутри)
            let mut j = i + 1;
            let mut name = String::new();
            let mut closed = false;
            while j < chars.len() {
                if chars[j] == '\'' {
                    if chars.get(j + 1) == Some(&'\'') {
                        name.push('\'');
                        j += 2;
                        continue;
                    }
                    closed = true;
                    break;
                }
                name.push(chars[j]);
                j += 1;
            }
            if closed && chars.get(j + 1) == Some(&'!') {
                if let Some(r) = parse_quoted(&name) {
                    if let Some(rep) = replace(&r) {
                        out.push_str(&rep);
                        i = j + 1; // '!' остаётся
                        continue;
                    }
                }
            }
            // не наше — переносим как есть
            let end = if closed { j + 1 } else { chars.len() };
            out.extend(&chars[i..end]);
            i = end;
            continue;
        }
        if c == '[' {
            let prev_ident = i > 0
                && (chars[i - 1].is_alphanumeric() || chars[i - 1] == '_' || chars[i - 1] == ']');
            if !prev_ident {
                if let Some(close) = chars[i + 1..]
                    .iter()
                    .position(|&x| x == ']')
                    .map(|p| i + 1 + p)
                {
                    let book: String = chars[i + 1..close].iter().collect();
                    let mut k = close + 1;
                    while k < chars.len() && sheet_char(chars[k]) {
                        k += 1;
                    }
                    let sheet: String = chars[close + 1..k].iter().collect();
                    let good_book = !book.is_empty()
                        && !book.starts_with('@')
                        && !book.contains(['[', '\'', '"']);
                    if good_book && !sheet.is_empty() && chars.get(k) == Some(&'!') {
                        let r = ExternalRef {
                            path: String::new(),
                            book,
                            sheet,
                        };
                        if let Some(rep) = replace(&r) {
                            out.push_str(&rep);
                            i = k;
                            continue;
                        }
                    }
                }
            }
        }
        out.push(c);
        i += 1;
    }
    out
}

/// Содержимое кавычек `C:\Папка\[Книга.xlsx]Лист` → ссылка; без `[…]` — не ссылка на книгу.
fn parse_quoted(name: &str) -> Option<ExternalRef> {
    let open = name.find('[')?;
    let close = open + name[open..].find(']')?;
    let book = &name[open + 1..close];
    let sheet = &name[close + 1..];
    if book.is_empty() || sheet.is_empty() {
        return None;
    }
    Some(ExternalRef {
        path: name[..open].to_string(),
        book: book.to_string(),
        sheet: sheet.to_string(),
    })
}

/// Текст ссылки на служебный лист: `'[Книга.xlsx]Лист'`.
pub fn quoted_internal(file: &str, sheet: &str) -> String {
    format!("'{}'", external_sheet_name(file, sheet).replace('\'', "''"))
}

/// Формула в виде, понятном разборщику: ссылки `[1]Лист!` (номер связи Excel, `books` —
/// имена файлов по порядку) и `[Книга.xlsx]Лист!`, `'C:\…\[Книга.xlsx]Лист'!` → `'[Книга.xlsx]Лист'!`.
pub fn to_internal_refs(formula: &str, books: &[String]) -> String {
    rewrite_external_refs(formula, &mut |r| {
        let file = match r.book.parse::<usize>() {
            Ok(n) if r.path.is_empty() => books.get(n.checked_sub(1)?)?.clone(),
            _ => r.book.clone(),
        };
        Some(quoted_internal(&file, &r.sheet))
    })
}

/// Формула для xlsx: `'[Книга.xlsx]Лист'!` → `[n]Лист!` (n — номер книги среди `books`, с 1).
pub fn to_excel_refs(formula: &str, books: &[String]) -> String {
    rewrite_external_refs(formula, &mut |r| {
        if !r.path.is_empty() {
            return None;
        }
        let n = books.iter().position(|b| {
            b.eq_ignore_ascii_case(&r.book) || b.to_lowercase() == r.book.to_lowercase()
        })? + 1;
        let inner = format!("[{n}]{}", r.sheet);
        let simple = r.sheet.chars().all(sheet_char)
            && !r
                .sheet
                .starts_with(|c: char| c.is_ascii_digit() || c == '.');
        Some(if simple {
            inner
        } else {
            format!("'{}'", inner.replace('\'', "''"))
        })
    })
}

/// Пустой скрытый служебный лист для листа другой книги (для импорта xlsx).
pub fn new_external_worksheet(file: &str, sheet: &str, sheet_id: u32) -> crate::types::Worksheet {
    let mut ws = Model::new_empty_worksheet(&external_sheet_name(file, sheet), sheet_id, &[&0]);
    ws.state = SheetState::VeryHidden;
    ws
}

/// Значение ячейки листа другой книги (кэш).
#[derive(Debug, Clone, PartialEq)]
pub enum ExternalValue {
    Number(f64),
    Text(String),
    Bool(bool),
    /// Ошибка по-английски: `#N/A`.
    Error(String),
}

impl<'a> Model<'a> {
    /// Номер первого служебного листа (или число листов, если их нет).
    pub fn first_external_sheet(&self) -> u32 {
        self.workbook
            .worksheets
            .iter()
            .position(|w| is_external_sheet_name(&w.name))
            .unwrap_or(self.workbook.worksheets.len()) as u32
    }

    /// Книга-источник по имени файла; добавляет её, если нет. `target` — путь, как его
    /// написали в формуле или как он записан в xlsx; пусто — только имя файла.
    pub fn ensure_external_book(&mut self, file: &str, target: &str) -> usize {
        let books = &mut self.workbook.ens_external_books;
        if let Some(i) = books
            .iter()
            .position(|b| b.file.to_lowercase() == file.to_lowercase())
        {
            if !target.is_empty() && books[i].target != target {
                books[i].target = target.to_string();
            }
            return i;
        }
        books.push(ExternalBook {
            target: if target.is_empty() {
                file.to_string()
            } else {
                target.to_string()
            },
            file: file.to_string(),
        });
        books.len() - 1
    }

    /// Служебный лист для листа другой книги; создаёт его, если нет (скрыт, в конце).
    /// Возвращает номер листа и признак «создан» (тогда формулы надо разобрать заново).
    pub fn ensure_external_sheet(&mut self, file: &str, sheet: &str) -> (u32, bool) {
        let name = external_sheet_name(file, sheet);
        let lower = name.to_lowercase();
        if let Some(i) = self
            .workbook
            .worksheets
            .iter()
            .position(|w| w.name.to_lowercase() == lower)
        {
            return (i as u32, false);
        }
        let sheet_id = self.get_new_sheet_id();
        let view_ids: Vec<&u32> = self.workbook.views.keys().collect();
        let mut ws = Model::new_empty_worksheet(&name, sheet_id, &view_ids);
        ws.state = SheetState::VeryHidden;
        self.workbook.worksheets.push(ws);
        ((self.workbook.worksheets.len() - 1) as u32, true)
    }

    /// Перед разбором введённой формулы: книги и листы, на которые она ссылается,
    /// появляются в модели. `true` — что-то добавлено.
    pub(crate) fn register_external_refs(&mut self, formula: &str) -> bool {
        if !formula.starts_with('=') || !formula.contains('[') {
            return false;
        }
        let mut found: Vec<ExternalRef> = Vec::new();
        rewrite_external_refs(formula, &mut |r| {
            found.push(r.clone());
            None
        });
        let books: Vec<String> = self
            .workbook
            .ens_external_books
            .iter()
            .map(|b| b.file.clone())
            .collect();
        let mut added = false;
        for r in found {
            let file = match r.book.parse::<usize>() {
                Ok(n) if r.path.is_empty() => match n.checked_sub(1).and_then(|k| books.get(k)) {
                    Some(f) => f.clone(),
                    None => continue,
                },
                _ => r.book.clone(),
            };
            let target = if r.path.is_empty() {
                String::new()
            } else {
                format!("{}{}", r.path, file)
            };
            let before = self.workbook.ens_external_books.len();
            self.ensure_external_book(&file, &target);
            added |= self.workbook.ens_external_books.len() != before;
            added |= self.ensure_external_sheet(&file, &r.sheet).1;
        }
        if added {
            self.reset_parsed_structures();
        }
        added
    }

    /// Записывает значения листа другой книги (кэш из xlsx или свежие из файла).
    pub fn set_external_values(
        &mut self,
        file: &str,
        sheet: &str,
        values: &[(i32, i32, ExternalValue)],
    ) -> Result<(), String> {
        let (index, _) = self.ensure_external_sheet(file, sheet);
        let ws = self.workbook.worksheet_mut(index)?;
        ws.sheet_data = Default::default();
        for (row, column, value) in values {
            match value {
                ExternalValue::Number(n) => self
                    .workbook
                    .worksheet_mut(index)?
                    .set_cell_with_number(*row, *column, *n, 0)?,
                ExternalValue::Bool(b) => self
                    .workbook
                    .worksheet_mut(index)?
                    .set_cell_with_boolean(*row, *column, *b, 0)?,
                ExternalValue::Text(t) => self.set_cell_with_string(index, *row, *column, t, 0)?,
                ExternalValue::Error(e) => {
                    let error = get_error_by_english_name(e)
                        .unwrap_or(crate::expressions::token::Error::NA);
                    self.workbook
                        .worksheet_mut(index)?
                        .set_cell_with_error(*row, *column, error, 0)?
                }
            }
        }
        Ok(())
    }

    /// Значения листа этой книги (для переноса в другую как источник).
    pub fn sheet_values(&self, sheet: u32) -> Vec<(i32, i32, ExternalValue)> {
        let Ok(ws) = self.workbook.worksheet(sheet) else {
            return Vec::new();
        };
        let mut out = Vec::new();
        for row in ws.sheet_data.rows() {
            for column in ws.sheet_data.columns_in_row(row) {
                let v = match self.get_cell_value_by_index(sheet, row, column) {
                    Ok(CellValue::Number(n)) => ExternalValue::Number(n),
                    Ok(CellValue::Boolean(b)) => ExternalValue::Bool(b),
                    Ok(CellValue::String(s)) => {
                        if s.starts_with('#') && get_error_by_english_name(&s).is_some() {
                            ExternalValue::Error(s)
                        } else {
                            ExternalValue::Text(s)
                        }
                    }
                    _ => continue,
                };
                out.push((row, column, v));
            }
        }
        out
    }

    /// Листы другой книги, которые есть в этой модели (имена в источнике).
    pub fn external_sheets_of(&self, file: &str) -> Vec<String> {
        let lower = file.to_lowercase();
        self.workbook
            .worksheets
            .iter()
            .filter_map(|w| split_external_sheet_name(&w.name))
            .filter(|(f, _)| f.to_lowercase() == lower)
            .map(|(_, s)| s.to_string())
            .collect()
    }

    /// Обновляет значения всех листов книги `file` из модели книги-источника.
    /// Возвращает листы, которых в источнике нет (ссылки на них дадут #ССЫЛКА!, как в Excel).
    pub fn update_external_book_from(&mut self, file: &str, source: &Model) -> Vec<String> {
        let mut missing = Vec::new();
        for sheet in self.external_sheets_of(file) {
            let lower = sheet.to_lowercase();
            let found =
                source.workbook.worksheets.iter().position(|w| {
                    w.name.to_lowercase() == lower && !is_external_sheet_name(&w.name)
                });
            match found {
                Some(i) => {
                    let values = source.sheet_values(i as u32);
                    let _ = self.set_external_values(file, &sheet, &values);
                }
                None => missing.push(sheet),
            }
        }
        missing
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rewrite_forms() {
        let books = vec!["Источник.xlsx".to_string()];
        assert_eq!(
            to_internal_refs("=[1]Цены!A1+1", &books),
            "='[Источник.xlsx]Цены'!A1+1"
        );
        assert_eq!(
            to_internal_refs("='[1]Мои цены'!A1", &books),
            "='[Источник.xlsx]Мои цены'!A1"
        );
        assert_eq!(
            to_internal_refs("=[Книга2.xlsx]Лист1!B2", &books),
            "='[Книга2.xlsx]Лист1'!B2"
        );
        assert_eq!(
            to_internal_refs(r"='C:\Отчёты\[Книга2.xlsx]Лист 1'!B2", &books),
            "='[Книга2.xlsx]Лист 1'!B2"
        );
        // строки, структурные ссылки и обычные листы не трогаем
        assert_eq!(
            to_internal_refs("=\"[1]Цены!A1\"&Таблица1[Сумма]", &books),
            "=\"[1]Цены!A1\"&Таблица1[Сумма]"
        );
        assert_eq!(
            to_internal_refs("='Лист 2'!A1+[@Цена]", &books),
            "='Лист 2'!A1+[@Цена]"
        );
        // и обратно, для xlsx
        let books = vec!["Источник.xlsx".to_string(), "Книга2.xlsx".to_string()];
        assert_eq!(
            to_excel_refs("='[Источник.xlsx]Цены'!A1", &books),
            "=[1]Цены!A1"
        );
        assert_eq!(
            to_excel_refs("=SUM('[Книга2.xlsx]Лист 1'!B2:B5)", &books),
            "=SUM('[2]Лист 1'!B2:B5)"
        );
        assert_eq!(file_name_of("file:///D:/Работа/Цены.xlsx"), "Цены.xlsx");
        assert_eq!(file_name_of(r"C:\Отчёты\Цены.xlsx"), "Цены.xlsx");
    }
}
