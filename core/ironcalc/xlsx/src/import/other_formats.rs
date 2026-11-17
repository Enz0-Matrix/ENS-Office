// SPDX-License-Identifier: AGPL-3.0-or-later
//! ENS-Office: «всеядность» — книги не в xlsx: старый `.xls` (Excel 97–2003), `.ods` (OpenDocument,
//! LibreOffice, Р7) и `.csv`. Переносятся листы, значения, формулы и даты; оформление — нет
//! (об этом предупреждение, а исходный файл не перезаписывается: книга открывается как новая).

use std::io::Cursor;

use calamine::{open_workbook_auto_from_rs, Data, Reader, Sheets};
use ironcalc_base::expressions::utils::number_to_column;
use ironcalc_base::Model;

use crate::error::XlsxError;

/// Что не переносится из `.xls` и `.ods` (для предупреждения при открытии).
const OTHER_FORMAT_LOSSES: &[&str] = &[
    "оформление ячеек (шрифты, цвета, границы, форматы чисел, кроме дат)",
    "ширина столбцов и высота строк",
    "объединённые ячейки",
    "примечания, рисунки и диаграммы",
];

fn excel_date_format(value: f64) -> &'static str {
    if value.fract().abs() < 1e-9 {
        "dd.mm.yyyy"
    } else if value < 1.0 {
        "hh:mm:ss"
    } else {
        "dd.mm.yyyy hh:mm"
    }
}

/// Дата и время → число Excel (система 1900).
fn serial_of(dt: chrono::NaiveDateTime) -> f64 {
    let base = chrono::NaiveDate::from_ymd_opt(1899, 12, 30)
        .unwrap_or_default()
        .and_hms_opt(0, 0, 0)
        .unwrap_or_default();
    (dt - base).num_milliseconds() as f64 / 86_400_000.0
}

/// ISO-дата `2026-11-17` или `2026-11-17T08:30:00` (ODS) → число Excel.
fn iso_to_serial(s: &str) -> Option<f64> {
    let (date, time) = match s.split_once('T') {
        Some((d, t)) => (d, Some(t)),
        None => (s, None),
    };
    let mut it = date.split('-');
    let y: i32 = it.next()?.parse().ok()?;
    let m: u32 = it.next()?.parse().ok()?;
    let d: u32 = it.next()?.parse().ok()?;
    let mut v = ironcalc_base::formatter::dates::date_to_serial_number(d, m, y).ok()? as f64;
    if let Some(t) = time {
        let parts: Vec<f64> = t
            .trim_end_matches('Z')
            .split(':')
            .filter_map(|x| x.parse().ok())
            .collect();
        let secs = parts.first().unwrap_or(&0.0) * 3600.0
            + parts.get(1).unwrap_or(&0.0) * 60.0
            + parts.get(2).unwrap_or(&0.0);
        v += secs / 86400.0;
    }
    Some(v)
}

/// ISO-длительность `PT12H30M00S` (ODS, время) → доля суток.
fn iso_duration(s: &str) -> Option<f64> {
    let t = s.strip_prefix("PT").or_else(|| s.strip_prefix("-PT"))?;
    let (mut h, mut m, mut sec, mut num) = (0.0, 0.0, 0.0, String::new());
    for c in t.chars() {
        match c {
            'H' => h = num.parse().ok()?,
            'M' => m = num.parse().ok()?,
            'S' => sec = num.parse().ok()?,
            _ => {
                num.push(c);
                continue;
            }
        }
        num.clear();
    }
    let v = (h * 3600.0 + m * 60.0 + sec) / 86400.0;
    Some(if s.starts_with('-') { -v } else { v })
}

/// Формула OpenFormula (ODS: `of:=SUM([.A1:.B2];[Лист2.C3])`) → формула Excel (`SUM(A1:B2,Лист2!C3)`).
pub(crate) fn openformula_to_excel(f: &str) -> String {
    let f = f.trim();
    let f = f.strip_prefix("of:").unwrap_or(f);
    let f = f.strip_prefix('=').unwrap_or(f);
    let chars: Vec<char> = f.chars().collect();
    let mut out = String::with_capacity(f.len());
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '"' {
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
        if c == '[' {
            if let Some(end) = chars[i..].iter().position(|&x| x == ']').map(|p| i + p) {
                let inner: String = chars[i + 1..end].iter().collect();
                out.push_str(&odf_reference(&inner));
                i = end + 1;
                continue;
            }
        }
        out.push(match c {
            ';' => ',',
            '~' => ',', // объединение диапазонов в OpenFormula
            _ => c,
        });
        i += 1;
    }
    out
}

/// `.A1`, `.A1:.B2`, `$Лист2.$C$3`, `'Мой лист'.A1:.A5` → ссылка Excel.
fn odf_reference(inner: &str) -> String {
    inner
        .split(':')
        .enumerate()
        .map(|(k, part)| {
            let (sheet, cell) = match part.rfind('.') {
                Some(p) => (&part[..p], &part[p + 1..]),
                None => ("", part),
            };
            let sheet = sheet.trim_start_matches('$');
            if sheet.is_empty() || k > 0 {
                cell.to_string()
            } else {
                let name = sheet.trim_matches('\'').replace("''", "'");
                format!(
                    "{}!{cell}",
                    ironcalc_base::expressions::utils::quote_name(&name)
                )
            }
        })
        .collect::<Vec<_>>()
        .join(":")
}

/// `.xls` (calamine): имена листов в формулах приходят без кавычек (`Второй лист!A1`).
/// Ставит кавычки там, где они нужны, вне строк в кавычках.
fn quote_sheet_names(formula: &str, sheets: &[String]) -> String {
    let mut names: Vec<&String> = sheets
        .iter()
        .filter(|n| ironcalc_base::expressions::utils::quote_name(n).starts_with('\''))
        .collect();
    if names.is_empty() || !formula.contains('!') {
        return formula.to_string();
    }
    names.sort_by_key(|n| std::cmp::Reverse(n.chars().count()));
    let mut out = String::with_capacity(formula.len() + 8);
    let mut rest = formula;
    let mut in_string = false;
    'outer: while let Some(c) = rest.chars().next() {
        if c == '"' {
            in_string = !in_string;
        } else if !in_string && !out.ends_with('\'') {
            for n in &names {
                let pat = format!("{n}!");
                let boundary = out
                    .chars()
                    .last()
                    .is_none_or(|p| !(p.is_alphanumeric() || p == '_' || p == '.'));
                if boundary && rest.starts_with(&pat) {
                    out.push_str(&ironcalc_base::expressions::utils::quote_name(n));
                    out.push('!');
                    rest = &rest[pat.len()..];
                    continue 'outer;
                }
            }
        }
        out.push(c);
        rest = &rest[c.len_utf8()..];
    }
    out
}

/// Открывает `.xls` или `.ods`: листы, значения, формулы, даты.
/// Книга создаётся по-английски (формулы xls и ods записаны по-английски), затем переводится
/// в язык и локаль пользователя.
pub fn load_from_other_bytes<'a>(
    bytes: &[u8],
    name: &str,
    locale: &str,
    tz: &str,
    language: &'a str,
) -> Result<Model<'a>, XlsxError> {
    let mut book = open_workbook_auto_from_rs(Cursor::new(bytes.to_vec()))
        .map_err(|e| XlsxError::Workbook(format!("Не удалось прочитать книгу: {e}")))?;
    let is_ods = matches!(book, Sheets::Ods(_));
    let mut m = Model::new_empty(name, "en", tz, "en").map_err(XlsxError::Workbook)?;
    let names = book.sheet_names().to_vec();
    if names.is_empty() {
        return Err(XlsxError::Workbook("В книге нет листов".to_string()));
    }
    for (i, sheet_name) in names.iter().enumerate() {
        if i == 0 {
            m.rename_sheet_by_index(0, sheet_name)
                .map_err(XlsxError::Workbook)?;
        } else {
            m.add_sheet(sheet_name).map_err(XlsxError::Workbook)?;
        }
    }
    let mut date_cells: Vec<(u32, i32, i32, &'static str)> = Vec::new();
    for (i, sheet_name) in names.iter().enumerate() {
        let sheet = i as u32;
        let range = book
            .worksheet_range(sheet_name)
            .map_err(|e| XlsxError::Workbook(format!("Лист «{sheet_name}»: {e}")))?;
        if let Some((r0, c0)) = range.start() {
            for (r, c, v) in range.cells() {
                let row = (r0 as usize + r) as i32 + 1;
                let col = (c0 as usize + c) as i32 + 1;
                let res = match v {
                    Data::Empty => Ok(()),
                    Data::Int(n) => m.update_cell_with_number(sheet, row, col, *n as f64),
                    Data::Float(f) => m.update_cell_with_number(sheet, row, col, *f),
                    Data::Bool(b) => m.update_cell_with_bool(sheet, row, col, *b),
                    Data::String(s) => m.update_cell_with_text(sheet, row, col, s),
                    Data::DateTime(dt) => {
                        let v = if dt.is_duration() {
                            date_cells.push((sheet, row, col, "[h]:mm:ss"));
                            dt.as_f64()
                        } else {
                            // as_datetime учитывает систему дат книги (1900 или 1904)
                            let v = dt
                                .as_datetime()
                                .map(serial_of)
                                .unwrap_or_else(|| dt.as_f64());
                            date_cells.push((sheet, row, col, excel_date_format(v)));
                            v
                        };
                        m.update_cell_with_number(sheet, row, col, v)
                    }
                    Data::DateTimeIso(s) => match iso_to_serial(s) {
                        Some(v) => {
                            date_cells.push((sheet, row, col, excel_date_format(v)));
                            m.update_cell_with_number(sheet, row, col, v)
                        }
                        None => m.update_cell_with_text(sheet, row, col, s),
                    },
                    Data::DurationIso(s) => match iso_duration(s) {
                        Some(v) => {
                            date_cells.push((sheet, row, col, "[h]:mm:ss"));
                            m.update_cell_with_number(sheet, row, col, v)
                        }
                        None => m.update_cell_with_text(sheet, row, col, s),
                    },
                    Data::Error(e) => m.set_user_input(sheet, row, col, e.to_string()),
                };
                res.map_err(XlsxError::Workbook)?;
            }
        }
        // формулы — поверх значений (значения остаются как кэш до пересчёта)
        if let Ok(formulas) = book.worksheet_formula(sheet_name) {
            if let Some((r0, c0)) = formulas.start() {
                for (r, c, f) in formulas.cells() {
                    if f.trim().is_empty() {
                        continue;
                    }
                    let row = (r0 as usize + r) as i32 + 1;
                    let col = (c0 as usize + c) as i32 + 1;
                    let text = if is_ods {
                        openformula_to_excel(f)
                    } else {
                        quote_sheet_names(f.trim_start_matches('='), &names)
                    };
                    // формула, которую не удалось разобрать, остаётся текстом, а не пропадает
                    if m.set_user_input(sheet, row, col, format!("={text}"))
                        .is_err()
                    {
                        let _ = m.update_cell_with_text(sheet, row, col, &format!("={text}"));
                    }
                }
            }
        }
    }
    for (sheet, row, col, fmt) in date_cells {
        let style = m
            .get_style_for_cell(sheet, row, col)
            .map_err(XlsxError::Workbook)?;
        let mut style = style;
        style.num_fmt = fmt.to_string();
        m.set_cell_style(sheet, row, col, &style)
            .map_err(XlsxError::Workbook)?;
    }
    let mut workbook = m.workbook;
    workbook.name = name.to_string();
    workbook.settings.locale = locale.to_string();
    workbook.ens_lost_features = OTHER_FORMAT_LOSSES.iter().map(|s| s.to_string()).collect();
    if !is_ods {
        // calamine не читает общие формулы BIFF (протянутые вниз): такие ячейки приходят значениями
        workbook
            .ens_lost_features
            .push("часть формул .xls (протянутые в соседние ячейки) приходит значениями — проверьте расчёты".to_string());
    }
    let mut model = Model::from_workbook(workbook, language).map_err(XlsxError::Workbook)?;
    model.evaluate();
    Ok(model)
}

/// Кодировка CSV: UTF-8 (с BOM или без), иначе Windows-1251 — так сохраняют Excel и 1С.
fn decode_csv(bytes: &[u8]) -> String {
    if let Some(rest) = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]) {
        return String::from_utf8_lossy(rest).into_owned();
    }
    if bytes.starts_with(&[0xFF, 0xFE]) {
        return encoding_rs::UTF_16LE.decode(bytes).0.into_owned();
    }
    match std::str::from_utf8(bytes) {
        Ok(s) => s.to_string(),
        Err(_) => encoding_rs::WINDOWS_1251.decode(bytes).0.into_owned(),
    }
}

/// Разделитель: тот из `;`, `,`, табуляции, что встречается в первых строках чаще и ровнее.
fn detect_separator(text: &str, default: char) -> char {
    let lines: Vec<&str> = text
        .lines()
        .filter(|l| !l.trim().is_empty())
        .take(20)
        .collect();
    let mut best = (default, 0usize);
    for sep in [';', '\t', ','] {
        let counts: Vec<usize> = lines
            .iter()
            .map(|l| split_csv_line(l, sep).len() - 1)
            .collect();
        let first = counts.first().copied().unwrap_or(0);
        if first == 0 {
            continue;
        }
        let same = counts.iter().filter(|&&n| n == first).count();
        let score = same * 100 + first;
        if score > best.1 {
            best = (sep, score);
        }
    }
    best.0
}

/// Поля одной строки (кавычки внутри строки не переносятся на следующую — для подсчёта).
fn split_csv_line(line: &str, sep: char) -> Vec<String> {
    parse_csv(line, sep).into_iter().next().unwrap_or_default()
}

/// Разбор CSV по RFC 4180: поля в кавычках, `""` — кавычка, переводы строк внутри кавычек.
pub(crate) fn parse_csv(text: &str, sep: char) -> Vec<Vec<String>> {
    let mut rows = Vec::new();
    let mut row = Vec::new();
    let mut field = String::new();
    let mut quoted = false;
    let mut chars = text.chars().peekable();
    let mut any = false;
    while let Some(c) = chars.next() {
        any = true;
        if quoted {
            if c == '"' {
                if chars.peek() == Some(&'"') {
                    field.push('"');
                    chars.next();
                } else {
                    quoted = false;
                }
            } else {
                field.push(c);
            }
        } else if c == '"' && field.is_empty() {
            quoted = true;
        } else if c == sep {
            row.push(std::mem::take(&mut field));
        } else if c == '\n' || c == '\r' {
            if c == '\r' && chars.peek() == Some(&'\n') {
                chars.next();
            }
            row.push(std::mem::take(&mut field));
            rows.push(std::mem::take(&mut row));
            any = false;
        } else {
            field.push(c);
        }
    }
    if any {
        row.push(field);
        rows.push(row);
    }
    rows
}

/// Открывает `.csv`: значения читаются, как при вводе с клавиатуры в локали пользователя
/// («1 234,56» — число, «17.11.2026» — дата), как это делает Excel.
pub fn load_from_csv_bytes<'a>(
    bytes: &[u8],
    name: &str,
    locale: &str,
    tz: &str,
    language: &'a str,
) -> Result<Model<'a>, XlsxError> {
    let text = decode_csv(bytes);
    let mut m = Model::new_empty(name, locale, tz, language).map_err(XlsxError::Workbook)?;
    let comma_decimal =
        ironcalc_base::locale::get_locale(locale).is_ok_and(|l| l.numbers.symbols.decimal == ",");
    let default_sep = if comma_decimal { ';' } else { ',' };
    let sep = detect_separator(&text, default_sep);
    let sheet_name: String = name
        .chars()
        .filter(|c| !"\\/*?:[]".contains(*c))
        .take(31)
        .collect();
    if !sheet_name.trim().is_empty() {
        let _ = m.rename_sheet_by_index(0, sheet_name.trim());
    }
    for (r, fields) in parse_csv(&text, sep).iter().enumerate() {
        for (c, value) in fields.iter().enumerate() {
            if value.is_empty() {
                continue;
            }
            let (row, col) = (r as i32 + 1, c as i32 + 1);
            if number_to_column(col).is_none() {
                break;
            }
            if m.set_user_input(0, row, col, value.clone()).is_err() {
                m.update_cell_with_text(0, row, col, value)
                    .map_err(XlsxError::Workbook)?;
            }
        }
    }
    let mut workbook = m.workbook;
    // CSV открывается как новая книга: исходный файл не перезаписывается, терять нечего
    let mut model = Model::from_workbook(workbook, language).map_err(XlsxError::Workbook)?;
    model.evaluate();
    Ok(model)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn openformula() {
        assert_eq!(
            openformula_to_excel("of:=SUM([.A1:.B2];[$Лист2.C3])"),
            "SUM(A1:B2,Лист2!C3)"
        );
        assert_eq!(
            openformula_to_excel("of:=[.A1]&\"[.x]; y\""),
            "A1&\"[.x]; y\""
        );
        assert_eq!(
            openformula_to_excel("of:=['Мой лист'.$A$1]*2"),
            "'Мой лист'!$A$1*2"
        );
    }

    #[test]
    fn xls_sheet_names() {
        let sheets = vec!["Второй лист".to_string(), "Лист1".to_string()];
        assert_eq!(
            quote_sheet_names("Второй лист!A1*2+Лист1!B2", &sheets),
            "'Второй лист'!A1*2+Лист1!B2"
        );
        assert_eq!(
            quote_sheet_names("\"Второй лист!A1\"&'Второй лист'!A1", &sheets),
            "\"Второй лист!A1\"&'Второй лист'!A1"
        );
    }

    #[test]
    fn csv() {
        assert_eq!(
            parse_csv("a;\"b;c\";\"d\"\"e\"\r\n1;2\n", ';'),
            vec![vec!["a", "b;c", "d\"e"], vec!["1", "2"]]
        );
        assert_eq!(
            parse_csv("\"много\nстрок\",x", ','),
            vec![vec!["много\nстрок", "x"]]
        );
        assert_eq!(detect_separator("a;b;c\n1;2;3\n", ','), ';');
        assert_eq!(detect_separator("a,b\n1,2\n", ';'), ',');
        assert_eq!(
            decode_csv(&encoding_rs::WINDOWS_1251.encode("Цена").0),
            "Цена"
        );
        assert_eq!(iso_to_serial("2026-11-17"), Some(46343.0));
        assert_eq!(iso_duration("PT12H30M00S"), Some(0.5208333333333334));
    }
}
