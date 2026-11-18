// SPDX-License-Identifier: AGPL-3.0-or-later
//! ENS-Office: «умные» таблицы (ADR-013). Имена столбцов таблицы — это текст ячеек её
//! заголовка, как в Excel. Изменили заголовок — столбец переименовывается, и формулы со
//! структурированными ссылками (`Таблица1[Цена]`) переписываются на новое имя.
//!
//! Сверка идёт после каждого действия пользователя, отмены, повтора и шага журнала,
//! поэтому неважно, каким путём изменилась ячейка (ввод, вставка, очистка, отмена).

use crate::{
    expressions::{
        lexer::LexerMode,
        parser::{
            stringify::{to_english_string, to_rc_format, visit_structured_mut},
            Node,
        },
        token::TableReference,
        types::CellReferenceRC,
        utils::parse_reference_a1,
    },
    language::get_default_language,
    locale::get_default_locale,
    model::Model,
    types::Table,
};

/// Границы таблицы `A1:F5` → (строка1, столбец1, строка2, столбец2).
pub fn table_bounds(reference: &str) -> Option<(i32, i32, i32, i32)> {
    let mut parts = reference.split(':');
    let a = parse_reference_a1(parts.next()?)?;
    let b = match parts.next() {
        Some(p) => parse_reference_a1(p)?,
        None => a.clone(),
    };
    Some((a.row, a.column, b.row, b.column))
}

impl<'a> Model<'a> {
    /// Имена столбцов по ячейкам заголовка: пустая ячейка — прежнее имя, повтор — с номером
    /// («Цена2»), без учёта регистра. Так же их пишет xlsx.
    pub fn ens_header_names(&self, sheet: u32, table: &Table) -> Vec<String> {
        let header = table_bounds(&table.reference).filter(|_| table.header_row_count > 0);
        let mut names: Vec<String> = Vec::with_capacity(table.columns.len());
        for (i, col) in table.columns.iter().enumerate() {
            let from_cell = header
                .and_then(|(r1, c1, _, _)| self.get_formatted_cell_value(sheet, r1, c1 + i as i32).ok())
                .filter(|s| !s.trim().is_empty());
            let mut name = from_cell.unwrap_or_else(|| col.name.clone());
            let base = name.clone();
            let mut n = 2;
            while names.iter().any(|x| x.to_lowercase() == name.to_lowercase()) {
                name = format!("{base}{n}");
                n += 1;
            }
            names.push(name);
        }
        names
    }

    /// Сверяет имена столбцов всех таблиц с заголовками; переименованные — в формулах тоже.
    /// `true` — что-то переименовано (формулы разобраны заново).
    pub fn ens_sync_table_headers(&mut self) -> bool {
        let mut renames: Vec<(String, Vec<(String, String)>)> = Vec::new();
        let mut keys: Vec<String> = self.workbook.tables.keys().cloned().collect();
        keys.sort();
        for key in keys {
            let table = &self.workbook.tables[&key];
            let Some(sheet) = self.get_sheet_index_by_name(&table.sheet_name) else {
                continue;
            };
            let names = self.ens_header_names(sheet, table);
            let changed: Vec<(String, String)> = table
                .columns
                .iter()
                .zip(names.iter())
                .filter(|(c, n)| c.name != **n)
                .map(|(c, n)| (c.name.clone(), n.clone()))
                .collect();
            if changed.is_empty() {
                continue;
            }
            if let Some(t) = self.workbook.tables.get_mut(&key) {
                for (col, name) in t.columns.iter_mut().zip(names) {
                    col.name = name;
                }
            }
            renames.push((key, changed));
        }
        if renames.is_empty() {
            return false;
        }
        self.ens_rename_in_formulas(&renames);
        self.reset_parsed_structures();
        true
    }

    /// Переписывает имена столбцов в формулах ячеек и в именах книги.
    fn ens_rename_in_formulas(&mut self, renames: &[(String, Vec<(String, String)>)]) {
        let rename = |node: &mut Node| {
            let mut changed = false;
            visit_structured_mut(node, &mut |table_name, reference| {
                let Some((_, list)) = renames.iter().find(|(t, _)| t.eq_ignore_ascii_case(table_name)) else {
                    return;
                };
                let fix = |c: &mut String, changed: &mut bool| {
                    // одновременная перестановка имён: «старое» берём по исходному списку
                    if let Some((_, new)) = list.iter().find(|(old, _)| old.to_lowercase() == c.to_lowercase()) {
                        *c = new.clone();
                        *changed = true;
                    }
                };
                match reference {
                    Some(TableReference::ColumnReference(c)) => fix(c, &mut changed),
                    Some(TableReference::RangeReference((a, b))) => {
                        fix(a, &mut changed);
                        fix(b, &mut changed);
                    }
                    None => {}
                }
            });
            changed
        };
        let names: Vec<String> = renames.iter().map(|(t, _)| t.to_lowercase()).collect();
        let mentions = |f: &str| {
            let lower = f.to_lowercase();
            names.iter().any(|n| lower.contains(n.as_str()))
        };
        // Разбор — по старым именам столбцов: парсер ещё знает прежние таблицы.
        // Формулы ячеек хранятся по-английски в R1C1.
        let (locale, language) = (self.locale, self.language);
        self.parser.set_locale(get_default_locale());
        self.parser.set_language(get_default_language());
        self.parser.set_lexer_mode(LexerMode::R1C1);
        for index in 0..self.workbook.worksheets.len() {
            let context = CellReferenceRC {
                sheet: self.workbook.worksheets[index].get_name(),
                row: 1,
                column: 1,
            };
            let formulas = self.workbook.worksheets[index].shared_formulas.clone();
            let mut out = Vec::with_capacity(formulas.len());
            for f in formulas {
                if !mentions(&f) {
                    out.push(f);
                    continue;
                }
                let mut node = self.parser.parse(&f, &context);
                if rename(&mut node) {
                    out.push(to_rc_format(&node));
                } else {
                    out.push(f);
                }
            }
            self.workbook.worksheets[index].shared_formulas = out;
        }
        self.parser.set_lexer_mode(LexerMode::A1);
        self.parser.set_locale(locale);
        self.parser.set_language(language);
        let context = self.defined_name_context();
        let mut defined = self.workbook.defined_names.clone();
        for dn in defined.iter_mut() {
            if !mentions(&dn.formula) {
                continue;
            }
            let had_equals = dn.formula.trim_start().starts_with('=');
            let body = dn.formula.trim_start().trim_start_matches('=').to_string();
            let mut node = self.parse_internal_formula(&body, &context);
            if rename(&mut node) {
                let text = to_english_string(&node, &context);
                dn.formula = if had_equals { format!("={text}") } else { text };
            }
        }
        self.workbook.defined_names = defined;
    }
}
