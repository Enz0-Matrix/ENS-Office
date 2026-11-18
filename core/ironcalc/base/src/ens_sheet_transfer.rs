// SPDX-License-Identifier: AGPL-3.0-or-later
//! ENS-Office: перенос листа в другую книгу («Переместить или скопировать», ADR-012).
//!
//! Лист уезжает «пакетом» (`SheetPackage`), в котором нет ничего, что зависит от книги-источника:
//! стили, общие строки и форматы условного форматирования лежат в пакете списками,
//! ссылки на другие листы книги-источника переписаны во внешние ссылки `'[Книга.xlsx]Лист'!A1`
//! (как делает Excel), а значения этих листов приложены — в новой книге формулы сразу
//! показывают те же числа, а «Обновить связи» подтянет свежие из файла.
//! Явные ссылки листа на самого себя в пакете ведут на служебное имя [`SELF_NAME`]; при вставке
//! лист сначала получает это имя, затем обычное переименование книги переписывает все ссылки.

use std::collections::HashMap;

use bitcode::{Decode, Encode};

use crate::{
    cf_types::CfRule,
    expressions::{
        lexer::LexerMode,
        parser::stringify::{for_each_sheet_ref, to_english_string, to_rc_format},
        types::CellReferenceRC,
    },
    external::{external_sheet_name, is_external_sheet_name, split_external_sheet_name, ExternalValue},
    model::Model,
    sheet_data::SheetData,
    types::{Cell, DefinedName, Dxf, RichRun, SheetState, Style, Table, Worksheet},
};

/// Служебное имя листа внутри пакета (такого имени у пользователя не бывает).
pub const SELF_NAME: &str = "ENS перенос ⁂";

/// Значения листа другой книги, приложенные к пакету.
#[derive(Encode, Decode, Clone, Debug, PartialEq)]
pub struct PackedExternal {
    /// Имя файла книги (`Цены.xlsx`) и путь, как записан (`C:\Папка\Цены.xlsx`; пусто — имя файла).
    pub file: String,
    pub target: String,
    pub sheet: String,
    /// (строка, столбец, вид: 0 — число, 1 — текст, 2 — логическое, 3 — ошибка, число, текст).
    pub values: Vec<(i32, i32, u8, f64, String)>,
}

/// Лист, готовый к вставке в любую книгу.
#[derive(Encode, Decode, Clone)]
pub struct SheetPackage {
    /// Имя листа в книге-источнике.
    pub name: String,
    /// Книга-источник (имя файла, как в формулах `[...]`).
    pub source_file: String,
    /// Лист: стили — номера в `styles`, строки — номера в `strings`, форматы правил — в `dxfs`.
    pub worksheet: Worksheet,
    pub styles: Vec<Style>,
    pub strings: Vec<String>,
    /// Оформленные части текста: номер в `strings` → части.
    pub rich: Vec<(i32, Vec<RichRun>)>,
    pub dxfs: Vec<Dxf>,
    pub tables: Vec<Table>,
    /// Имена, видимые только на этом листе: (имя, формула).
    pub names: Vec<(String, String)>,
    pub external: Vec<PackedExternal>,
}

fn pack_value(v: &ExternalValue) -> (u8, f64, String) {
    match v {
        ExternalValue::Number(n) => (0, *n, String::new()),
        ExternalValue::Text(t) => (1, 0.0, t.clone()),
        ExternalValue::Bool(b) => (2, if *b { 1.0 } else { 0.0 }, String::new()),
        ExternalValue::Error(e) => (3, 0.0, e.clone()),
    }
}

fn unpack_value(kind: u8, n: f64, t: &str) -> ExternalValue {
    match kind {
        0 => ExternalValue::Number(n),
        2 => ExternalValue::Bool(n != 0.0),
        3 => ExternalValue::Error(t.to_string()),
        _ => ExternalValue::Text(t.to_string()),
    }
}

/// Номер формата у правила условного форматирования.
fn map_cf_dxf(rule: &CfRule, map: &mut dyn FnMut(u32) -> u32) -> CfRule {
    let mut rule = rule.clone();
    match &mut rule {
        CfRule::CellIs { dxf_id, .. }
        | CfRule::Formula { dxf_id, .. }
        | CfRule::Text { dxf_id, .. }
        | CfRule::TimePeriod { dxf_id, .. }
        | CfRule::DuplicateValues { dxf_id, .. }
        | CfRule::UniqueValues { dxf_id, .. }
        | CfRule::Blanks { dxf_id, .. }
        | CfRule::NotBlanks { dxf_id, .. }
        | CfRule::Errors { dxf_id, .. }
        | CfRule::NoErrors { dxf_id, .. }
        | CfRule::AboveAverage { dxf_id, .. }
        | CfRule::BelowAverage { dxf_id, .. }
        | CfRule::Top10 { dxf_id, .. }
        | CfRule::Bottom10 { dxf_id, .. } => *dxf_id = map(*dxf_id),
        _ => {}
    }
    rule
}

fn set_cell_style(cell: &mut Cell, map: &mut dyn FnMut(i32) -> i32) {
    let s = cell.get_style();
    cell.set_style(map(s));
}

impl<'a> Model<'a> {
    /// Пакет листа `sheet` для вставки в другую книгу. `source_file` — имя этой книги
    /// для внешних ссылок, `source_target` — её путь на диске (пусто — неизвестен).
    pub fn ens_export_sheet(&mut self, sheet: u32, source_file: &str, source_target: &str) -> Result<SheetPackage, String> {
        let source = self.workbook.worksheet(sheet)?.clone();
        if is_external_sheet_name(&source.name) {
            return Err("Служебный лист другой книги не переносится".to_string());
        }
        let names: Vec<String> = self.workbook.get_worksheet_names();
        // Как называется каждый лист книги в пакете: этот — служебное имя, другие — внешние ссылки.
        let outside: Vec<Option<String>> = names
            .iter()
            .enumerate()
            .map(|(i, n)| {
                if i as u32 == sheet {
                    Some(SELF_NAME.to_string())
                } else if is_external_sheet_name(n) {
                    Some(n.clone())
                } else {
                    Some(external_sheet_name(source_file, n))
                }
            })
            .collect();
        let mut used = vec![false; names.len()];

        // Формулы ячеек (R1C1, разбираются в контексте этого листа).
        self.parser.set_lexer_mode(LexerMode::R1C1);
        let context = CellReferenceRC { sheet: source.name.clone(), row: 1, column: 1 };
        let mut formulas = Vec::with_capacity(source.shared_formulas.len());
        for formula in &source.shared_formulas {
            let mut node = self.parser.parse(formula, &context);
            for_each_sheet_ref(&mut node, &mut |name, index| {
                let i = index as usize;
                if name.is_some() && i < outside.len() {
                    used[i] = true;
                    *name = outside[i].clone();
                }
            });
            formulas.push(to_rc_format(&node));
        }
        self.parser.set_lexer_mode(LexerMode::A1);

        // Имена, видимые только на этом листе.
        let name_context = self.defined_name_context();
        let mut local_names = Vec::new();
        for dn in self.workbook.defined_names.clone() {
            if dn.sheet_id != Some(source.sheet_id) {
                continue;
            }
            let had_equals = dn.formula.trim_start().starts_with('=');
            let body = dn.formula.trim_start().strip_prefix('=').unwrap_or(&dn.formula).to_string();
            let mut node = self.parse_internal_formula(&body, &name_context);
            for_each_sheet_ref(&mut node, &mut |name, index| {
                let i = index as usize;
                if name.is_some() && i < outside.len() {
                    used[i] = true;
                    *name = outside[i].clone();
                }
            });
            let text = to_english_string(&node, &name_context);
            local_names.push((dn.name.clone(), if had_equals { format!("={text}") } else { text }));
        }

        // Стили, строки, оформленный текст.
        let mut styles: Vec<Style> = Vec::new();
        let mut style_map: HashMap<i32, i32> = HashMap::new();
        let mut style_of = |index: i32, model: &Model| -> i32 {
            if let Some(i) = style_map.get(&index) {
                return *i;
            }
            let Ok(style) = model.workbook.styles.get_style(index).or_else(|_| model.workbook.styles.get_style(0)) else {
                return 0;
            };
            let i = styles.len() as i32;
            styles.push(style);
            style_map.insert(index, i);
            i
        };
        let mut strings: Vec<String> = Vec::new();
        let mut string_map: HashMap<i32, i32> = HashMap::new();
        let mut rich = Vec::new();
        let mut data = SheetData::new();
        for (row, column, cell) in source.sheet_data.cells() {
            let mut cell = cell.into_owned();
            set_cell_style(&mut cell, &mut |s| style_of(s, self));
            if let Cell::SharedString { si, .. } = &mut cell {
                let old = *si;
                let new = *string_map.entry(old).or_insert_with(|| {
                    strings.push(self.workbook.shared_strings.get(old as usize).cloned().unwrap_or_default());
                    (strings.len() - 1) as i32
                });
                if let Some(runs) = self.workbook.ens_rich.get(&old) {
                    rich.push((new, runs.clone()));
                }
                *si = new;
            }
            data.set_cell(row, column, cell);
        }
        let mut ws = source.clone();
        ws.sheet_data = data;
        ws.shared_formulas = formulas;
        for r in ws.rows.iter_mut() {
            r.s = style_of(r.s, self);
        }
        for c in ws.cols.iter_mut() {
            if let Some(s) = c.style {
                c.style = Some(style_of(s, self));
            }
        }

        // Форматы условного форматирования и таблиц.
        let mut dxfs: Vec<Dxf> = Vec::new();
        let mut dxf_map: HashMap<u32, u32> = HashMap::new();
        let all_dxfs = self.workbook.styles.dxfs.clone();
        let mut dxf_of = |id: u32| -> u32 {
            *dxf_map.entry(id).or_insert_with(|| {
                dxfs.push(all_dxfs.get(id as usize).cloned().unwrap_or_default());
                (dxfs.len() - 1) as u32
            })
        };
        for cf in ws.conditional_formatting.iter_mut() {
            cf.cf_rule = map_cf_dxf(&cf.cf_rule, &mut dxf_of);
        }
        let mut tables: Vec<Table> = Vec::new();
        for t in self.workbook.tables.values() {
            if t.sheet_name != source.name {
                continue;
            }
            let mut t = t.clone();
            t.header_row_dxf_id = t.header_row_dxf_id.map(&mut dxf_of);
            t.data_dxf_id = t.data_dxf_id.map(&mut dxf_of);
            t.totals_row_dxf_id = t.totals_row_dxf_id.map(&mut dxf_of);
            tables.push(t);
        }
        tables.sort_by(|a, b| a.name.cmp(&b.name));

        // Значения листов, на которые ссылается лист.
        let mut external = Vec::new();
        for (i, name) in names.iter().enumerate() {
            if !used[i] || i as u32 == sheet {
                continue;
            }
            let values = self
                .sheet_values(i as u32)
                .into_iter()
                .map(|(r, c, v)| {
                    let (k, n, t) = pack_value(&v);
                    (r, c, k, n, t)
                })
                .collect();
            let (file, target, sheet_name) = match split_external_sheet_name(name) {
                Some((file, s)) => {
                    let target = self
                        .workbook
                        .ens_external_books
                        .iter()
                        .find(|b| b.file.eq_ignore_ascii_case(file))
                        .map(|b| b.target.clone())
                        .unwrap_or_default();
                    (file.to_string(), target, s.to_string())
                }
                None => (source_file.to_string(), source_target.to_string(), name.clone()),
            };
            external.push(PackedExternal { file, target, sheet: sheet_name, values });
        }

        Ok(SheetPackage {
            name: source.name.clone(),
            source_file: source_file.to_string(),
            worksheet: ws,
            styles,
            strings,
            rich,
            dxfs,
            tables,
            names: local_names,
            external,
        })
    }

    /// Вставляет лист из пакета перед номером `index` (не дальше служебных листов) под именем
    /// `name`. Возвращает номер нового листа.
    pub fn ens_import_sheet(&mut self, pkg: &SheetPackage, index: u32, name: &str) -> Result<u32, String> {
        if self.get_sheet_index_by_name(SELF_NAME).is_some() {
            return Err("Лист с служебным именем уже есть в книге".to_string());
        }
        if self.get_sheet_index_by_name(name).is_some() {
            return Err(format!("Лист «{name}» уже есть в книге"));
        }
        let index = index.min(self.first_external_sheet());

        // Значения других книг (и книги-источника), на которые ссылается лист.
        for e in &pkg.external {
            // Ссылка обратно на эту же книгу: лист вернулся домой — значения не нужны,
            // но служебный лист всё равно нужен формулам; обычные листы не трогаем.
            self.ensure_external_book(&e.file, &e.target);
            let values: Vec<(i32, i32, ExternalValue)> =
                e.values.iter().map(|(r, c, k, n, t)| (*r, *c, unpack_value(*k, *n, t))).collect();
            self.set_external_values(&e.file, &e.sheet, &values)?;
        }

        // Стили, строки, форматы правил — в номера этой книги.
        let styles: Vec<i32> = pkg.styles.iter().map(|s| self.workbook.styles.get_style_index_or_create(s)).collect();
        let style_of = |s: i32| styles.get(s as usize).copied().unwrap_or(0);
        let dxf_base = self.workbook.styles.dxfs.len() as u32;
        self.workbook.styles.dxfs.extend(pkg.dxfs.iter().cloned());
        let rich: HashMap<i32, &Vec<RichRun>> = pkg.rich.iter().map(|(i, r)| (*i, r)).collect();
        let mut string_of: HashMap<i32, i32> = HashMap::new();
        let mut ws = pkg.worksheet.clone();
        let mut data = SheetData::new();
        for (row, column, cell) in pkg.worksheet.sheet_data.cells() {
            let mut cell = cell.into_owned();
            cell.set_style(style_of(cell.get_style()));
            if let Cell::SharedString { si, .. } = &mut cell {
                let text = pkg.strings.get(*si as usize).cloned().unwrap_or_default();
                let new = if let Some(runs) = rich.get(si) {
                    // оформленная строка принадлежит одной ячейке: всегда своя
                    let i = self.workbook.shared_strings.len() as i32;
                    self.workbook.shared_strings.push(text);
                    self.workbook.ens_rich.insert(i, (*runs).clone());
                    i
                } else if let Some(i) = string_of.get(si) {
                    *i
                } else {
                    let i = match self.shared_strings.get(&text) {
                        Some(i) => *i as i32,
                        None => {
                            let i = self.workbook.shared_strings.len();
                            self.workbook.shared_strings.push(text.clone());
                            self.shared_strings.insert(text, i);
                            i as i32
                        }
                    };
                    string_of.insert(*si, i);
                    i
                };
                *si = new;
            }
            data.set_cell(row, column, cell);
        }
        ws.sheet_data = data;
        for r in ws.rows.iter_mut() {
            r.s = style_of(r.s);
        }
        for c in ws.cols.iter_mut() {
            c.style = c.style.map(style_of);
        }
        for cf in ws.conditional_formatting.iter_mut() {
            cf.cf_rule = map_cf_dxf(&cf.cf_rule, &mut |id| id + dxf_base);
        }
        let sheet_id = self.get_new_sheet_id();
        let view_ids: Vec<&u32> = self.workbook.views.keys().collect();
        ws.views = Model::new_empty_worksheet(SELF_NAME, sheet_id, &view_ids).views;
        ws.name = SELF_NAME.to_string();
        ws.sheet_id = sheet_id;
        ws.state = SheetState::Visible;
        ws.ens_code_name = String::new();
        self.workbook.worksheets.insert(index as usize, ws);

        // Таблицы: имя таблицы уникально во всей книге.
        for t in &pkg.tables {
            let mut t = t.clone();
            let mut n = 1;
            let base = t.name.clone();
            while self.workbook.tables.keys().any(|k| k.eq_ignore_ascii_case(&t.name)) {
                n += 1;
                t.name = format!("{base}_{n}");
                t.display_name = t.name.clone();
                t.ens_raw_xml = String::new();
            }
            t.sheet_name = SELF_NAME.to_string();
            t.header_row_dxf_id = t.header_row_dxf_id.map(|i| i + dxf_base);
            t.data_dxf_id = t.data_dxf_id.map(|i| i + dxf_base);
            t.totals_row_dxf_id = t.totals_row_dxf_id.map(|i| i + dxf_base);
            self.workbook.tables.insert(t.name.clone(), t);
        }
        for (n, formula) in &pkg.names {
            self.workbook.defined_names.push(DefinedName { name: n.clone(), formula: formula.clone(), sheet_id: Some(sheet_id) });
        }
        self.reset_parsed_structures();
        // Обычное переименование переписывает ссылки листа на себя в формулах и именах.
        self.rename_sheet_by_index(index, name)?;
        // Таблицы переименование листа не трогает.
        for t in self.workbook.tables.values_mut() {
            if t.sheet_name == SELF_NAME {
                t.sheet_name = name.to_string();
            }
        }
        self.reset_parsed_structures();
        Ok(index)
    }
}

/// Пакет в байты и обратно (для передачи между окнами программы).
pub fn sheet_package_to_bytes(pkg: &SheetPackage) -> Vec<u8> {
    bitcode::encode(pkg)
}

pub fn sheet_package_from_bytes(bytes: &[u8]) -> Result<SheetPackage, String> {
    bitcode::decode(bytes).map_err(|e| format!("Повреждённые данные листа: {e}"))
}

