//! ENS-Office: drop-down lists (data validation of the "list" kind): which cells have them, their items,
//! and the check of what is typed into such a cell.

use crate::expressions::utils::parse_reference_a1;

use super::{common::UserModel, history::Diff};
use crate::types::Validation;

/// One range of an `sqref`: "B2:B10" or "D4" → (row1, column1, row2, column2).
fn parse_span(token: &str) -> Option<(i32, i32, i32, i32)> {
    let (a, b) = token.split_once(':').unwrap_or((token, token));
    let (a, b) = (parse_reference_a1(a)?, parse_reference_a1(b)?);
    Some((a.row.min(b.row), a.column.min(b.column), a.row.max(b.row), a.column.max(b.column)))
}

fn covers(v: &Validation, row: i32, column: i32) -> bool {
    v.sqref.split_whitespace().filter_map(parse_span).any(|(r1, c1, r2, c2)| row >= r1 && row <= r2 && column >= c1 && column <= c2)
}

/// A span after `delta` lines were inserted (> 0) or deleted (< 0) at `at`; `None` — nothing is left of it.
fn shift_span(lo: i32, hi: i32, at: i32, delta: i32) -> Option<(i32, i32)> {
    if delta > 0 {
        return Some((if lo >= at { lo + delta } else { lo }, if hi >= at { hi + delta } else { hi }));
    }
    let n = -delta;
    if hi < at {
        return Some((lo, hi));
    }
    if lo >= at + n {
        return Some((lo - n, hi - n));
    }
    let new_lo = lo.min(at);
    let new_hi = if hi >= at + n { hi - n } else { at - 1 };
    (new_hi >= new_lo).then_some((new_lo, new_hi))
}

impl UserModel<'_> {
    /// The drop-down lists of the sheet.
    pub fn get_validations(&self, sheet: u32) -> Result<Vec<Validation>, String> {
        Ok(self.model.workbook.worksheet(sheet)?.ens_validations.clone())
    }

    /// Replaces the drop-down lists of the sheet; an undoable step.
    pub fn set_validations(&mut self, sheet: u32, validations: Vec<Validation>) -> Result<(), String> {
        self.guard_sheet(sheet)?;
        for v in &validations {
            if v.source.trim().is_empty() || v.sqref.split_whitespace().next().is_none() || v.sqref.split_whitespace().any(|t| parse_span(t).is_none()) {
                return Err("Неверный список: укажите ячейки и источник.".to_string());
            }
        }
        let old_value = self.get_validations(sheet)?;
        if old_value == validations {
            return Ok(());
        }
        self.model.workbook.worksheet_mut(sheet)?.ens_validations = validations.clone();
        self.push_diff_list(vec![Diff::SetValidations { sheet, new_value: validations, old_value }]);
        Ok(())
    }

    fn validation_at(&self, sheet: u32, row: i32, column: i32) -> Option<Validation> {
        self.model.workbook.worksheet(sheet).ok()?.ens_validations.iter().find(|v| covers(v, row, column)).cloned()
    }

    /// Whether the cell has a drop-down list with an arrow button.
    pub fn has_dropdown(&self, sheet: u32, row: i32, column: i32) -> bool {
        self.validation_at(sheet, row, column).is_some_and(|v| !v.no_arrow)
    }

    fn items_of(&self, sheet: u32, v: &Validation) -> Vec<String> {
        let src = v.source.trim().trim_start_matches('=');
        if let Some(list) = src.strip_prefix('"') {
            return list.trim_end_matches('"').split(',').map(|s| s.trim().to_string()).filter(|s| !s.is_empty()).collect();
        }
        // A defined name stands for a range.
        let named = self.model.get_defined_name_list().into_iter().find(|(n, scope, _)| n.eq_ignore_ascii_case(src) && scope.is_none_or(|s| s == sheet));
        let src = named.as_ref().map_or(src, |(_, _, f)| f.trim_start_matches('='));
        let (sheet_name, cells) = src.rsplit_once('!').map_or((None, src), |(s, c)| (Some(s.trim_matches('\'').replace("''", "'")), c));
        let target = match sheet_name {
            Some(name) => match self.model.workbook.worksheets.iter().position(|w| w.name.eq_ignore_ascii_case(&name)) {
                Some(i) => i as u32,
                None => return vec![],
            },
            None => sheet,
        };
        let Some((r1, c1, r2, c2)) = parse_span(&cells.replace('$', "")) else { return vec![] };
        let mut out = Vec::new();
        for r in r1..=r2.min(r1 + 5000) {
            for c in c1..=c2.min(c1 + 200) {
                if let Ok(t) = self.model.get_formatted_cell_value(target, r, c) {
                    if !t.is_empty() {
                        out.push(t);
                    }
                }
            }
        }
        out
    }

    /// The items of the drop-down list of the cell, if it has one.
    pub fn get_dropdown_items(&self, sheet: u32, row: i32, column: i32) -> Option<Vec<String>> {
        let v = self.validation_at(sheet, row, column)?;
        Some(self.items_of(sheet, &v))
    }

    /// Error if `value` may not be typed into the cell (it has a list and the value is not in it).
    pub(crate) fn check_validation(&self, sheet: u32, row: i32, column: i32, value: &str) -> Result<(), String> {
        let Some(v) = self.validation_at(sheet, row, column) else { return Ok(()) };
        let value = value.trim();
        if value.is_empty() {
            return if v.allow_blank { Ok(()) } else { Err(self.validation_error(&v)) };
        }
        let items = self.items_of(sheet, &v);
        if items.iter().any(|i| i.to_lowercase() == value.to_lowercase()) {
            Ok(())
        } else {
            Err(self.validation_error(&v))
        }
    }

    fn validation_error(&self, v: &Validation) -> String {
        if v.error.is_empty() { "Значение не входит в список допустимых.".to_string() } else { v.error.clone() }
    }

    /// After rows or columns were inserted or deleted: the lists follow their cells. Returns the steps for undo.
    pub(crate) fn shift_validations(&mut self, sheet: u32, is_row: bool, at: i32, delta: i32) -> Result<Vec<Diff>, String> {
        let old_value = self.get_validations(sheet)?;
        let mut new_value = Vec::new();
        for v in &old_value {
            let parts: Vec<String> = v
                .sqref
                .split_whitespace()
                .filter_map(parse_span)
                .filter_map(|(r1, c1, r2, c2)| {
                    let (lo, hi) = if is_row { (r1, r2) } else { (c1, c2) };
                    let (lo, hi) = shift_span(lo, hi, at, delta)?;
                    let (r1, r2, c1, c2) = if is_row { (lo, hi, c1, c2) } else { (r1, r2, lo, hi) };
                    let cell = |r: i32, c: i32| format!("{}{r}", crate::expressions::utils::number_to_column(c).unwrap_or_default());
                    Some(if (r1, c1) == (r2, c2) { cell(r1, c1) } else { format!("{}:{}", cell(r1, c1), cell(r2, c2)) })
                })
                .collect();
            if !parts.is_empty() {
                new_value.push(Validation { sqref: parts.join(" "), ..v.clone() });
            }
        }
        if new_value == old_value {
            return Ok(vec![]);
        }
        self.model.workbook.worksheet_mut(sheet)?.ens_validations = new_value.clone();
        Ok(vec![Diff::SetValidations { sheet, new_value, old_value }])
    }
}
