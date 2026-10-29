//! ENS-Office: drop-down lists (data validation of the "list" kind): which cells have them, their items,
//! and the check of what is typed into such a cell.

use crate::expressions::utils::parse_reference_a1;

use super::{common::UserModel, history::Diff};
use crate::formatter::format::parse_date;
use crate::types::Validation;
use serde::Serialize;

/// A message for the user: the hint of a cell, or the warning about a value that breaks the rule.
#[derive(Serialize, Debug, PartialEq, Clone)]
pub struct ValidationMessage {
    pub style: u8,
    pub title: String,
    pub message: String,
}

fn parse_number(t: &str) -> Option<f64> {
    let t = t.trim().replace(',', ".");
    t.parse::<f64>().ok().filter(|n| n.is_finite())
}

/// "13:30" or "13:30:15" → a fraction of a day.
fn parse_time(t: &str) -> Option<f64> {
    let mut parts = t.trim().split(':');
    let h: f64 = parts.next()?.trim().parse().ok()?;
    let m: f64 = parts.next()?.trim().parse().ok()?;
    let sec: f64 = parts.next().map_or(Some(0.0), |p| p.trim().parse().ok())?;
    if parts.next().is_some() || !(0.0..24.0).contains(&h) || !(0.0..60.0).contains(&m) || !(0.0..60.0).contains(&sec) {
        return None;
    }
    Some((h * 3600.0 + m * 60.0 + sec) / 86400.0)
}

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
        self.validation_at(sheet, row, column).is_some_and(|v| matches!(v.kind.as_str(), "" | "list") && !v.no_arrow)
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
        let v = self.validation_at(sheet, row, column).filter(|v| matches!(v.kind.as_str(), "" | "list"))?;
        Some(self.items_of(sheet, &v))
    }

    /// The check applied to `value` typed into the cell; `Some` — the value breaks the rule.
    fn violated(&self, sheet: u32, row: i32, column: i32, value: &str) -> Option<Validation> {
        let v = self.validation_at(sheet, row, column)?;
        let value = value.trim();
        if value.is_empty() {
            return (!v.allow_blank).then_some(v);
        }
        let ok = match v.kind.as_str() {
            "" | "list" => {
                let low = value.to_lowercase();
                self.items_of(sheet, &v).iter().any(|i| i.to_lowercase() == low)
            }
            "custom" => true,
            _ => self.value_fits(sheet, &v, value),
        };
        (!ok).then_some(v)
    }

    /// A bound of a rule: a number, a date, a time, or a reference to a cell.
    fn bound(&self, sheet: u32, kind: &str, text: &str) -> Option<f64> {
        let t = text.trim().trim_start_matches('=');
        if let Some(n) = parse_number(t) {
            return Some(n);
        }
        if kind == "date" {
            if let Ok((serial, _)) = parse_date(t, self.model.locale) {
                return Some(serial as f64);
            }
        }
        if kind == "time" {
            if let Some(n) = parse_time(t) {
                return Some(n);
            }
        }
        let r = parse_reference_a1(&t.replace('$', ""))?;
        match self.model.get_cell_value_by_index(sheet, r.row, r.column).ok()? {
            crate::cell::CellValue::Number(n) => Some(n),
            _ => None,
        }
    }

    fn value_fits(&self, sheet: u32, v: &Validation, value: &str) -> bool {
        if value.starts_with('=') {
            return true; // a formula is not checked
        }
        let kind = v.kind.as_str();
        let x = match kind {
            "textLength" => value.chars().count() as f64,
            "whole" | "decimal" => match parse_number(value) {
                Some(n) if kind == "decimal" || n.fract() == 0.0 => n,
                _ => return false,
            },
            "date" => match parse_number(value).or_else(|| parse_date(value, self.model.locale).ok().map(|(d, _)| d as f64)) {
                Some(n) => n,
                None => return false,
            },
            "time" => match parse_number(value).or_else(|| parse_time(value)) {
                Some(n) => n,
                None => return false,
            },
            _ => return true,
        };
        let Some(a) = self.bound(sheet, kind, &v.source) else { return true };
        let eps = 1e-9;
        match v.op.as_str() {
            "between" | "" | "notBetween" => {
                let Some(b) = self.bound(sheet, kind, &v.source2) else { return true };
                let inside = x >= a.min(b) - eps && x <= a.max(b) + eps;
                inside == (v.op != "notBetween")
            }
            "equal" => (x - a).abs() <= eps,
            "notEqual" => (x - a).abs() > eps,
            "greaterThan" => x > a + eps,
            "lessThan" => x < a - eps,
            "greaterThanOrEqual" => x >= a - eps,
            "lessThanOrEqual" => x <= a + eps,
            _ => true,
        }
    }

    /// Error if `value` may not be typed into the cell (the rule has the "stop" style and the value breaks it).
    pub(crate) fn check_validation(&self, sheet: u32, row: i32, column: i32, value: &str) -> Result<(), String> {
        match self.violated(sheet, row, column, value) {
            Some(v) if v.error_style == 0 => Err(self.validation_error(&v)),
            _ => Ok(()),
        }
    }

    /// For the "warning" and "information" styles: what to tell the user before the value is kept.
    pub fn get_validation_notice(&self, sheet: u32, row: i32, column: i32, value: &str) -> Option<ValidationMessage> {
        let v = self.violated(sheet, row, column, value).filter(|v| v.error_style != 0)?;
        Some(ValidationMessage { style: v.error_style, title: v.error_title.clone(), message: self.validation_error(&v) })
    }

    /// The hint to show when the cell is selected.
    pub fn get_validation_prompt(&self, sheet: u32, row: i32, column: i32) -> Option<ValidationMessage> {
        let v = self.validation_at(sheet, row, column)?;
        (!v.prompt.is_empty() || !v.prompt_title.is_empty()).then(|| ValidationMessage { style: 0, title: v.prompt_title, message: v.prompt })
    }

    fn validation_error(&self, v: &Validation) -> String {
        if !v.error.is_empty() {
            return v.error.clone();
        }
        match v.kind.as_str() {
            "" | "list" => "Значение не входит в список допустимых.".to_string(),
            _ => "Введённое значение недопустимо.".to_string(),
        }
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
