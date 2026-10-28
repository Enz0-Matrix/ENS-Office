//! ENS-Office: sort a range, find and replace. Each change is one undo step.

use std::borrow::Cow;
use std::cmp::Ordering;

use serde::Serialize;

use crate::cell::CellValue;
use crate::expressions::types::{Area, CellReferenceIndex};

use super::{common::UserModel, history::Diff};

/// A cell found by [UserModel::find_all].
#[derive(Serialize, Debug, PartialEq)]
pub struct FoundCell {
    /// Sheet index.
    pub sheet: u32,
    /// Row (1-based).
    pub row: i32,
    /// Column (1-based).
    pub column: i32,
}

/// Options of find / replace (as in Excel's dialog).
#[derive(Clone, Copy, Debug, Default)]
pub struct FindOptions {
    /// Match upper/lower case exactly.
    pub match_case: bool,
    /// The whole cell must equal the text.
    pub whole_cell: bool,
    /// Look in formulas (cell content) instead of shown values.
    pub in_formulas: bool,
    /// Only this sheet (None: the whole workbook).
    pub sheet: Option<u32>,
}

fn matches(haystack: &str, needle: &str, o: FindOptions) -> bool {
    if o.match_case {
        if o.whole_cell { haystack == needle } else { haystack.contains(needle) }
    } else {
        let (h, n) = (haystack.to_lowercase(), needle.to_lowercase());
        if o.whole_cell { h == n } else { h.contains(&n) }
    }
}

/// Replaces every occurrence of `needle` (case-insensitively unless `match_case`).
fn replace_text(text: &str, needle: &str, with: &str, match_case: bool) -> String {
    if needle.is_empty() {
        return text.to_string();
    }
    if match_case {
        return text.replace(needle, with);
    }
    let lower: Vec<char> = text.to_lowercase().chars().collect();
    let chars: Vec<char> = text.chars().collect();
    let pat: Vec<char> = needle.to_lowercase().chars().collect();
    // Lowercasing may change length for a few exotic letters: fall back to exact.
    if lower.len() != chars.len() {
        return text.replace(needle, with);
    }
    let mut out = String::new();
    let mut i = 0;
    while i < chars.len() {
        if i + pat.len() <= lower.len() && lower[i..i + pat.len()] == pat[..] {
            out.push_str(with);
            i += pat.len();
        } else {
            out.push(chars[i]);
            i += 1;
        }
    }
    out
}

/// Excel's sort order: numbers < text < FALSE < TRUE; blanks always last.
fn compare(a: &CellValue, b: &CellValue) -> Ordering {
    fn rank(v: &CellValue) -> u8 {
        match v {
            CellValue::Number(_) => 0,
            CellValue::String(_) => 1,
            CellValue::Boolean(_) => 2,
            CellValue::None => 3,
        }
    }
    match (a, b) {
        (CellValue::Number(x), CellValue::Number(y)) => x.partial_cmp(y).unwrap_or(Ordering::Equal),
        (CellValue::String(x), CellValue::String(y)) => x.to_lowercase().cmp(&y.to_lowercase()),
        (CellValue::Boolean(x), CellValue::Boolean(y)) => x.cmp(y),
        _ => rank(a).cmp(&rank(b)),
    }
}

/// Lines touched when `count` lines from `start` move by `delta`.
fn moved_span(start: i32, count: i32, delta: i32) -> (i32, i32) {
    let end = start + count - 1;
    (start.min(start + delta), end.max(end + delta))
}

/// Where the line now at `target` was before `count` lines from `start` moved
/// by `delta` (the moved block lands at `start + delta`, the lines it jumped
/// over shift by `count` the other way).
fn moved_from(target: i32, start: i32, count: i32, delta: i32) -> i32 {
    let end = start + count - 1;
    if target >= start + delta && target <= end + delta {
        target - delta
    } else if delta > 0 && target >= start && target < start + delta {
        target + count
    } else if delta < 0 && target > end + delta && target <= end {
        target - count
    } else {
        target
    }
}

/// Where rows with an empty sort key go.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Blanks {
    /// At the end, in both orders (as Excel).
    Last,
    /// At the beginning.
    First,
    /// Empty counts as 0: it sorts among the numbers.
    AsZero,
}

impl UserModel<'_> {
    /// ENS-Office: heights of the rows that differ from the default, for the grid:
    /// `[row, height, row, height, …]` (hidden rows have height 0). The grid
    /// computes positions from this without asking for every row.
    pub fn get_row_layout(&self, sheet: u32) -> Result<Vec<f64>, String> {
        let ws = self.model.workbook.worksheet(sheet)?;
        let mut out = Vec::with_capacity(ws.rows.len() * 2);
        for r in &ws.rows {
            out.push(f64::from(r.r));
            out.push(if r.hidden { 0.0 } else { r.height * crate::constants::ROW_HEIGHT_FACTOR });
        }
        Ok(out)
    }

    /// ENS-Office: page setup for printing (JSON, see `Worksheet::ens_page_setup`).
    pub fn get_page_setup(&self, sheet: u32) -> Result<String, String> {
        Ok(self.model.workbook.worksheet(sheet)?.ens_page_setup.clone())
    }

    /// ENS-Office: changes the page setup (an undoable step, so the book counts as changed).
    pub fn set_page_setup(&mut self, sheet: u32, json: &str) -> Result<(), String> {
        let ws = self.model.workbook.worksheet_mut(sheet)?;
        if ws.ens_page_setup == json {
            return Ok(());
        }
        let old_value = std::mem::replace(&mut ws.ens_page_setup, json.to_string());
        self.push_diff_list(vec![Diff::SetPageSetup { sheet, new_value: json.to_string(), old_value }]);
        Ok(())
    }

    /// ENS-Office: where the outline "+/−" buttons stand, `[below, right]` (see `Worksheet::ens_summary`).
    pub fn get_outline_summary(&self, sheet: u32) -> Result<(bool, bool), String> {
        Ok(self.model.workbook.worksheet(sheet)?.ens_summary)
    }

    /// ENS-Office: moves the outline "+/−" buttons (an undoable step).
    pub fn set_outline_summary(&mut self, sheet: u32, below: bool, right: bool) -> Result<(), String> {
        let ws = self.model.workbook.worksheet_mut(sheet)?;
        let old_value = std::mem::replace(&mut ws.ens_summary, (below, right));
        if old_value != (below, right) {
            self.push_diff_list(vec![Diff::SetOutlineSummary { sheet, new_value: (below, right), old_value }]);
        }
        Ok(())
    }

    /// ENS-Office: what the sheet's filters show (JSON, see `Worksheet::ens_filter_criteria`).
    pub fn get_filter_criteria(&self, sheet: u32) -> Result<String, String> {
        Ok(self.model.workbook.worksheet(sheet)?.ens_filter_criteria.clone())
    }

    /// ENS-Office: remembers what the sheet's filters show (saved to xlsx).
    pub fn set_filter_criteria(&mut self, sheet: u32, json: &str) -> Result<(), String> {
        self.model.workbook.worksheet_mut(sheet)?.ens_filter_criteria = json.to_string();
        Ok(())
    }

    /// ENS-Office: outline (grouping) of rows, `[row, level, collapsed (0/1), …]`
    /// for the rows in a group or carrying a "collapsed" flag.
    pub fn get_row_outline(&self, sheet: u32) -> Result<Vec<i32>, String> {
        let ws = self.model.workbook.worksheet(sheet)?;
        let mut rows: Vec<&crate::types::Row> = ws.rows.iter().filter(|r| r.outline_level > 0 || r.collapsed).collect();
        rows.sort_by_key(|r| r.r);
        Ok(rows.iter().flat_map(|r| [r.r, i32::from(r.outline_level), i32::from(r.collapsed)]).collect())
    }

    /// ENS-Office: outline (grouping) of columns, `[column, level, collapsed (0/1), …]`.
    pub fn get_column_outline(&self, sheet: u32) -> Result<Vec<i32>, String> {
        let ws = self.model.workbook.worksheet(sheet)?;
        let mut out = Vec::new();
        for c in ws.cols.iter().filter(|c| c.outline_level > 0 || c.collapsed) {
            for column in c.min..=c.max.min(c.min + 1000) {
                out.extend([column, i32::from(c.outline_level), i32::from(c.collapsed)]);
            }
        }
        Ok(out)
    }

    /// ENS-Office: sets the outline level of rows `r1..=r2` (one undo step).
    pub fn set_rows_outline(&mut self, sheet: u32, r1: i32, r2: i32, level: u8) -> Result<(), String> {
        let mut diff_list = Vec::new();
        for row in r1..=r2 {
            let ws = self.model.workbook.worksheet_mut(sheet)?;
            let old = ws.row_outline(row);
            let new = (level.min(7), old.1);
            if old != new {
                ws.set_row_outline(row, new.0, new.1)?;
                diff_list.push(Diff::SetRowOutline { sheet, row, new_value: new, old_value: old });
            }
        }
        if !diff_list.is_empty() {
            self.push_diff_list(diff_list);
        }
        Ok(())
    }

    /// ENS-Office: marks the group ending next to `row` as collapsed or not.
    pub fn set_row_collapsed(&mut self, sheet: u32, row: i32, collapsed: bool) -> Result<(), String> {
        let ws = self.model.workbook.worksheet_mut(sheet)?;
        let old = ws.row_outline(row);
        let new = (old.0, collapsed);
        if old != new {
            ws.set_row_outline(row, new.0, new.1)?;
            self.push_diff_list(vec![Diff::SetRowOutline { sheet, row, new_value: new, old_value: old }]);
        }
        Ok(())
    }

    /// ENS-Office: sets the outline level of columns `c1..=c2` (one undo step).
    pub fn set_columns_outline(&mut self, sheet: u32, c1: i32, c2: i32, level: u8) -> Result<(), String> {
        let mut diff_list = Vec::new();
        for column in c1..=c2 {
            let ws = self.model.workbook.worksheet_mut(sheet)?;
            let old = ws.column_outline(column);
            let new = (level.min(7), old.1);
            if old != new {
                ws.set_column_outline(column, new.0, new.1)?;
                diff_list.push(Diff::SetColumnOutline { sheet, column, new_value: new, old_value: old });
            }
        }
        if !diff_list.is_empty() {
            self.push_diff_list(diff_list);
        }
        Ok(())
    }

    /// ENS-Office: marks the column group ending next to `column` as collapsed or not.
    pub fn set_column_collapsed(&mut self, sheet: u32, column: i32, collapsed: bool) -> Result<(), String> {
        let ws = self.model.workbook.worksheet_mut(sheet)?;
        let old = ws.column_outline(column);
        let new = (old.0, collapsed);
        if old != new {
            ws.set_column_outline(column, new.0, new.1)?;
            self.push_diff_list(vec![Diff::SetColumnOutline { sheet, column, new_value: new, old_value: old }]);
        }
        Ok(())
    }

    /// New workbooks: cells are centred vertically by default (the user's wish;
    /// Excel's default is "bottom"). Changes the workbook's default cell format,
    /// so it is saved to xlsx and Excel shows the same. Not an undo step.
    pub fn set_default_vertical_center(&mut self) {
        if let Some(xf) = self.model.workbook.styles.cell_xfs.first_mut() {
            let mut alignment = xf.alignment.clone().unwrap_or_default();
            alignment.vertical = crate::types::VerticalAlignment::Center;
            xf.alignment = Some(alignment);
            xf.apply_alignment = true;
        }
    }

    /// ENS-Office: a step that restores the rules of the sheet after rows or columns were deleted
    /// (empty if nothing changed).
    pub(crate) fn cf_list_diff(&self, sheet: u32, old_value: Vec<crate::cf_types::ConditionalFormatting>) -> Result<Vec<Diff>, String> {
        let new_value = self.model.workbook.worksheet(sheet)?.conditional_formatting.clone();
        Ok(if new_value == old_value { vec![] } else { vec![Diff::SetConditionalFormattingList { sheet, old_value, new_value }] })
    }

    /// ENS-Office: the default font of a new book (the font of the first cell style).
    pub fn set_default_font(&mut self, name: &str, size: i32) {
        let styles = &mut self.model.workbook.styles;
        let Some(id) = styles.cell_xfs.first().map(|xf| xf.font_id) else { return };
        if let Some(font) = usize::try_from(id).ok().and_then(|i| styles.fonts.get_mut(i)) {
            font.name = name.to_string();
            font.sz = size;
        }
    }

    /// Position in the undo history, for [`Self::merge_undo_since`].
    pub fn undo_mark(&self) -> usize {
        self.history.undo_stack.len()
    }

    /// Joins every undo step made since `mark` into one, so one Ctrl+Z undoes
    /// an action applied to several ranges (a Ctrl selection) at once.
    pub fn merge_undo_since(&mut self, mark: usize) {
        let stack = &mut self.history.undo_stack;
        if stack.len() > mark + 1 {
            let merged: Vec<Diff> = stack.drain(mark..).flatten().collect();
            stack.push(merged);
        }
    }

    /// The cell's value shown with number format `code`, without changing the
    /// workbook (the sample in the "Format cells" dialog). Text and booleans are
    /// shown as they are; an empty cell gives "".
    pub fn format_cell_value_as(&self, sheet: u32, row: i32, column: i32, code: &str) -> Result<String, String> {
        let ws = self.model.workbook.worksheet(sheet)?;
        let Some(cell) = ws.cell(row, column) else {
            return Ok(String::new());
        };
        Ok(cell.formatted_value(&self.model.workbook.shared_strings, self.model.language, |value| {
            crate::formatter::format::format_number(value, code, self.model.locale).text
        }))
    }

    /// Palette colour (Excel index 1..56, `[Red]` = 3) that the number format
    /// gives the cell's value, 0 if none: "#,##0;[Red]-#,##0" paints negatives red.
    pub fn get_formatted_cell_color(&self, sheet: u32, row: i32, column: i32) -> Result<i32, String> {
        let ws = self.model.workbook.worksheet(sheet)?;
        let Some(cell) = ws.cell(row, column) else {
            return Ok(0);
        };
        let CellValue::Number(value) = cell.value(&self.model.workbook.shared_strings, self.model.language) else {
            return Ok(0);
        };
        let format = self.model.get_style_for_cell(sheet, row, column)?.num_fmt;
        Ok(crate::formatter::format::format_number(value, &format, self.model.locale).color.unwrap_or(0))
    }

    /// Moves rows `[row, row + row_count - 1]` by exactly `delta` rows (the rows
    /// in between shift the other way), as Excel's "Insert cut cells". Unlike
    /// `move_rows_action` hidden rows are not skipped, and row heights and the
    /// hidden state travel with their rows. One undo step.
    pub fn move_rows_exact(&mut self, sheet: u32, row: i32, row_count: i32, delta: i32) -> Result<(), String> {
        self.guard_sheet(sheet)?;
        if delta == 0 || row_count <= 0 {
            return Ok(());
        }
        let (lo, hi) = moved_span(row, row_count, delta);
        let mut before = Vec::new();
        for r in lo..=hi {
            before.push((self.model.get_row_height(sheet, r)?, self.model.is_row_hidden(sheet, r)?));
        }
        let old_merged_cells = self.model.get_merged_cells(sheet)?.to_vec();
        self.model.move_rows_action(sheet, row, row_count, delta)?;
        let mut diff_list = self.merged_cells_snapshot_diffs(sheet, old_merged_cells)?;
        diff_list.push(Diff::MoveRows { sheet, row, row_count, delta });
        for r in lo..=hi {
            let (height, hidden) = before[(moved_from(r, row, row_count, delta) - lo) as usize];
            // The engine may already move some of it: compare with what is there now.
            let (old_height, old_hidden) = (self.model.get_row_height(sheet, r)?, self.model.is_row_hidden(sheet, r)?);
            if height != old_height {
                self.model.set_row_height(sheet, r, height)?;
                diff_list.push(Diff::SetRowHeight { sheet, row: r, new_value: height, old_value: old_height });
            }
            if hidden != old_hidden {
                self.model.set_row_hidden(sheet, r, hidden)?;
                diff_list.push(Diff::SetRowHidden { sheet, row: r, new_value: hidden, old_value: old_hidden });
            }
        }
        self.push_diff_list(diff_list);
        self.evaluate_if_not_paused();
        Ok(())
    }

    /// Columns version of [`Self::move_rows_exact`]: widths and the hidden
    /// state travel with their columns.
    pub fn move_columns_exact(&mut self, sheet: u32, column: i32, column_count: i32, delta: i32) -> Result<(), String> {
        self.guard_sheet(sheet)?;
        if delta == 0 || column_count <= 0 {
            return Ok(());
        }
        let (lo, hi) = moved_span(column, column_count, delta);
        let mut before = Vec::new();
        for c in lo..=hi {
            before.push((self.model.get_column_width(sheet, c)?, self.model.is_column_hidden(sheet, c)?));
        }
        let old_merged_cells = self.model.get_merged_cells(sheet)?.to_vec();
        self.model.move_columns_action(sheet, column, column_count, delta)?;
        let mut diff_list = self.merged_cells_snapshot_diffs(sheet, old_merged_cells)?;
        diff_list.push(Diff::MoveColumns { sheet, column, column_count, delta });
        for c in lo..=hi {
            let (width, hidden) = before[(moved_from(c, column, column_count, delta) - lo) as usize];
            let (old_width, old_hidden) = (self.model.get_column_width(sheet, c)?, self.model.is_column_hidden(sheet, c)?);
            if width != old_width {
                self.model.set_column_width(sheet, c, width)?;
                diff_list.push(Diff::SetColumnWidth { sheet, column: c, new_value: width, old_value: old_width });
            }
            if hidden != old_hidden {
                self.model.set_column_hidden(sheet, c, hidden)?;
                diff_list.push(Diff::SetColumnHidden { sheet, column: c, new_value: hidden, old_value: old_hidden });
            }
        }
        self.push_diff_list(diff_list);
        self.evaluate_if_not_paused();
        Ok(())
    }

    /// Sorts the rows of `area` by the values in `key_column` (a column of the
    /// sheet inside the area). Formulas move with their rows and their relative
    /// references are adjusted, as in Excel. Blank keys stay at the end.
    pub fn sort_range(&mut self, area: &Area, key_column: i32, descending: bool) -> Result<(), String> {
        self.sort_range_by(area, &[(key_column, descending)], Blanks::Last)
    }

    /// Sorts the rows of `area` by several columns (`(column, descending)`, the
    /// first one decides, the next ones break ties). `blanks` says where rows
    /// with an empty key go.
    pub fn sort_range_by(&mut self, area: &Area, keys: &[(i32, bool)], blanks: Blanks) -> Result<(), String> {
        self.sort_range_by_rows(area, keys, blanks, false)
    }

    /// As [`Self::sort_range_by`]; with `skip_hidden` the hidden rows (rows
    /// filtered out by the autofilter) stay where they are and only the
    /// visible rows are reordered among themselves, as Excel does with an
    /// active filter.
    pub fn sort_range_by_rows(&mut self, area: &Area, keys: &[(i32, bool)], blanks: Blanks, skip_hidden: bool) -> Result<(), String> {
        self.guard_area(area)?;
        let sheet = area.sheet;
        if keys.is_empty() {
            return Ok(());
        }
        for &(key_column, _) in keys {
            if key_column < area.column || key_column >= area.column + area.width {
                return Err("Столбец сортировки вне диапазона".to_string());
            }
        }
        let (r1, c1) = (area.row, area.column);
        let (r2, c2) = (area.row + area.height - 1, area.column + area.width - 1);
        // As Excel: merged cells are sorted as blocks, so they must all have the same
        // size and stand the same way in every block of rows.
        let block = {
            let merges: Vec<_> = self.model.get_merged_cells(sheet)?
                .into_iter()
                .filter(|m| m.intersects(r1, c1, area.width, area.height))
                .collect();
            let message = "Для сортировки все объединённые ячейки должны быть одного размера и стоять одинаково в каждой строке (блоке строк)".to_string();
            if merges.is_empty() {
                1
            } else {
                let h = merges[0].height;
                if area.height % h != 0
                    || merges.iter().any(|m| {
                        m.height != h || m.row < r1 || m.row + m.height - 1 > r2 || m.column < c1 || m.column + m.width - 1 > c2 || (m.row - r1) % h != 0
                    })
                {
                    return Err(message);
                }
                let mut patterns = vec![std::collections::BTreeSet::new(); (area.height / h) as usize];
                for m in &merges {
                    patterns[((m.row - r1) / h) as usize].insert((m.column, m.width));
                }
                if patterns.iter().any(|p| p != &patterns[0]) {
                    return Err(message);
                }
                h
            }
        };
        let hidden: std::collections::HashSet<i32> = if skip_hidden {
            self.model.workbook.worksheet(sheet)?.rows.iter().filter(|r| r.hidden).map(|r| r.r).collect()
        } else {
            Default::default()
        };
        // The rows that take part; a sorted row goes to the slot of its position.
        let slots: Vec<i32> = (r1..=r2).step_by(block as usize).filter(|r| !hidden.contains(r)).collect();
        // values[k][i] — key k of row slots[i]
        let values: Vec<Vec<CellValue>> = keys
            .iter()
            .map(|&(column, _)| {
                slots.iter()
                    .map(|&r| match self.model.get_cell_value_by_index(sheet, r, column).unwrap_or(CellValue::None) {
                        // "Пустые как ноль": строки не разъезжаются — пустое стоит среди чисел как 0.
                        CellValue::None if blanks == Blanks::AsZero => CellValue::Number(0.0),
                        v => v,
                    })
                    .collect()
            })
            .collect();
        let mut order: Vec<usize> = (0..slots.len()).collect();
        order.sort_by(|&a, &b| {
            for (k, &(_, descending)) in keys.iter().enumerate() {
                let (ka, kb) = (&values[k][a], &values[k][b]);
                let first_blank = if blanks == Blanks::First { Ordering::Less } else { Ordering::Greater };
                let o = match (ka, kb) {
                    (CellValue::None, CellValue::None) => Ordering::Equal,
                    (CellValue::None, _) => first_blank,
                    (_, CellValue::None) => first_blank.reverse(),
                    _ if descending => compare(kb, ka),
                    _ => compare(ka, kb),
                };
                if o != Ordering::Equal {
                    return o;
                }
            }
            Ordering::Equal
        });
        if order.iter().enumerate().all(|(i, &j)| i == j) {
            return Ok(()); // already sorted
        }
        // Read everything first: the target rows are also sources.
        let mut planned = Vec::new();
        for (i, &j) in order.iter().enumerate() {
            let (source_row, target_row) = (slots[j], slots[i]);
            if source_row == target_row {
                continue;
            }
            for offset in 0..block {
                let (source_row, target_row) = (source_row + offset, target_row + offset);
                for column in c1..=c2 {
                    let content = self.model.get_localized_cell_content(sheet, source_row, column)?;
                    let source = CellReferenceIndex { sheet, row: source_row, column };
                    let target = CellReferenceIndex { sheet, row: target_row, column };
                    let value = if content.is_empty() {
                        String::new()
                    } else {
                        self.model.extend_copied_value(&content, &source, &target)?
                    };
                    let style = self.model.get_cell_style_or_none(sheet, source_row, column)?;
                    planned.push((target_row, column, value, style));
                }
            }
        }
        let mut diff_list = Vec::new();
        for (row, column, value, style) in planned {
            let old_value = self.model.workbook.worksheet(sheet)?.cell(row, column).map(Cow::into_owned);
            let old_style = self.model.get_cell_style_or_none(sheet, row, column)?;
            if value.is_empty() {
                if old_value.is_some() {
                    self.model.range_clear_contents(&Area { sheet, row, column, width: 1, height: 1 })?;
                    diff_list.push(Diff::RangeClearContents {
                        sheet, row, column, width: 1, height: 1, old_value: vec![vec![old_value]],
                    });
                }
            } else {
                self.model.set_user_input(sheet, row, column, value.clone())?;
                diff_list.push(Diff::SetCellValue { sheet, row, column, new_value: value, old_value: Box::new(old_value) });
            }
            let new_style = style.unwrap_or_default();
            if old_style.as_ref() != Some(&new_style) {
                self.model.set_cell_style(sheet, row, column, &new_style)?;
                diff_list.push(Diff::SetCellStyle {
                    sheet, row, column, old_value: Box::new(old_style), new_value: Box::new(new_style),
                });
            }
        }
        self.push_diff_list(diff_list);
        self.evaluate_if_not_paused();
        Ok(())
    }

    /// All cells whose shown value (or formula) contains `text`, sheet by sheet,
    /// row by row, as Excel's "Find All".
    pub fn find_all(&self, text: &str, options: FindOptions) -> Result<Vec<FoundCell>, String> {
        let mut found = Vec::new();
        if text.is_empty() {
            return Ok(found);
        }
        for (index, ws) in self.model.workbook.worksheets.iter().enumerate() {
            let sheet = index as u32;
            if options.sheet.is_some_and(|s| s != sheet) {
                continue;
            }
            let mut cells: Vec<(i32, i32)> = ws.sheet_data.cells().map(|(r, c, _)| (r, c)).collect();
            cells.sort();
            for (row, column) in cells {
                let hay = if options.in_formulas {
                    self.model.get_localized_cell_content(sheet, row, column)?
                } else {
                    self.model.get_formatted_cell_value(sheet, row, column)?
                };
                if !hay.is_empty() && matches(&hay, text, options) {
                    found.push(FoundCell { sheet, row, column });
                }
            }
        }
        Ok(found)
    }

    /// Replaces `text` with `with` in the content of the given cells (or of
    /// every match when `cells` is None). Returns how many cells changed.
    pub fn replace_all(&mut self, text: &str, with: &str, options: FindOptions, cells: Option<Vec<FoundCell>>) -> Result<usize, String> {
        let targets = match cells {
            Some(c) => c,
            None => self.find_all(text, FindOptions { in_formulas: true, ..options })?,
        };
        let mut diff_list = Vec::new();
        for FoundCell { sheet, row, column } in targets {
            let content = self.model.get_localized_cell_content(sheet, row, column)?;
            let new_value = if options.whole_cell {
                if matches(&content, text, options) { with.to_string() } else { continue }
            } else {
                replace_text(&content, text, with, options.match_case)
            };
            if new_value == content {
                continue;
            }
            let old_value = self.model.workbook.worksheet(sheet)?.cell(row, column).map(Cow::into_owned);
            self.model.set_user_input(sheet, row, column, new_value.clone())?;
            diff_list.push(Diff::SetCellValue { sheet, row, column, new_value, old_value: Box::new(old_value) });
        }
        let changed = diff_list.len();
        if changed > 0 {
            self.push_diff_list(diff_list);
            self.evaluate_if_not_paused();
        }
        Ok(changed)
    }
}

/// Excel keeps the print area of a sheet in this sheet-scoped name.
pub const PRINT_AREA: &str = "_xlnm.Print_Area";

/// Excel keeps the rows and columns repeated on every printed page in this sheet-scoped name.
pub const PRINT_TITLES: &str = "_xlnm.Print_Titles";

/// "Лист1!$A:$B,Лист1!$1:$3" → [row1, row2, column1, column2]; 0 — not set.
fn parse_print_titles(formula: &str) -> [i32; 4] {
    let mut out = [0; 4];
    for part in formula.split(',') {
        let range = part.rsplit('!').next().unwrap_or_default().trim().trim_start_matches('=').replace('$', "");
        let (a, b) = range.split_once(':').unwrap_or((&range, &range));
        if let (Ok(r1), Ok(r2)) = (a.parse::<i32>(), b.parse::<i32>()) {
            out[0] = r1;
            out[1] = r2;
        } else if let (Ok(c1), Ok(c2)) = (
            crate::expressions::utils::column_to_number(a),
            crate::expressions::utils::column_to_number(b),
        ) {
            out[2] = c1;
            out[3] = c2;
        }
    }
    out
}

/// Excel keeps the autofilter range of a sheet in this hidden, sheet-scoped name.
pub const FILTER_DATABASE: &str = "_xlnm._FilterDatabase";

/// "Лист1!$A$1:$C$6" → (row1, column1, row2, column2).
pub fn parse_filter_range(formula: &str) -> Option<(i32, i32, i32, i32)> {
    let range = formula.rsplit('!').next()?.trim_start_matches('=');
    let mut parts = range.split(':');
    let a = crate::expressions::utils::parse_reference_a1(parts.next()?)?;
    let b = match parts.next() {
        Some(p) => crate::expressions::utils::parse_reference_a1(p)?,
        None => crate::expressions::utils::parse_reference_a1(range)?,
    };
    Some((a.row, a.column, b.row, b.column))
}

impl UserModel<'_> {
    /// ENS-Office: the print area of the sheet, [row1, column1, row2, column2].
    /// Several areas (Excel allows a list) — the first one.
    pub fn get_print_area(&self, sheet: u32) -> Result<Option<[i32; 4]>, String> {
        let sheet_id = self.model.workbook.worksheet(sheet)?.sheet_id;
        Ok(self
            .model
            .workbook
            .defined_names
            .iter()
            .find(|d| d.sheet_id == Some(sheet_id) && d.name.eq_ignore_ascii_case(PRINT_AREA))
            .and_then(|d| parse_filter_range(d.formula.split(',').next().unwrap_or_default()))
            .map(|(r1, c1, r2, c2)| [r1, c1, r2, c2]))
    }

    pub(crate) fn write_print_area(&mut self, sheet: u32, range: Option<[i32; 4]>) -> Result<(), String> {
        let ws = self.model.workbook.worksheet(sheet)?;
        let (sheet_id, name) = (ws.sheet_id, ws.get_name());
        let names = &mut self.model.workbook.defined_names;
        names.retain(|d| !(d.sheet_id == Some(sheet_id) && d.name.eq_ignore_ascii_case(PRINT_AREA)));
        if let Some([r1, c1, r2, c2]) = range {
            let col = |c: i32| crate::expressions::utils::number_to_column(c).unwrap_or_default();
            names.push(crate::types::DefinedName {
                name: PRINT_AREA.to_string(),
                formula: format!("{}!${}${r1}:${}${r2}", crate::expressions::utils::quote_name(&name), col(c1), col(c2)),
                sheet_id: Some(sheet_id),
            });
        }
        Ok(())
    }

    /// ENS-Office: sets (or with None removes) the print area of the sheet; an undoable step.
    pub fn set_print_area(&mut self, sheet: u32, range: Option<[i32; 4]>) -> Result<(), String> {
        let old_value = self.get_print_area(sheet)?;
        if old_value == range {
            return Ok(());
        }
        self.write_print_area(sheet, range)?;
        self.push_diff_list(vec![Diff::SetPrintArea { sheet, new_value: range, old_value }]);
        Ok(())
    }

    /// ENS-Office: rows and columns repeated on every printed page, [row1, row2, column1, column2]
    /// (0 — not set), Excel's `_xlnm.Print_Titles`.
    pub fn get_print_titles(&self, sheet: u32) -> Result<[i32; 4], String> {
        let sheet_id = self.model.workbook.worksheet(sheet)?.sheet_id;
        Ok(self
            .model
            .workbook
            .defined_names
            .iter()
            .find(|d| d.sheet_id == Some(sheet_id) && d.name.eq_ignore_ascii_case(PRINT_TITLES))
            .map(|d| parse_print_titles(&d.formula))
            .unwrap_or([0; 4]))
    }

    pub(crate) fn write_print_titles(&mut self, sheet: u32, titles: [i32; 4]) -> Result<(), String> {
        let ws = self.model.workbook.worksheet(sheet)?;
        let (sheet_id, name) = (ws.sheet_id, ws.get_name());
        let names = &mut self.model.workbook.defined_names;
        names.retain(|d| !(d.sheet_id == Some(sheet_id) && d.name.eq_ignore_ascii_case(PRINT_TITLES)));
        let q = crate::expressions::utils::quote_name(&name);
        let col = |c: i32| crate::expressions::utils::number_to_column(c).unwrap_or_default();
        let mut parts = Vec::new();
        if titles[2] > 0 && titles[3] >= titles[2] {
            parts.push(format!("{q}!${}:${}", col(titles[2]), col(titles[3])));
        }
        if titles[0] > 0 && titles[1] >= titles[0] {
            parts.push(format!("{q}!${}:${}", titles[0], titles[1]));
        }
        if !parts.is_empty() {
            names.push(crate::types::DefinedName { name: PRINT_TITLES.to_string(), formula: parts.join(","), sheet_id: Some(sheet_id) });
        }
        Ok(())
    }

    /// ENS-Office: sets the repeated rows and columns, [row1, row2, column1, column2] (0 — none); an undoable step.
    pub fn set_print_titles(&mut self, sheet: u32, titles: [i32; 4]) -> Result<(), String> {
        let old_value = self.get_print_titles(sheet)?;
        if old_value == titles {
            return Ok(());
        }
        self.write_print_titles(sheet, titles)?;
        self.push_diff_list(vec![Diff::SetPrintTitles { sheet, new_value: titles, old_value }]);
        Ok(())
    }

    pub(crate) fn write_comment(
        &mut self,
        sheet: u32,
        row: i32,
        column: i32,
        comment: Option<crate::types::Comment>,
    ) -> Result<(), String> {
        let cell_ref = comment_ref(row, column)?;
        let ws = self.model.workbook.worksheet_mut(sheet)?;
        ws.comments.retain(|c| !c.cell_ref.eq_ignore_ascii_case(&cell_ref));
        if let Some(c) = comment {
            ws.comments.push(c);
        }
        Ok(())
    }

    pub(crate) fn write_rich_text(
        &mut self,
        sheet: u32,
        row: i32,
        column: i32,
        runs: Option<Vec<crate::types::RichRun>>,
        text: &str,
    ) -> Result<(), String> {
        let style = self.model.workbook.worksheet(sheet)?.cell(row, column).map_or(0, |c| c.get_style());
        match runs {
            Some(r) => self.model.set_cell_with_rich_text(sheet, row, column, r, style),
            None => self.model.set_user_input(sheet, row, column, text.to_string()),
        }
    }

    /// ENS-Office: the runs of the rich text of a cell (`None` — a plain cell).
    pub fn get_rich_text(&self, sheet: u32, row: i32, column: i32) -> Result<Option<Vec<crate::types::RichRun>>, String> {
        Ok(match self.model.workbook.worksheet(sheet)?.cell(row, column).as_deref() {
            Some(crate::types::Cell::SharedString { si, .. }) => self.model.workbook.ens_rich.get(si).cloned(),
            _ => None,
        })
    }

    /// ENS-Office: cells with a rich text in the rectangle, flat `[row, column, …]`.
    pub fn rich_text_cells_in(&self, sheet: u32, r1: i32, c1: i32, r2: i32, c2: i32) -> Result<Vec<i32>, String> {
        let mut out = Vec::new();
        if self.model.workbook.ens_rich.is_empty() {
            return Ok(out);
        }
        let ws = self.model.workbook.worksheet(sheet)?;
        for row in r1..=r2 {
            for (column, cell) in ws.sheet_data.cells_in_row_between(row, c1, c2) {
                if let crate::types::Cell::SharedString { si, .. } = &*cell {
                    if self.model.workbook.ens_rich.contains_key(si) {
                        out.extend([row, column]);
                    }
                }
            }
        }
        Ok(out)
    }

    /// ENS-Office: sets a text cell made of formatted runs (no runs, or one plain run — a plain text);
    /// an undoable step.
    pub fn set_rich_text(&mut self, sheet: u32, row: i32, column: i32, runs: Vec<crate::types::RichRun>) -> Result<(), String> {
        self.guard_cells(sheet, row, column, row, column)?;
        let new_text: String = runs.iter().map(|r| r.text.as_str()).collect();
        let plain = |r: &crate::types::RichRun| !(r.b || r.i || r.u || r.strike || r.sz.is_some() || r.color.is_some() || r.name.is_some() || r.vert != 0);
        let new_value = if runs.iter().all(plain) { None } else { Some(runs) };
        let old_value = self.get_rich_text(sheet, row, column)?;
        let old_text = self.model.get_localized_cell_content(sheet, row, column)?;
        if old_value == new_value && old_text == new_text {
            return Ok(());
        }
        self.write_rich_text(sheet, row, column, new_value.clone(), &new_text)?;
        self.model.evaluate();
        self.push_diff_list(vec![Diff::SetRichText { sheet, row, column, new_value, old_value, new_text, old_text }]);
        Ok(())
    }

    /// ENS-Office: the pictures of the sheet.
    pub fn get_images(&self, sheet: u32) -> Result<Vec<crate::types::Image>, String> {
        Ok(self.model.workbook.worksheet(sheet)?.ens_images.clone())
    }

    fn change_images(&mut self, sheet: u32, change: impl FnOnce(&mut Vec<crate::types::Image>) -> Result<(), String>) -> Result<(), String> {
        self.guard_sheet(sheet)?;
        let old_value = self.get_images(sheet)?;
        let mut new_value = old_value.clone();
        change(&mut new_value)?;
        if new_value == old_value {
            return Ok(());
        }
        self.model.workbook.worksheet_mut(sheet)?.ens_images = new_value.clone();
        self.push_diff_list(vec![Diff::SetImages { sheet, new_value, old_value }]);
        Ok(())
    }

    /// ENS-Office: puts a picture on the sheet; an undoable step. Returns its index.
    pub fn add_image(&mut self, sheet: u32, image: crate::types::Image) -> Result<usize, String> {
        if !matches!(image.ext.as_str(), "png" | "jpeg" | "gif") || image.data.is_empty() {
            return Err("Поддерживаются рисунки PNG, JPEG и GIF.".to_string());
        }
        let mut index = 0;
        self.change_images(sheet, |list| {
            list.push(image);
            index = list.len() - 1;
            Ok(())
        })?;
        Ok(index)
    }

    /// ENS-Office: moves / resizes a picture (top left corner at the cell, shift in px); an undoable step.
    pub fn place_image(&mut self, sheet: u32, index: usize, row: i32, column: i32, dx: i32, dy: i32, width: i32, height: i32) -> Result<(), String> {
        self.change_images(sheet, |list| {
            let img = list.get_mut(index).ok_or("No such picture")?;
            (img.row, img.column, img.dx, img.dy) = (row.max(1), column.max(1), dx, dy);
            (img.width, img.height) = (width.max(8), height.max(8));
            Ok(())
        })
    }

    /// ENS-Office: deletes a picture; an undoable step.
    pub fn delete_image(&mut self, sheet: u32, index: usize) -> Result<(), String> {
        self.change_images(sheet, |list| {
            if index >= list.len() {
                return Err("No such picture".to_string());
            }
            list.remove(index);
            Ok(())
        })
    }

    /// ENS-Office: all comments (notes) of the sheet.
    pub fn get_comments(&self, sheet: u32) -> Result<Vec<crate::types::Comment>, String> {
        Ok(self.model.workbook.worksheet(sheet)?.comments.clone())
    }

    /// ENS-Office: the comment of a cell.
    pub fn get_comment(&self, sheet: u32, row: i32, column: i32) -> Result<Option<crate::types::Comment>, String> {
        let cell_ref = comment_ref(row, column)?;
        Ok(self
            .model
            .workbook
            .worksheet(sheet)?
            .comments
            .iter()
            .find(|c| c.cell_ref.eq_ignore_ascii_case(&cell_ref))
            .cloned())
    }

    /// ENS-Office: sets (an empty text deletes) the comment of a cell; an undoable step.
    pub fn set_comment(&mut self, sheet: u32, row: i32, column: i32, text: &str, author: &str) -> Result<(), String> {
        self.guard_cells(sheet, row, column, row, column)?;
        let old_value = self.get_comment(sheet, row, column)?;
        let new_value = if text.is_empty() {
            None
        } else {
            Some(crate::types::Comment {
                text: text.to_string(),
                author_name: author.to_string(),
                author_id: old_value.as_ref().and_then(|c| c.author_id.clone()),
                cell_ref: comment_ref(row, column)?,
            })
        };
        if old_value == new_value {
            return Ok(());
        }
        self.write_comment(sheet, row, column, new_value.clone())?;
        self.push_diff_list(vec![Diff::SetComment { sheet, row, column, new_value, old_value }]);
        Ok(())
    }

    /// ENS-Office: deletes the comment of a cell; an undoable step.
    pub fn delete_comment(&mut self, sheet: u32, row: i32, column: i32) -> Result<(), String> {
        self.set_comment(sheet, row, column, "", "")
    }

    pub(crate) fn write_page_breaks(&mut self, sheet: u32, breaks: (Vec<i32>, Vec<i32>)) -> Result<(), String> {
        let ws = self.model.workbook.worksheet_mut(sheet)?;
        ws.ens_row_breaks = breaks.0;
        ws.ens_col_breaks = breaks.1;
        Ok(())
    }

    /// ENS-Office: manual page breaks, (rows after which a page ends, columns after which a page ends).
    pub fn get_page_breaks(&self, sheet: u32) -> Result<(Vec<i32>, Vec<i32>), String> {
        let ws = self.model.workbook.worksheet(sheet)?;
        Ok((ws.ens_row_breaks.clone(), ws.ens_col_breaks.clone()))
    }

    /// ENS-Office: inserts (`on`) or removes a break after the row (`is_row`) or column `index`; an undoable step.
    pub fn set_page_break(&mut self, sheet: u32, is_row: bool, index: i32, on: bool) -> Result<(), String> {
        if index < 1 {
            return Err(format!("Invalid page break position: {index}"));
        }
        let old_value = self.get_page_breaks(sheet)?;
        let mut new_value = old_value.clone();
        let list = if is_row { &mut new_value.0 } else { &mut new_value.1 };
        list.retain(|b| *b != index);
        if on {
            list.push(index);
            list.sort_unstable();
        }
        if new_value == old_value {
            return Ok(());
        }
        self.write_page_breaks(sheet, new_value.clone())?;
        self.push_diff_list(vec![Diff::SetPageBreaks { sheet, new_value, old_value }]);
        Ok(())
    }

    /// ENS-Office: removes all manual page breaks; an undoable step.
    pub fn clear_page_breaks(&mut self, sheet: u32) -> Result<(), String> {
        let old_value = self.get_page_breaks(sheet)?;
        if old_value.0.is_empty() && old_value.1.is_empty() {
            return Ok(());
        }
        let new_value = (vec![], vec![]);
        self.write_page_breaks(sheet, new_value.clone())?;
        self.push_diff_list(vec![Diff::SetPageBreaks { sheet, new_value, old_value }]);
        Ok(())
    }

    /// ENS-Office: the autofilter range of the sheet, [row1, column1, row2, column2].
    pub fn get_auto_filter(&self, sheet: u32) -> Result<Option<[i32; 4]>, String> {
        let sheet_id = self.model.workbook.worksheet(sheet)?.sheet_id;
        Ok(self
            .model
            .workbook
            .defined_names
            .iter()
            .find(|d| d.sheet_id == Some(sheet_id) && d.name.eq_ignore_ascii_case(FILTER_DATABASE))
            .and_then(|d| parse_filter_range(&d.formula))
            .map(|(r1, c1, r2, c2)| [r1, c1, r2, c2]))
    }

    /// ENS-Office: set (or remove with None) the autofilter range of the sheet.
    /// Saved to xlsx as `<autoFilter>`, as Excel does.
    pub fn set_auto_filter(&mut self, sheet: u32, range: Option<[i32; 4]>) -> Result<(), String> {
        let ws = self.model.workbook.worksheet(sheet)?;
        let (sheet_id, name) = (ws.sheet_id, ws.get_name());
        let names = &mut self.model.workbook.defined_names;
        names.retain(|d| !(d.sheet_id == Some(sheet_id) && d.name.eq_ignore_ascii_case(FILTER_DATABASE)));
        if let Some([r1, c1, r2, c2]) = range {
            let col = |c: i32| crate::expressions::utils::number_to_column(c).unwrap_or_default();
            names.push(crate::types::DefinedName {
                name: FILTER_DATABASE.to_string(),
                formula: format!("{}!${}${r1}:${}${r2}", crate::expressions::utils::quote_name(&name), col(c1), col(c2)),
                sheet_id: Some(sheet_id),
            });
        }
        Ok(())
    }

    /// Used range of the sheet: [min_row, min_column, max_row, max_column]
    /// (all 1 for an empty sheet). For printing.
    pub fn get_sheet_dimension(&self, sheet: u32) -> Result<[i32; 4], String> {
        let d = self.model.workbook.worksheet(sheet)?.dimension();
        Ok([d.min_row, d.min_column, d.max_row, d.max_column])
    }

    /// "Paste values": the values (not formulas) of the source range, at the
    /// selected cell, optionally transposed. Numbers keep full precision.
    pub fn paste_values(&mut self, source_sheet: u32, source: (i32, i32, i32, i32), transpose: bool) -> Result<(), String> {
        self.guard_selection()?;
        let view = self.get_selected_view();
        let (sheet, row0, column0) = (view.sheet, view.row, view.column);
        let (r1, c1, r2, c2) = source;
        let decimal = self.model.locale.numbers.symbols.decimal.clone();
        let (yes, no) = (self.model.language.booleans.r#true.clone(), self.model.language.booleans.r#false.clone());
        let mut planned = Vec::new();
        for r in r1..=r2 {
            for c in c1..=c2 {
                let (dr, dc) = if transpose { (c - c1, r - r1) } else { (r - r1, c - c1) };
                let text = match self.model.get_cell_value_by_index(source_sheet, r, c)? {
                    CellValue::Number(n) => n.to_string().replace('.', &decimal),
                    // A text that would be read as a number or formula stays text.
                    CellValue::String(t) => {
                        if t.starts_with('=') || t.parse::<f64>().is_ok() || t.replace(&decimal, ".").parse::<f64>().is_ok() {
                            format!("'{t}")
                        } else {
                            t
                        }
                    }
                    CellValue::Boolean(b) => if b { yes.clone() } else { no.clone() },
                    CellValue::None => String::new(),
                };
                planned.push((row0 + dr, column0 + dc, text));
            }
        }
        let mut diff_list = Vec::new();
        for (row, column, value) in planned {
            let old_value = self.model.workbook.worksheet(sheet)?.cell(row, column).map(Cow::into_owned);
            if value.is_empty() {
                if old_value.is_some() {
                    self.model.range_clear_contents(&Area { sheet, row, column, width: 1, height: 1 })?;
                    diff_list.push(Diff::RangeClearContents { sheet, row, column, width: 1, height: 1, old_value: vec![vec![old_value]] });
                }
                continue;
            }
            self.model.set_user_input(sheet, row, column, value.clone())?;
            diff_list.push(Diff::SetCellValue { sheet, row, column, new_value: value, old_value: Box::new(old_value) });
        }
        if !diff_list.is_empty() {
            self.push_diff_list(diff_list);
            self.evaluate_if_not_paused();
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    #![allow(clippy::unwrap_used)]
    use super::*;

    fn model() -> UserModel<'static> {
        UserModel::new_empty("model", "ru", "UTC", "ru").unwrap()
    }

    #[test]
    fn ens_sort_moves_rows_with_formulas() {
        let mut m = model();
        for (r, name, n) in [(1, "груша", "3"), (2, "", "1"), (3, "Яблоко", "2"), (4, "арбуз", "5")] {
            m.set_user_input(0, r, 1, name).unwrap();
            m.set_user_input(0, r, 2, n).unwrap();
            m.set_user_input(0, r, 3, &format!("=B{r}*10")).unwrap();
        }
        let area = Area { sheet: 0, row: 1, column: 1, width: 3, height: 4 };
        m.sort_range(&area, 1, false).unwrap();
        let col = |m: &UserModel, c| (1..=4).map(|r| m.get_formatted_cell_value(0, r, c).unwrap()).collect::<Vec<_>>();
        assert_eq!(col(&m, 1), ["арбуз", "груша", "Яблоко", ""]); // blank last
        assert_eq!(col(&m, 3), ["50", "30", "20", "10"]); // formulas follow their rows
        assert_eq!(m.get_cell_content(0, 1, 3).unwrap(), "=B1*10");
        m.sort_range(&area, 2, false).unwrap();
        assert_eq!(col(&m, 2), ["1", "2", "3", "5"]);
        m.undo().unwrap();
        assert_eq!(col(&m, 1), ["арбуз", "груша", "Яблоко", ""]);
    }

    #[test]
    fn ens_paste_values_keeps_precision() {
        let mut m = model();
        m.set_user_input(0, 1, 1, "1,23456789012345").unwrap();
        m.set_user_input(0, 2, 1, "=A1*2").unwrap();
        m.set_user_input(0, 3, 1, "007").unwrap();
        m.update_range_style(&Area { sheet: 0, row: 1, column: 1, width: 1, height: 2 }, "num_fmt", "0,00").unwrap();
        m.set_selected_cell(1, 3).unwrap();
        m.paste_values(0, (1, 1, 3, 1), false).unwrap();
        assert_eq!(m.get_cell_content(0, 2, 3).unwrap(), "2,4691357802469");
        m.set_selected_cell(5, 1).unwrap();
        m.paste_values(0, (1, 1, 2, 1), true).unwrap();
        assert_eq!(m.get_cell_content(0, 5, 2).unwrap(), "2,4691357802469");
        assert_eq!(m.get_cell_content(0, 5, 1).unwrap(), "1,23456789012345");
        m.undo().unwrap();
        assert_eq!(m.get_cell_content(0, 5, 2).unwrap(), "");
    }

    #[test]
    fn ens_row_and_column_outline() {
        let mut m = model();
        m.set_rows_outline(0, 2, 4, 1).unwrap();
        m.set_rows_outline(0, 3, 3, 2).unwrap();
        m.set_row_collapsed(0, 5, true).unwrap();
        assert_eq!(m.get_row_outline(0).unwrap(), [2, 1, 0, 3, 2, 0, 4, 1, 0, 5, 0, 1]);
        m.undo().unwrap();
        assert_eq!(m.get_row_outline(0).unwrap(), [2, 1, 0, 3, 2, 0, 4, 1, 0]);
        m.set_columns_width(0, 2, 6, 120.0).unwrap();
        m.set_columns_outline(0, 3, 4, 1).unwrap();
        assert_eq!(m.get_column_outline(0).unwrap(), [3, 1, 0, 4, 1, 0]);
        assert_eq!(m.get_column_width(0, 5).unwrap(), 120.0);
        assert_eq!(m.get_column_width(0, 3).unwrap(), 120.0);
        m.undo().unwrap();
        assert!(m.get_column_outline(0).unwrap().is_empty());
        // Вставка строки над группой сдвигает её вместе со строками.
        m.insert_rows(0, 1, 1).unwrap();
        assert_eq!(&m.get_row_outline(0).unwrap()[..3], [3, 1, 0]);
    }

    #[test]
    fn ens_text_rotation_style() {
        let mut m = model();
        let area = Area { sheet: 0, row: 1, column: 1, width: 1, height: 1 };
        m.update_range_style(&area, "alignment.text_rotation", "90").unwrap();
        assert_eq!(m.get_cell_style(0, 1, 1).unwrap().alignment.unwrap().text_rotation, 90);
        assert!(m.update_range_style(&area, "alignment.text_rotation", "200").is_err());
        m.update_range_style(&area, "alignment.text_rotation", "255").unwrap();
        assert_eq!(m.get_cell_style(0, 1, 1).unwrap().alignment.unwrap().text_rotation, 255);
    }

    #[test]
    fn ens_cf_range_shrinks_and_undo_restores() {
        use crate::cf_types::CfRuleInput;
        let mut m = model();
        let rule = CfRuleInput::Blanks { format: Default::default(), stop_if_true: false };
        m.add_conditional_formatting(0, "A3:A8", rule).unwrap();
        let range = |m: &UserModel| m.get_conditional_formatting_list(0).unwrap().iter().map(|c| c.range.clone()).collect::<Vec<_>>();
        assert_eq!(range(&m), ["A3:A8"]);
        m.delete_rows(0, 5, 2).unwrap(); // inside: shrinks
        assert_eq!(range(&m), ["A3:A6"]);
        m.delete_rows(0, 1, 3).unwrap(); // edge: A3:A6 loses A3 and moves up
        assert_eq!(range(&m), ["A1:A3"]);
        m.delete_rows(0, 1, 5).unwrap(); // whole range: the rule goes
        assert!(range(&m).is_empty());
        m.undo().unwrap();
        assert_eq!(range(&m), ["A1:A3"]);
        m.undo().unwrap();
        assert_eq!(range(&m), ["A3:A6"]);
    }

    #[test]
    fn ens_cyrillic_defined_name() {
        let mut m = model();
        m.set_user_input(0, 1, 1, "10").unwrap();
        m.new_defined_name("Цена", None, "Sheet1!$A$1").unwrap();
        m.set_user_input(0, 1, 2, "=Цена*2").unwrap();
        assert_eq!(m.get_formatted_cell_value(0, 1, 2).unwrap(), "20");
    }

    #[test]
    fn ens_print_titles() {
        let mut m = model();
        assert_eq!(m.get_print_titles(0).unwrap(), [0; 4]);
        m.set_print_titles(0, [1, 3, 1, 2]).unwrap();
        assert_eq!(m.get_print_titles(0).unwrap(), [1, 3, 1, 2]);
        m.undo().unwrap();
        assert_eq!(m.get_print_titles(0).unwrap(), [0; 4]);
        m.redo().unwrap();
        m.set_print_titles(0, [2, 2, 0, 0]).unwrap();
        assert_eq!(m.get_print_titles(0).unwrap(), [2, 2, 0, 0]);
    }

    #[test]
    fn ens_default_font() {
        let mut m = model();
        m.set_default_font("Times New Roman", 12);
        let font = m.get_cell_style(0, 1, 1).unwrap().font;
        assert_eq!((font.name.as_str(), font.sz), ("Times New Roman", 12));
    }

    #[test]
    fn ens_default_vertical_center() {
        let mut m = model();
        m.set_default_vertical_center();
        m.set_user_input(0, 3, 3, "x").unwrap();
        let st = m.get_cell_style(0, 3, 3).unwrap();
        assert_eq!(st.alignment.map(|a| a.vertical), Some(crate::types::VerticalAlignment::Center));
        // Новое оформление ячейки сохраняет выравнивание по умолчанию.
        let area = Area { sheet: 0, row: 3, column: 3, width: 1, height: 1 };
        m.update_range_style(&area, "font.b", "true").unwrap();
        let st = m.get_cell_style(0, 3, 3).unwrap();
        assert!(st.font.b);
        assert_eq!(st.alignment.map(|a| a.vertical), Some(crate::types::VerticalAlignment::Center));
    }

    #[test]
    fn ens_merge_undo_steps() {
        let mut m = model();
        m.set_user_input(0, 1, 1, "x").unwrap();
        let mark = m.undo_mark();
        m.set_user_input(0, 2, 1, "a").unwrap();
        m.set_user_input(0, 3, 1, "b").unwrap();
        m.merge_undo_since(mark);
        m.undo().unwrap();
        assert_eq!(m.get_formatted_cell_value(0, 2, 1).unwrap(), "");
        assert_eq!(m.get_formatted_cell_value(0, 3, 1).unwrap(), "");
        assert_eq!(m.get_formatted_cell_value(0, 1, 1).unwrap(), "x");
        m.redo().unwrap();
        assert_eq!(m.get_formatted_cell_value(0, 3, 1).unwrap(), "b");
    }

    #[test]
    fn ens_move_rows_and_columns_exact() {
        let mut m = model();
        for r in 1..=6 {
            m.set_user_input(0, r, 1, &format!("{r}")).unwrap();
        }
        m.set_user_input(0, 1, 3, "=SUM(A2:A3)").unwrap();
        m.set_rows_height(0, 2, 2, 40.0).unwrap();
        m.set_rows_hidden(0, 5, 5, true).unwrap();
        // Строки 2–3 (вырезаны) — вставить перед строкой 6: встают на 4–5.
        m.move_rows_exact(0, 2, 2, 2).unwrap();
        let col = |m: &UserModel| (1..=6).map(|r| m.get_formatted_cell_value(0, r, 1).unwrap()).collect::<Vec<_>>();
        assert_eq!(col(&m), ["1", "4", "5", "2", "3", "6"]);
        assert_eq!(m.get_cell_content(0, 1, 3).unwrap(), "=СУММ(A4:A5)");
        assert_eq!(m.get_row_height(0, 4).unwrap(), 40.0);
        assert!(m.model.is_row_hidden(0, 3).unwrap(), "скрытая строка 5 переехала на 3");
        assert!(!m.model.is_row_hidden(0, 5).unwrap());
        m.undo().unwrap();
        assert_eq!(col(&m), ["1", "2", "3", "4", "5", "6"]);
        assert_eq!(m.get_row_height(0, 2).unwrap(), 40.0);
        assert!(m.model.is_row_hidden(0, 5).unwrap());
        // Столбец C — перед A.
        m.set_user_input(0, 1, 2, "b").unwrap();
        m.set_columns_width(0, 3, 3, 200.0).unwrap();
        m.move_columns_exact(0, 3, 1, -2).unwrap();
        assert_eq!(m.get_cell_content(0, 1, 1).unwrap(), "=СУММ(B2:B3)");
        assert_eq!(m.get_formatted_cell_value(0, 1, 3).unwrap(), "b");
        assert_eq!(m.get_column_width(0, 1).unwrap(), 200.0);
    }

    #[test]
    fn ens_sort_blocks_with_merged_cells() {
        let mut m = model();
        // Блоки по две строки: A объединён по вертикали, B — обычный.
        for (r, v) in [(1, "30"), (3, "10"), (5, "20")] {
            m.set_user_input(0, r, 1, v).unwrap();
            m.merge_cells(&Area { sheet: 0, row: r, column: 1, width: 1, height: 2 }).unwrap();
            m.set_user_input(0, r, 2, &format!("a{v}")).unwrap();
            m.set_user_input(0, r + 1, 2, &format!("b{v}")).unwrap();
        }
        let area = Area { sheet: 0, row: 1, column: 1, width: 2, height: 6 };
        m.sort_range_by(&area, &[(1, false)], Blanks::Last).unwrap();
        let col = |c| (1..=6).map(|r| m.get_formatted_cell_value(0, r, c).unwrap()).collect::<Vec<_>>();
        assert_eq!(col(1), ["10", "", "20", "", "30", ""]);
        assert_eq!(col(2), ["a10", "b10", "a20", "b20", "a30", "b30"]);
        // Разные размеры объединений — понятная ошибка.
        m.unmerge_cells(&Area { sheet: 0, row: 3, column: 1, width: 1, height: 2 }).unwrap();
        assert!(m.sort_range_by(&area, &[(1, false)], Blanks::Last).is_err());
    }

    #[test]
    fn ens_sort_visible_rows_only() {
        let mut m = model();
        // Строки 2 и 4 скрыты фильтром: они остаются на местах.
        for (r, v) in [(1, "50"), (2, "1"), (3, "30"), (4, "2"), (5, "10")] {
            m.set_user_input(0, r, 1, v).unwrap();
            m.set_user_input(0, r, 2, &format!("=A{r}*2")).unwrap();
        }
        m.set_rows_hidden(0, 2, 2, true).unwrap();
        m.set_rows_hidden(0, 4, 4, true).unwrap();
        let area = Area { sheet: 0, row: 1, column: 1, width: 2, height: 5 };
        m.sort_range_by_rows(&area, &[(1, false)], Blanks::Last, true).unwrap();
        let col = |m: &UserModel, c| (1..=5).map(|r| m.get_formatted_cell_value(0, r, c).unwrap()).collect::<Vec<_>>();
        assert_eq!(col(&m, 1), ["10", "1", "30", "2", "50"]);
        assert_eq!(col(&m, 2), ["20", "2", "60", "4", "100"]);
        assert_eq!(m.get_cell_content(0, 1, 2).unwrap(), "=A1*2");
        m.undo().unwrap();
        assert_eq!(col(&m, 1), ["50", "1", "30", "2", "10"]);
        // Без skip_hidden — все строки, как раньше.
        m.sort_range_by(&area, &[(1, false)], Blanks::Last).unwrap();
        assert_eq!(col(&m, 1), ["1", "2", "10", "30", "50"]);
    }

    #[test]
    fn ens_sort_by_two_columns_and_blank_modes() {
        let mut m = model();
        // Отдел, Сумма
        for (r, d, s) in [(1, "Б", "5"), (2, "А", "7"), (3, "Б", "1"), (4, "", "3"), (5, "А", "2")] {
            m.set_user_input(0, r, 1, d).unwrap();
            m.set_user_input(0, r, 2, s).unwrap();
        }
        let area = Area { sheet: 0, row: 1, column: 1, width: 2, height: 5 };
        let col = |m: &UserModel, c| (1..=5).map(|r| m.get_formatted_cell_value(0, r, c).unwrap()).collect::<Vec<_>>();
        m.sort_range_by(&area, &[(1, false), (2, true)], Blanks::Last).unwrap();
        assert_eq!(col(&m, 1), ["А", "А", "Б", "Б", ""]);
        assert_eq!(col(&m, 2), ["7", "2", "5", "1", "3"]);
        m.sort_range_by(&area, &[(1, false)], Blanks::First).unwrap();
        assert_eq!(col(&m, 1), ["", "А", "А", "Б", "Б"]);
        // Пустые как ноль: среди чисел, по возрастанию — перед 1.
        m.range_clear_contents(&Area { sheet: 0, row: 3, column: 2, width: 1, height: 1 }).unwrap();
        m.sort_range_by(&area, &[(2, false)], Blanks::AsZero).unwrap();
        assert_eq!(col(&m, 2)[0], "");
        m.sort_range_by(&area, &[(2, true)], Blanks::AsZero).unwrap();
        assert_eq!(col(&m, 2)[4], "");
        assert!(m.get_row_layout(0).unwrap().is_empty());
    }

    #[test]
    fn ens_auto_filter_range() {
        let mut m = model();
        assert_eq!(m.get_auto_filter(0).unwrap(), None);
        m.set_auto_filter(0, Some([1, 1, 6, 3])).unwrap();
        assert_eq!(m.get_auto_filter(0).unwrap(), Some([1, 1, 6, 3]));
        m.set_auto_filter(0, Some([2, 2, 9, 4])).unwrap();
        assert_eq!(m.get_auto_filter(0).unwrap(), Some([2, 2, 9, 4]));
        m.set_auto_filter(0, None).unwrap();
        assert_eq!(m.get_auto_filter(0).unwrap(), None);
        assert_eq!(parse_filter_range("'Мой лист'!$A$1:$C$6"), Some((1, 1, 6, 3)));
    }

    #[test]
    fn ens_find_and_replace() {
        let mut m = model();
        m.new_sheet().unwrap();
        m.set_user_input(0, 1, 1, "Москва").unwrap();
        m.set_user_input(0, 2, 1, "москва-сити").unwrap();
        m.set_user_input(1, 3, 2, "Казань").unwrap();
        m.set_user_input(0, 3, 1, "=ДЛСТР(\"москва\")").unwrap();
        let o = FindOptions::default();
        assert_eq!(m.find_all("москва", o).unwrap().len(), 2);
        assert_eq!(m.find_all("москва", FindOptions { in_formulas: true, ..o }).unwrap().len(), 3);
        assert_eq!(m.find_all("Москва", FindOptions { match_case: true, ..o }).unwrap().len(), 1);
        assert_eq!(m.find_all("москва", FindOptions { whole_cell: true, ..o }).unwrap().len(), 1);
        assert_eq!(m.find_all("каз", o).unwrap(), vec![FoundCell { sheet: 1, row: 3, column: 2 }]);
        let n = m.replace_all("москва", "Тверь", o, None).unwrap();
        assert_eq!(n, 3);
        assert_eq!(m.get_formatted_cell_value(0, 1, 1).unwrap(), "Тверь");
        assert_eq!(m.get_formatted_cell_value(0, 2, 1).unwrap(), "Тверь-сити");
        assert_eq!(m.get_formatted_cell_value(0, 3, 1).unwrap(), "5");
        m.undo().unwrap();
        assert_eq!(m.get_formatted_cell_value(0, 1, 1).unwrap(), "Москва");
        assert_eq!(m.get_formatted_cell_value(0, 3, 1).unwrap(), "6");
    }
}

fn comment_ref(row: i32, column: i32) -> Result<String, String> {
    let col = crate::expressions::utils::number_to_column(column).ok_or_else(|| format!("Invalid column: {column}"))?;
    if row < 1 {
        return Err(format!("Invalid row: {row}"));
    }
    Ok(format!("{col}{row}"))
}
