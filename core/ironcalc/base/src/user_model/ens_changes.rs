//! ENS-Office: what was changed in the book since it was opened (or last saved),
//! for the change log kept in the file ("who changed what and when").
//!
//! Every recorded step (`push_diff_list`) adds the touched area of each sheet;
//! areas are kept small: contained rectangles are dropped, neighbours merged and,
//! when there are too many, the sheet keeps one bounding rectangle.

use crate::constants::{LAST_COLUMN, LAST_ROW};
use crate::expressions::utils::number_to_column;

use super::{common::UserModel, history::Diff};

const MAX_RECTS: usize = 64;

#[derive(Clone, Default)]
pub(crate) struct ChangeTracker {
    /// Sheet name → touched rectangles [r1, c1, r2, c2].
    sheets: Vec<(String, Vec<[i32; 4]>)>,
    /// Structural changes: new, deleted and renamed sheets, deleted rows/columns.
    notes: Vec<String>,
}

fn json_str(s: &str) -> String {
    let mut out = String::from("\"");
    for c in s.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out.push('"');
    out
}

fn contains(a: &[i32; 4], b: &[i32; 4]) -> bool {
    a[0] <= b[0] && a[1] <= b[1] && a[2] >= b[2] && a[3] >= b[3]
}

impl ChangeTracker {
    fn add(&mut self, sheet: &str, rect: [i32; 4]) {
        let i = match self.sheets.iter().position(|(n, _)| n == sheet) {
            Some(i) => i,
            None => {
                self.sheets.push((sheet.to_string(), Vec::new()));
                self.sheets.len() - 1
            }
        };
        let rects = &mut self.sheets[i].1;
        if rects.iter().any(|r| contains(r, &rect)) {
            return;
        }
        rects.retain(|r| !contains(&rect, r));
        let mut rect = rect;
        // Neighbours in the same rows or columns become one rectangle (typing down a column).
        while let Some(j) = rects.iter().position(|r| {
            (r[0] == rect[0] && r[2] == rect[2] && r[1] <= rect[3] + 1 && rect[1] <= r[3] + 1)
                || (r[1] == rect[1] && r[3] == rect[3] && r[0] <= rect[2] + 1 && rect[0] <= r[2] + 1)
        }) {
            let r = rects.remove(j);
            rect = [r[0].min(rect[0]), r[1].min(rect[1]), r[2].max(rect[2]), r[3].max(rect[3])];
        }
        rects.push(rect);
        if rects.len() > MAX_RECTS {
            let b = rects.iter().fold(rects[0], |b, r| [b[0].min(r[0]), b[1].min(r[1]), b[2].max(r[2]), b[3].max(r[3])]);
            *rects = vec![b];
        }
    }

    fn note(&mut self, text: String) {
        if !self.notes.contains(&text) && self.notes.len() < 100 {
            self.notes.push(text);
        }
    }

    fn rename(&mut self, old: &str, new: &str) {
        if let Some(s) = self.sheets.iter_mut().find(|(n, _)| n == old) {
            s.0 = new.to_string();
        }
    }
}

/// "A1:C20", "5:7" (whole rows), "B:D" (whole columns), "A1" (one cell).
fn rect_text(r: &[i32; 4]) -> String {
    let col = |c: i32| number_to_column(c).unwrap_or_default();
    if r[1] == 1 && r[3] >= LAST_COLUMN {
        return if r[0] == r[2] { format!("{0}:{0}", r[0]) } else { format!("{}:{}", r[0], r[2]) };
    }
    if r[0] == 1 && r[2] >= LAST_ROW {
        return format!("{}:{}", col(r[1]), col(r[3]));
    }
    if r[0] == r[2] && r[1] == r[3] {
        return format!("{}{}", col(r[1]), r[0]);
    }
    format!("{}{}:{}{}", col(r[1]), r[0], col(r[3]), r[2])
}

impl UserModel<'_> {
    fn sheet_name(&self, sheet: u32) -> String {
        self.model.workbook.worksheets.get(sheet as usize).map(|w| w.name.clone()).unwrap_or_default()
    }

    /// Notes what a recorded step changed (called after the step is applied).
    pub(crate) fn track_changes(&mut self, diffs: &[Diff]) {
        let rows = |r: i32, n: i32| [r, 1, r + n - 1, LAST_COLUMN];
        let cols = |c: i32, n: i32| [1, c, LAST_ROW, c + n - 1];
        for d in diffs {
            let (sheet, rect) = match d {
                Diff::SetCellValue { sheet, row, column, .. }
                | Diff::CellClearFormatting { sheet, row, column, .. }
                | Diff::SetCellStyle { sheet, row, column, .. }
                | Diff::SetCellLink { sheet, row, column, .. } => (*sheet, [*row, *column, *row, *column]),
                Diff::SetArrayValue { sheet, row, column, width, height, .. }
                | Diff::RangeClearContents { sheet, row, column, width, height, .. }
                | Diff::RangeClearAll { sheet, row, column, width, height, .. } => {
                    (*sheet, [*row, *column, row + height - 1, column + width - 1])
                }
                Diff::SetRowStyle { sheet, row, .. } | Diff::DeleteRowStyle { sheet, row, .. } => (*sheet, rows(*row, 1)),
                Diff::SetColumnStyle { sheet, column, .. } | Diff::DeleteColumnStyle { sheet, column, .. } => {
                    (*sheet, cols(*column, 1))
                }
                Diff::InsertRows { sheet, row, count } => (*sheet, rows(*row, *count)),
                Diff::InsertColumns { sheet, column, count } => (*sheet, cols(*column, *count)),
                Diff::MoveRows { sheet, row, row_count, delta } => {
                    let (a, b) = (*row.min(&(row + delta)), (row + row_count - 1).max(row + row_count - 1 + delta));
                    (*sheet, [a, 1, b, LAST_COLUMN])
                }
                Diff::MoveColumns { sheet, column, column_count, delta } => {
                    let (a, b) = (*column.min(&(column + delta)), (column + column_count - 1).max(column + column_count - 1 + delta));
                    (*sheet, [1, a, LAST_ROW, b])
                }
                Diff::DeleteRows { sheet, row, count, .. } => {
                    let name = self.sheet_name(*sheet);
                    let what = if *count == 1 { format!("строка {row}") } else { format!("строки {row}–{}", row + count - 1) };
                    self.ens_changes.note(format!("«{name}»: удалена {what}").replace("удалена строки", "удалены строки"));
                    continue;
                }
                Diff::DeleteColumns { sheet, column, count, .. } => {
                    let name = self.sheet_name(*sheet);
                    let c = |c: i32| number_to_column(c).unwrap_or_default();
                    let what = if *count == 1 { format!("удалён столбец {}", c(*column)) } else { format!("удалены столбцы {}–{}", c(*column), c(column + count - 1)) };
                    self.ens_changes.note(format!("«{name}»: {what}"));
                    continue;
                }
                Diff::NewSheet { name, .. } => {
                    self.ens_changes.note(format!("добавлен лист «{name}»"));
                    continue;
                }
                Diff::DuplicateSheet { new_index, .. } => {
                    let name = self.sheet_name(*new_index);
                    self.ens_changes.note(format!("добавлен лист «{name}»"));
                    continue;
                }
                Diff::DeleteSheet { old_data, .. } => {
                    self.ens_changes.note(format!("удалён лист «{}»", old_data.name));
                    continue;
                }
                Diff::RenameSheet { old_value, new_value, .. } => {
                    self.ens_changes.rename(old_value, new_value);
                    self.ens_changes.note(format!("лист «{old_value}» переименован в «{new_value}»"));
                    continue;
                }
                _ => continue,
            };
            let name = self.sheet_name(sheet);
            if !name.is_empty() {
                let rect = [rect[0].max(1), rect[1].max(1), rect[2].min(LAST_ROW), rect[3].min(LAST_COLUMN)];
                self.ens_changes.add(&name, rect);
            }
        }
    }

    /// ENS-Office: what changed since the book was opened or last saved, JSON
    /// `{"sheets": [[name, "A1:C20 E5"], …], "notes": ["удалён лист «…»", …]}`.
    pub fn get_ens_changes(&self) -> String {
        let sheets: Vec<String> = self
            .ens_changes
            .sheets
            .iter()
            .filter(|(_, r)| !r.is_empty())
            .map(|(n, r)| format!("[{},{}]", json_str(n), json_str(&r.iter().map(rect_text).collect::<Vec<_>>().join(" "))))
            .collect();
        let notes: Vec<String> = self.ens_changes.notes.iter().map(|n| json_str(n)).collect();
        format!(r#"{{"sheets":[{}],"notes":[{}]}}"#, sheets.join(","), notes.join(","))
    }

    /// ENS-Office: starts counting changes anew (after the book is saved).
    pub fn clear_ens_changes(&mut self) {
        self.ens_changes = ChangeTracker::default();
    }

    /// ENS-Office: a fingerprint of the sheet's contents (values and formulas),
    /// to notice that the file was changed by another program since it was saved here.
    pub fn sheet_content_hash(&self, sheet: u32) -> Result<String, String> {
        let ws = self.model.workbook.worksheet(sheet)?;
        let mut h: u64 = 0xcbf2_9ce4_8422_2325;
        let mut eat = |bytes: &[u8]| {
            for b in bytes {
                h ^= *b as u64;
                h = h.wrapping_mul(0x0100_0000_01b3);
            }
        };
        let mut rows = ws.sheet_data.rows();
        rows.sort_unstable();
        for row in rows {
            let mut columns = ws.sheet_data.columns_in_row(row);
            columns.sort_unstable();
            for column in columns {
                let content = self.model.get_localized_cell_content(sheet, row, column)?;
                if content.is_empty() {
                    continue;
                }
                eat(&row.to_le_bytes());
                eat(&column.to_le_bytes());
                eat(content.as_bytes());
                eat(&[0]);
            }
        }
        Ok(format!("{h:016x}"))
    }

    /// ENS-Office: the change log kept in the file (JSON, see `Workbook::ens_change_log`).
    pub fn get_change_log(&self) -> String {
        self.model.workbook.ens_change_log.clone()
    }

    /// ENS-Office: replaces the change log (before saving).
    pub fn set_change_log(&mut self, json: &str) {
        self.model.workbook.ens_change_log = json.to_string();
    }

    /// ENS-Office: who saved the file last (`lastModifiedBy` in xlsx), e.g. after Excel.
    pub fn get_last_modified_by(&self) -> String {
        self.model.workbook.metadata.last_modified_by.clone()
    }

    /// ENS-Office: the user saving the file (written to `lastModifiedBy`, and to
    /// `creator` for a new book).
    pub fn set_last_modified_by(&mut self, name: &str) {
        let m = &mut self.model.workbook.metadata;
        if m.creator.is_empty() || m.creator == crate::new_empty::IRONCALC_USER {
            m.creator = name.to_string();
        }
        m.last_modified_by = name.to_string();
    }
}

#[cfg(test)]
mod tests {
    use crate::UserModel;

    #[test]
    fn changes_are_collected_as_ranges() {
        let mut m = UserModel::new_empty("Книга1", "ru", "UTC", "ru").unwrap();
        for r in 1..=5 {
            m.set_user_input(0, r, 1, "1").unwrap();
        }
        m.set_user_input(0, 3, 2, "x").unwrap();
        m.set_user_input(0, 10, 4, "y").unwrap();
        m.insert_rows(0, 20, 2).unwrap();
        let json = m.get_ens_changes();
        assert!(json.contains("A1:A5"), "{json}");
        assert!(json.contains("B3"), "{json}");
        assert!(json.contains("D10"), "{json}");
        assert!(json.contains("20:21"), "{json}");
        m.new_sheet().unwrap();
        assert!(m.get_ens_changes().contains("добавлен лист"), "{}", m.get_ens_changes());
        m.clear_ens_changes();
        assert_eq!(m.get_ens_changes(), r#"{"sheets":[],"notes":[]}"#);
        let h1 = m.sheet_content_hash(0).unwrap();
        m.set_user_input(0, 1, 1, "2").unwrap();
        assert_ne!(h1, m.sheet_content_hash(0).unwrap());
    }
}
