//! ENS-Office: undo and redo limited to a sheet ("only this sheet" / "all sheets but this one").
//!
//! The history stays one stack; a scoped undo takes the latest step that belongs to the
//! wanted sheets out of the stack and reverts it. Steps of the other sheets are skipped
//! when they only touch cells. A step that changes the book's structure (rows and columns
//! inserted, moved or deleted, sheets added, renamed or moved, names, theme…) or touches
//! both groups of sheets stops the search: it could shift references of the skipped steps.

use super::{common::UserModel, history::{Diff, DiffList, DiffType, QueueDiffs}};

/// Which steps a scoped undo or redo takes.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum UndoScope {
    /// Steps made on the given sheet only.
    Sheet,
    /// Steps made on any sheet but the given one.
    OtherSheets,
}

/// What a scoped undo or redo did.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum ScopedResult {
    /// The step was undone (redone).
    Done,
    /// There is nothing to undo (redo) in this scope.
    Nothing,
    /// The nearest step changes the book's structure or touches both groups of sheets.
    Blocked,
}

/// The sheet a step works on; `None` for steps that reach beyond one sheet's cells.
fn sheet_of(diff: &Diff) -> Option<u32> {
    match diff {
        Diff::SetCellValue { sheet, .. }
        | Diff::SetArrayValue { sheet, .. }
        | Diff::RangeClearContents { sheet, .. }
        | Diff::RangeClearAll { sheet, .. }
        | Diff::CellClearFormatting { sheet, .. }
        | Diff::SetCellStyle { sheet, .. }
        | Diff::ApplyNamedStyle { sheet, .. }
        | Diff::SetColumnWidth { sheet, .. }
        | Diff::SetColumnHidden { sheet, .. }
        | Diff::SetRowHeight { sheet, .. }
        | Diff::SetRowHidden { sheet, .. }
        | Diff::SetRowOutline { sheet, .. }
        | Diff::SetPrintArea { sheet, .. }
        | Diff::SetPrintTitles { sheet, .. }
        | Diff::SetConditionalFormattingList { sheet, .. }
        | Diff::SetPageSetup { sheet, .. }
        | Diff::SetOutlineSummary { sheet, .. }
        | Diff::SetColumnOutline { sheet, .. }
        | Diff::SetColumnStyle { sheet, .. }
        | Diff::SetRowStyle { sheet, .. }
        | Diff::DeleteColumnStyle { sheet, .. }
        | Diff::DeleteRowStyle { sheet, .. }
        | Diff::SetFrozenRowsCount { sheet, .. }
        | Diff::SetFrozenColumnsCount { sheet, .. }
        | Diff::SetShowGridLines { sheet, .. }
        | Diff::AddConditionalFormatting { sheet, .. }
        | Diff::DeleteConditionalFormatting { sheet, .. }
        | Diff::UpdateConditionalFormatting { sheet, .. }
        | Diff::SetCellLink { sheet, .. }
        | Diff::SwapConditionalFormattingPriority { sheet, .. }
        | Diff::SetMergedCells { sheet, .. } => Some(*sheet),
        _ => None,
    }
}

enum Verdict {
    Take,
    Skip,
    Block,
}

fn classify(list: &DiffList, sheet: u32, scope: UndoScope) -> Verdict {
    let (mut here, mut elsewhere) = (false, false);
    for diff in list {
        match sheet_of(diff) {
            None => return Verdict::Block,
            Some(s) if s == sheet => here = true,
            Some(_) => elsewhere = true,
        }
    }
    match (scope, here, elsewhere) {
        (_, false, false) => Verdict::Skip,
        (_, true, true) => Verdict::Block,
        (UndoScope::Sheet, true, false) | (UndoScope::OtherSheets, false, true) => Verdict::Take,
        _ => Verdict::Skip,
    }
}

/// Index of the step to take, scanning from the newest.
fn find(stack: &[DiffList], sheet: u32, scope: UndoScope) -> Result<usize, ScopedResult> {
    for (i, list) in stack.iter().enumerate().rev() {
        match classify(list, sheet, scope) {
            Verdict::Take => return Ok(i),
            Verdict::Skip => {}
            Verdict::Block => return Err(ScopedResult::Blocked),
        }
    }
    Err(ScopedResult::Nothing)
}

impl UserModel<'_> {
    /// Whether [`Self::undo_scoped`] would do something.
    pub fn can_undo_scoped(&self, sheet: u32, scope: UndoScope) -> bool {
        find(&self.history.undo_stack, sheet, scope).is_ok()
    }

    /// Whether [`Self::redo_scoped`] would do something.
    pub fn can_redo_scoped(&self, sheet: u32, scope: UndoScope) -> bool {
        find(&self.history.redo_stack, sheet, scope).is_ok()
    }

    /// Undoes the latest step in the scope.
    pub fn undo_scoped(&mut self, sheet: u32, scope: UndoScope) -> Result<ScopedResult, String> {
        let i = match find(&self.history.undo_stack, sheet, scope) {
            Ok(i) => i,
            Err(r) => return Ok(r),
        };
        let list = self.history.undo_stack.remove(i);
        self.history.redo_stack.push(list.clone());
        self.apply_undo_diff_list(&list)?;
        self.send_queue.push(QueueDiffs { r#type: DiffType::Undo, list });
        Ok(ScopedResult::Done)
    }

    /// Redoes the latest undone step in the scope.
    pub fn redo_scoped(&mut self, sheet: u32, scope: UndoScope) -> Result<ScopedResult, String> {
        let i = match find(&self.history.redo_stack, sheet, scope) {
            Ok(i) => i,
            Err(r) => return Ok(r),
        };
        let list = self.history.redo_stack.remove(i);
        self.history.undo_stack.push(list.clone());
        self.apply_diff_list(&list)?;
        self.send_queue.push(QueueDiffs { r#type: DiffType::Redo, list });
        Ok(ScopedResult::Done)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn model() -> UserModel<'static> {
        let mut m = UserModel::new_empty("Book", "en", "UTC", "en").unwrap();
        m.new_sheet().unwrap();
        m
    }

    #[test]
    fn scoped_undo_skips_other_sheets() {
        let mut m = model();
        m.set_user_input(0, 1, 1, "a").unwrap();
        m.set_user_input(1, 1, 1, "b").unwrap();
        m.set_user_input(0, 2, 1, "c").unwrap();
        // Only the second sheet: "b" goes, the first sheet keeps "a" and "c".
        assert_eq!(m.undo_scoped(1, UndoScope::Sheet).unwrap(), ScopedResult::Done);
        assert_eq!(m.get_formatted_cell_value(1, 1, 1).unwrap(), "");
        assert_eq!(m.get_formatted_cell_value(0, 2, 1).unwrap(), "c");
        // The next step back is the creation of the sheet: it stops the search.
        assert_eq!(m.undo_scoped(1, UndoScope::Sheet).unwrap(), ScopedResult::Blocked);
        assert_eq!(m.redo_scoped(1, UndoScope::Sheet).unwrap(), ScopedResult::Done);
        assert_eq!(m.get_formatted_cell_value(1, 1, 1).unwrap(), "b");
    }

    #[test]
    fn scoped_undo_stops_at_structure_changes() {
        let mut m = model();
        m.set_user_input(0, 1, 1, "a").unwrap();
        m.insert_rows(1, 1, 1).unwrap();
        assert_eq!(m.undo_scoped(0, UndoScope::Sheet).unwrap(), ScopedResult::Blocked);
        assert_eq!(m.get_formatted_cell_value(0, 1, 1).unwrap(), "a");
    }
}
