//! ENS-Office: protection of a sheet and of the book structure (as in Excel, no password yet).
//! On a protected sheet only unlocked cells can be changed; rows, columns and formatting are fixed.

use crate::expressions::types::Area;

use super::{common::UserModel, history::Diff};

const MAX_CHECKED_CELLS: i64 = 20_000;

impl UserModel<'_> {
    /// Whether the sheet is protected.
    pub fn is_sheet_protected(&self, sheet: u32) -> Result<bool, String> {
        Ok(self.model.workbook.worksheet(sheet)?.ens_protected)
    }

    /// Whether the protection has a password set in another program (this one cannot remove it).
    pub fn is_sheet_password_protected(&self, sheet: u32) -> Result<bool, String> {
        Ok(!self.model.workbook.worksheet(sheet)?.ens_protect_attrs.is_empty())
    }

    pub(crate) fn write_sheet_protected(&mut self, sheet: u32, on: bool) -> Result<(), String> {
        self.model.workbook.worksheet_mut(sheet)?.ens_protected = on;
        Ok(())
    }

    /// Protects the sheet or removes the protection; an undoable step.
    pub fn set_sheet_protected(&mut self, sheet: u32, on: bool) -> Result<(), String> {
        let old_value = self.is_sheet_protected(sheet)?;
        if old_value == on {
            return Ok(());
        }
        if !on && self.is_sheet_password_protected(sheet)? {
            return Err("Защита листа снабжена паролем, снять её здесь нельзя.".to_string());
        }
        self.write_sheet_protected(sheet, on)?;
        self.push_diff_list(vec![Diff::SetSheetProtected { sheet, new_value: on, old_value }]);
        Ok(())
    }

    /// Whether the book structure is locked.
    pub fn is_structure_locked(&self) -> bool {
        self.model.workbook.ens_structure_locked
    }

    pub(crate) fn write_structure_locked(&mut self, on: bool) {
        self.model.workbook.ens_structure_locked = on;
    }

    /// Locks the book structure (adding, deleting, renaming, moving, hiding sheets); an undoable step.
    pub fn set_structure_locked(&mut self, on: bool) -> Result<(), String> {
        let old_value = self.is_structure_locked();
        if old_value == on {
            return Ok(());
        }
        self.write_structure_locked(on);
        self.push_diff_list(vec![Diff::SetStructureLocked { new_value: on, old_value }]);
        Ok(())
    }

    /// Error if the sheet structure is locked.
    pub(crate) fn guard_structure(&self) -> Result<(), String> {
        if self.is_structure_locked() {
            return Err("Структура книги защищена: листы нельзя добавлять, удалять, переименовывать и перемещать.".to_string());
        }
        Ok(())
    }

    /// Error if the sheet is protected (rows, columns, formatting and the like).
    pub(crate) fn guard_sheet(&self, sheet: u32) -> Result<(), String> {
        if self.model.workbook.worksheet(sheet)?.ens_protected {
            return Err("Лист защищён. Снимите защиту листа, чтобы внести изменения.".to_string());
        }
        Ok(())
    }

    /// Whether the cell is locked (Excel's default: yes).
    fn is_cell_locked(&self, sheet: u32, row: i32, column: i32) -> bool {
        self.model
            .get_style_for_cell(sheet, row, column)
            .ok()
            .and_then(|s| s.protection)
            .is_none_or(|p| p.locked)
    }

    /// Error if the sheet is protected and the cells `r1..=r2` × `c1..=c2` are not all unlocked.
    pub(crate) fn guard_cells(&self, sheet: u32, r1: i32, c1: i32, r2: i32, c2: i32) -> Result<(), String> {
        if !self.model.workbook.worksheet(sheet)?.ens_protected {
            return Ok(());
        }
        let count = i64::from(r2 - r1 + 1) * i64::from(c2 - c1 + 1);
        let all_unlocked = count <= MAX_CHECKED_CELLS
            && (r1..=r2).all(|r| (c1..=c2).all(|c| !self.is_cell_locked(sheet, r, c)));
        if all_unlocked {
            return Ok(());
        }
        Err("Ячейка защищена от изменений. Снимите защиту листа или защиту ячейки.".to_string())
    }

    pub(crate) fn guard_area(&self, area: &Area) -> Result<(), String> {
        self.guard_cells(area.sheet, area.row, area.column, area.row + area.height - 1, area.column + area.width - 1)
    }

    /// Guard for operations that work on the selected range of the current view.
    pub(crate) fn guard_selection(&self) -> Result<(), String> {
        let v = self.get_selected_view();
        let [r1, c1, r2, c2] = v.range;
        self.guard_cells(v.sheet, r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2))
    }

    /// Whether the formula of the cell is hidden (the sheet is protected and the cell has the "hidden" flag).
    pub(crate) fn is_formula_hidden(&self, sheet: u32, row: i32, column: i32) -> bool {
        self.model.workbook.worksheet(sheet).is_ok_and(|w| w.ens_protected)
            && self
                .model
                .get_style_for_cell(sheet, row, column)
                .ok()
                .and_then(|s| s.protection)
                .is_some_and(|p| p.hidden)
    }
}
