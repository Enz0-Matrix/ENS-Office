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

    pub(crate) fn write_sheet_protected(&mut self, sheet: u32, state: (bool, String)) -> Result<(), String> {
        let ws = self.model.workbook.worksheet_mut(sheet)?;
        (ws.ens_protected, ws.ens_protect_attrs) = state;
        Ok(())
    }

    /// Protects the sheet; with a non-empty `password` the removal of the protection asks for it
    /// (`salt` — random bytes from the caller); an undoable step.
    pub fn protect_sheet(&mut self, sheet: u32, password: &str, salt: &[u8]) -> Result<(), String> {
        let ws = self.model.workbook.worksheet(sheet)?;
        let old_value = (ws.ens_protected, ws.ens_protect_attrs.clone());
        let attrs = if password.is_empty() { String::new() } else { password_attrs(password, salt) };
        let new_value = (true, attrs);
        if old_value == new_value {
            return Ok(());
        }
        self.write_sheet_protected(sheet, new_value.clone())?;
        self.push_diff_list(vec![Diff::SetSheetProtected { sheet, new_value, old_value }]);
        Ok(())
    }

    /// Removes the protection of the sheet; asks for the password if there is one; an undoable step.
    pub fn unprotect_sheet(&mut self, sheet: u32, password: &str) -> Result<(), String> {
        let ws = self.model.workbook.worksheet(sheet)?;
        let old_value = (ws.ens_protected, ws.ens_protect_attrs.clone());
        if !old_value.0 {
            return Ok(());
        }
        check_password(&old_value.1, password)?;
        let new_value = (false, String::new());
        self.write_sheet_protected(sheet, new_value.clone())?;
        self.push_diff_list(vec![Diff::SetSheetProtected { sheet, new_value, old_value }]);
        Ok(())
    }

    /// Protects the sheet without a password or removes a protection that has none.
    pub fn set_sheet_protected(&mut self, sheet: u32, on: bool) -> Result<(), String> {
        if on {
            self.protect_sheet(sheet, "", &[])
        } else {
            self.unprotect_sheet(sheet, "")
        }
    }

    /// Whether the book structure is locked.
    pub fn is_structure_locked(&self) -> bool {
        self.model.workbook.ens_structure_locked
    }

    /// Whether the lock of the structure has a password.
    pub fn is_structure_password_protected(&self) -> bool {
        !self.model.workbook.ens_structure_attrs.is_empty()
    }

    pub(crate) fn write_structure_locked(&mut self, state: (bool, String)) {
        (self.model.workbook.ens_structure_locked, self.model.workbook.ens_structure_attrs) = state;
    }

    /// Locks the book structure (adding, deleting, renaming, moving, hiding sheets), with a password if given; an undoable step.
    pub fn lock_structure(&mut self, password: &str, salt: &[u8]) -> Result<(), String> {
        let old_value = (self.model.workbook.ens_structure_locked, self.model.workbook.ens_structure_attrs.clone());
        let new_value = (true, if password.is_empty() { String::new() } else { password_attrs(password, salt) });
        if old_value == new_value {
            return Ok(());
        }
        self.write_structure_locked(new_value.clone());
        self.push_diff_list(vec![Diff::SetStructureLocked { new_value, old_value }]);
        Ok(())
    }

    /// Unlocks the book structure; asks for the password if there is one; an undoable step.
    pub fn unlock_structure(&mut self, password: &str) -> Result<(), String> {
        let old_value = (self.model.workbook.ens_structure_locked, self.model.workbook.ens_structure_attrs.clone());
        if !old_value.0 {
            return Ok(());
        }
        check_password(&old_value.1, password)?;
        let new_value = (false, String::new());
        self.write_structure_locked(new_value.clone());
        self.push_diff_list(vec![Diff::SetStructureLocked { new_value, old_value }]);
        Ok(())
    }

    /// Locks or unlocks the structure without a password.
    pub fn set_structure_locked(&mut self, on: bool) -> Result<(), String> {
        if on {
            self.lock_structure("", &[])
        } else {
            self.unlock_structure("")
        }
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
        // An invalid sheet is reported by the operation itself.
        if self.model.workbook.worksheet(sheet).is_ok_and(|w| w.ens_protected) {
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
        if !self.model.workbook.worksheet(sheet).is_ok_and(|w| w.ens_protected) {
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

const SPIN_COUNT: u32 = 100_000;
const B64: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

fn base64_encode(data: &[u8]) -> String {
    let mut out = String::new();
    for chunk in data.chunks(3) {
        let n = (u32::from(chunk[0]) << 16) | (u32::from(*chunk.get(1).unwrap_or(&0)) << 8) | u32::from(*chunk.get(2).unwrap_or(&0));
        for i in 0..4 {
            if i <= chunk.len() {
                out.push(B64[((n >> (18 - 6 * i)) & 63) as usize] as char);
            } else {
                out.push('=');
            }
        }
    }
    out
}

fn base64_decode(text: &str) -> Option<Vec<u8>> {
    let mut out = Vec::new();
    let (mut acc, mut bits) = (0u32, 0);
    for c in text.bytes().filter(|c| *c != b'=') {
        acc = (acc << 6) | B64.iter().position(|x| *x == c)? as u32;
        bits += 6;
        if bits >= 8 {
            bits -= 8;
            out.push((acc >> bits) as u8);
            acc &= (1 << bits) - 1;
        }
    }
    Some(out)
}

/// ECMA-376 (agile) hash of a password: SHA-512 of salt + UTF-16LE password, then `spin` rounds.
fn agile_hash(password: &str, salt: &[u8], spin: u32) -> Vec<u8> {
    use sha2::{Digest, Sha512};
    let mut data: Vec<u8> = salt.to_vec();
    data.extend(password.encode_utf16().flat_map(u16::to_le_bytes));
    let mut h = Sha512::digest(&data).to_vec();
    for i in 0..spin {
        let mut d = h.clone();
        d.extend(i.to_le_bytes());
        h = Sha512::digest(&d).to_vec();
    }
    h
}

/// The old 16-bit hash of Excel's `password="XXXX"` attribute.
fn legacy_hash(password: &str) -> u16 {
    let bytes: Vec<u8> = password.chars().map(|c| if (c as u32) < 256 { c as u8 } else { b'?' }).collect();
    let mut hash: u16 = 0;
    for b in bytes.iter().rev() {
        hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7fff);
        hash ^= u16::from(*b);
    }
    hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7fff);
    hash ^ bytes.len() as u16 ^ 0xCE4B
}

/// ` algorithmName="SHA-512" hashValue=… saltValue=… spinCount=…` for a new password.
fn password_attrs(password: &str, salt: &[u8]) -> String {
    let salt: Vec<u8> = if salt.len() >= 8 { salt.to_vec() } else { b"ENS-Office-salt".to_vec() };
    format!(
        " algorithmName=\"SHA-512\" hashValue=\"{}\" saltValue=\"{}\" spinCount=\"{SPIN_COUNT}\"",
        base64_encode(&agile_hash(password, &salt, SPIN_COUNT)),
        base64_encode(&salt)
    )
}

fn attr<'a>(attrs: &'a str, name: &str) -> Option<&'a str> {
    let key = format!(" {name}=\"");
    let start = attrs.find(&key)? + key.len();
    attrs[start..].split('"').next()
}

/// Ok if `attrs` holds no password or `password` fits it.
fn check_password(attrs: &str, password: &str) -> Result<(), String> {
    let wrong = || "Неверный пароль.".to_string();
    if attrs.is_empty() {
        return Ok(());
    }
    if let Some(hash) = attr(attrs, "hashValue") {
        if attr(attrs, "algorithmName") != Some("SHA-512") {
            return Err("Пароль защиты задан алгоритмом, который не поддерживается.".to_string());
        }
        let salt = base64_decode(attr(attrs, "saltValue").unwrap_or("")).ok_or_else(wrong)?;
        let spin: u32 = attr(attrs, "spinCount").and_then(|s| s.parse().ok()).unwrap_or(SPIN_COUNT);
        return if base64_encode(&agile_hash(password, &salt, spin)) == hash { Ok(()) } else { Err(wrong()) };
    }
    match attr(attrs, "password") {
        Some(p) if u16::from_str_radix(p, 16) == Ok(legacy_hash(password)) => Ok(()),
        Some(_) => Err(wrong()),
        None => Ok(()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn empty_password_has_excels_hash() {
        assert_eq!(legacy_hash(""), 0xCE4B);
    }

    #[test]
    fn base64_round_trip() {
        for data in [&b""[..], b"a", b"ab", b"abc", b"abcd", &[0, 255, 128, 7, 9]] {
            assert_eq!(base64_decode(&base64_encode(data)).unwrap(), data);
        }
        assert_eq!(base64_encode(b"Man"), "TWFu");
    }
}
