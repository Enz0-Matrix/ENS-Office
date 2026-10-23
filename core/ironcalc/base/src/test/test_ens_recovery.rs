#![allow(clippy::unwrap_used)]

//! ENS-Office §7: a snapshot plus the journal of edits made after it gives back
//! the workbook as it was at the moment of a crash (what ENS-Table does).

use crate::UserModel;

#[test]
fn ens_snapshot_plus_journal_restores_the_workbook() {
    let mut m = UserModel::new_empty("model", "ru", "UTC", "ru").unwrap();
    m.set_user_input(0, 1, 1, "10").unwrap();
    m.set_user_input(0, 2, 1, "=A1*2").unwrap();
    m.flush_send_queue();
    let snapshot = m.to_bytes();

    // Edits after the snapshot, flushed to the journal in two batches.
    let mut journal = Vec::new();
    m.set_user_input(0, 1, 1, "7,5").unwrap();
    m.set_user_input(0, 3, 1, "=СУММ(A1:A2)").unwrap();
    journal.push(m.flush_send_queue());
    m.new_sheet().unwrap();
    m.set_user_input(1, 1, 1, "второй лист").unwrap();
    m.undo().unwrap();
    journal.push(m.flush_send_queue());

    let mut restored = UserModel::from_bytes(&snapshot, "ru").unwrap();
    for diffs in &journal {
        restored.apply_external_diffs(diffs).unwrap();
    }
    restored.evaluate();
    for (sheet, row, column) in [(0, 1, 1), (0, 2, 1), (0, 3, 1), (1, 1, 1)] {
        assert_eq!(
            restored.get_formatted_cell_value(sheet, row, column).unwrap(),
            m.get_formatted_cell_value(sheet, row, column).unwrap(),
            "cell {sheet}:{row}:{column}"
        );
    }
    assert_eq!(restored.get_formatted_cell_value(0, 3, 1).unwrap(), "22,5");
    assert_eq!(restored.get_worksheets_properties().len(), 2);
}
