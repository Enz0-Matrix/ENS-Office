#![allow(clippy::unwrap_used)]

//! ENS-Office: incremental recalculation (incremental.rs). Each test edits a
//! UserModel, checks the values, and checks that only the affected formulas
//! ran. Run the whole suite with `ENS_VERIFY_INCREMENTAL=1` to compare every
//! incremental evaluation with a full one.

use crate::incremental::LastEvaluation;
use crate::UserModel;

fn new_model() -> UserModel<'static> {
    UserModel::new_empty("model", "en", "UTC", "en").unwrap()
}

fn text(model: &UserModel, cell: (i32, i32)) -> String {
    model.get_formatted_cell_value(0, cell.0, cell.1).unwrap()
}

fn ran(model: &UserModel) -> LastEvaluation {
    model.model.last_evaluation()
}

#[test]
fn ens_incremental_runs_only_dependents() {
    let mut model = new_model();
    model.pause_evaluation();
    for row in 1..=100 {
        model.set_user_input(0, row, 1, &row.to_string()).unwrap(); // A
        model.set_user_input(0, row, 2, &format!("=A{row}*2")).unwrap(); // B
    }
    model.set_user_input(0, 1, 3, "=SUM(B1:B100)").unwrap(); // C1
    model.set_user_input(0, 2, 3, "=A1+A2").unwrap(); // C2
    model.resume_evaluation();
    model.evaluate();
    assert_eq!(ran(&model), LastEvaluation::Full);
    assert_eq!(text(&model, (1, 3)), "10100");

    model.set_user_input(0, 50, 1, "1000").unwrap(); // A50: B50 and C1
    assert_eq!(ran(&model), LastEvaluation::Incremental(2));
    assert_eq!(text(&model, (50, 2)), "2000");
    assert_eq!(text(&model, (1, 3)), "12000");
    assert_eq!(text(&model, (2, 3)), "3");

    model.set_user_input(0, 1, 1, "5").unwrap(); // A1: B1, C1, C2
    assert_eq!(ran(&model), LastEvaluation::Incremental(3));
    assert_eq!(text(&model, (2, 3)), "7");
    assert_eq!(text(&model, (1, 3)), "12008");
}

#[test]
fn ens_incremental_empty_cell_inside_range_and_delete() {
    let mut model = new_model();
    model.set_user_input(0, 1, 1, "1").unwrap();
    model.set_user_input(0, 3, 1, "3").unwrap();
    model.set_user_input(0, 1, 2, "=SUM(A1:A5)").unwrap();
    model.set_user_input(0, 2, 2, "=COUNTA(A:A)").unwrap();
    model.set_user_input(0, 3, 2, "=SUMPRODUCT(A1:A5,A1:A5)").unwrap();
    model.evaluate();
    assert_eq!(text(&model, (1, 2)), "4");

    model.set_user_input(0, 2, 1, "10").unwrap(); // A2 was empty
    assert_eq!(text(&model, (1, 2)), "14");
    assert_eq!(text(&model, (2, 2)), "3");
    assert_eq!(text(&model, (3, 2)), "110");

    model.set_user_input(0, 1000, 1, "7").unwrap(); // far down, whole column
    assert_eq!(text(&model, (2, 2)), "4");
    assert_eq!(text(&model, (1, 2)), "14");

    model.range_clear_contents(&crate::expressions::types::Area {
        sheet: 0,
        row: 2,
        column: 1,
        width: 1,
        height: 1,
    })
    .unwrap();
    assert_eq!(text(&model, (1, 2)), "4");
    assert_eq!(text(&model, (2, 2)), "3");
}

#[test]
fn ens_incremental_vlookup_and_new_formula() {
    let mut model = new_model();
    for row in 1..=20 {
        model.set_user_input(0, row, 1, &format!("K{row}")).unwrap();
        model.set_user_input(0, row, 2, &(row * 10).to_string()).unwrap();
    }
    model.set_user_input(0, 1, 4, "=VLOOKUP(\"K5\",A1:B20,2,0)").unwrap();
    model.set_user_input(0, 2, 4, "=VLOOKUP(\"K9\",A1:B20,2,0)").unwrap();
    model.evaluate();
    assert_eq!(text(&model, (1, 4)), "50");

    model.set_user_input(0, 5, 2, "555").unwrap();
    assert_eq!(text(&model, (1, 4)), "555");
    assert_eq!(text(&model, (2, 4)), "90");

    // The key moves: K9 typed into row 3 is found first now.
    model.set_user_input(0, 3, 1, "k9").unwrap();
    assert_eq!(text(&model, (2, 4)), "30");

    // A new formula reading the others.
    model.set_user_input(0, 3, 4, "=D1+D2").unwrap();
    assert_eq!(text(&model, (3, 4)), "585");
    // Changing a formula changes its dependents.
    model.set_user_input(0, 1, 4, "=1").unwrap();
    assert_eq!(text(&model, (3, 4)), "31");
}

#[test]
fn ens_incremental_cycles_undo_and_volatile() {
    let mut model = new_model();
    model.set_user_input(0, 1, 1, "1").unwrap();
    model.set_user_input(0, 1, 2, "=A1+1").unwrap();
    model.set_user_input(0, 1, 3, "=B1+1").unwrap();
    model.evaluate();
    // A cycle appears...
    model.set_user_input(0, 1, 1, "=C1").unwrap();
    assert_eq!(text(&model, (1, 3)), "#CIRC!");
    // ...and goes away.
    model.set_user_input(0, 1, 1, "5").unwrap();
    assert_eq!(text(&model, (1, 3)), "7");

    model.undo().unwrap();
    assert_eq!(text(&model, (1, 3)), "#CIRC!");
    model.redo().unwrap();
    assert_eq!(text(&model, (1, 3)), "7");

    // A volatile formula runs every time, even when nothing it reads changed.
    model.set_user_input(0, 5, 1, "=RAND()").unwrap();
    model.set_user_input(0, 6, 1, "1").unwrap();
    assert_eq!(ran(&model), LastEvaluation::Incremental(1));
}

#[test]
fn ens_incremental_falls_back_to_full() {
    let mut model = new_model();
    model.set_user_input(0, 1, 1, "1").unwrap();
    model.set_user_input(0, 1, 2, "=A1*3").unwrap();
    model.evaluate();
    model.set_user_input(0, 1, 1, "2").unwrap();
    assert_eq!(ran(&model), LastEvaluation::Incremental(1));

    // Structure: a new sheet, a defined name.
    model.new_sheet().unwrap();
    assert_eq!(ran(&model), LastEvaluation::Full);
    model.set_user_input(0, 1, 1, "4").unwrap();
    assert_eq!(text(&model, (1, 2)), "12");

    // Dynamic arrays: always full.
    model.set_user_input(0, 3, 1, "=SEQUENCE(3)").unwrap();
    model.set_user_input(0, 1, 1, "5").unwrap();
    assert_eq!(ran(&model), LastEvaluation::Full);
    assert_eq!(text(&model, (1, 2)), "15");
    assert_eq!(text(&model, (5, 1)), "3");
}

#[test]
fn ens_incremental_hidden_rows_subtotal() {
    let mut model = new_model();
    for row in 1..=5 {
        model.set_user_input(0, row, 1, &row.to_string()).unwrap();
    }
    model.set_user_input(0, 1, 2, "=SUBTOTAL(109,A1:A5)").unwrap();
    model.set_user_input(0, 2, 2, "=SUBTOTAL(9,A1:B1)").unwrap();
    model.evaluate();
    assert_eq!(text(&model, (1, 2)), "15");
    model.set_rows_hidden(0, 2, 2, true).unwrap();
    assert_eq!(text(&model, (1, 2)), "13");
    model.set_user_input(0, 2, 1, "100").unwrap(); // hidden, still read
    assert_eq!(text(&model, (1, 2)), "13");
    model.set_rows_hidden(0, 2, 2, false).unwrap();
    assert_eq!(text(&model, (1, 2)), "113");
}

#[test]
fn ens_incremental_cycle_with_minimal_stack() {
    // Depth 1 forces the unwind/replay path; a cell marked circular there is
    // not run again, so its reads are unknown: it must still be recalculated.
    let mut model = new_model();
    model.model.set_max_evaluation_depth(1);
    model.set_user_input(0, 1, 1, "1").unwrap();
    model.set_user_input(0, 1, 2, "=A1+1").unwrap();
    model.set_user_input(0, 1, 3, "=B1+1").unwrap();
    model.set_user_input(0, 1, 1, "=C1").unwrap();
    assert_eq!(text(&model, (1, 3)), "#CIRC!");
    model.evaluate();
    model.model.evaluate(); // a full evaluation in between
    model.set_user_input(0, 1, 1, "5").unwrap();
    assert_eq!(text(&model, (1, 2)), "6");
    assert_eq!(text(&model, (1, 3)), "7");
}

/// Random edits on a small sheet; after each one the incremental result must
/// equal a full evaluation of the same model.
#[test]
fn ens_incremental_random_edits_with_cycles_match_full() {
    let mut seed: u64 = std::env::var("ENS_SEED")
        .ok()
        .and_then(|s| s.parse().ok())
        .unwrap_or(0x2545_f491_4f6c_dd1d);
    let mut next = move |n: u64| {
        seed ^= seed << 13;
        seed ^= seed >> 7;
        seed ^= seed << 17;
        seed % n
    };
    const ROWS: u64 = 12;
    const COLS: u64 = 6;
    let col = |c: u64| char::from(b'A' + c as u8);
    let mut model = new_model();
    model.evaluate();
    let mut incremental_runs = 0;
    for step in 0..400 {
        let row = 1 + next(ROWS) as i32;
        let column = 1 + next(COLS) as i32;
        let (r1, r2) = (1 + next(ROWS), 1 + next(ROWS));
        let (c1, c2) = (next(COLS), next(COLS));
        let (lo, hi) = (r1.min(r2), r1.max(r2));
        let input = match next(12) {
            0 | 1 | 2 => format!("{}", next(20)),
            3 => format!("K{}", next(5)),
            4 => format!("={}{}+1", col(c1), r1),
            5 => format!("=SUM({}{lo}:{}{hi})", col(c1), col(c1)),
            6 => format!("=VLOOKUP(\"K{}\",A1:F{ROWS},{},0)", next(5), 1 + next(COLS)),
            7 => format!("=IF({}{}>10,\"big\",{}{})", col(c1), r1, col(c2), r2),
            8 => format!("=COUNTIF({}:{},\">5\")", col(c1), col(c1)),
            9 => format!("=INDEX(A1:F{ROWS},MATCH({},B1:B{ROWS},0),{})", next(20), 1 + next(COLS)),
            10 => format!("=SUBTOTAL(9,{}1:{}{hi})", col(c1), col(c2)),
            _ => String::new(),
        };
        if input.is_empty() {
            model
                .range_clear_contents(&crate::expressions::types::Area {
                    sheet: 0,
                    row,
                    column,
                    width: 1,
                    height: 1,
                })
                .unwrap();
        } else {
            model.set_user_input(0, row, column, &input).unwrap();
        }
        if matches!(ran(&model), LastEvaluation::Incremental(_)) {
            incremental_runs += 1;
        } else if std::env::var("ENS_DEBUG").is_ok() {
            println!("full at step {step}: {input:?} valid={}", model.model.deps.valid);
        }
        let snapshot = |model: &UserModel| -> Vec<String> {
            let mut values = Vec::new();
            for r in 1..=ROWS as i32 {
                for c in 1..=COLS as i32 {
                    values.push(model.get_formatted_cell_value(0, r, c).unwrap());
                }
            }
            values
        };
        let after_incremental = snapshot(&model);
        model.model.evaluate();
        let after_full = snapshot(&model);
        if after_incremental != after_full {
            for (i, (a, b)) in after_incremental.iter().zip(after_full.iter()).enumerate() {
                if a != b {
                    let (r, c) = (i as i32 / COLS as i32 + 1, i as i32 % COLS as i32 + 1);
                    let content = model.get_cell_content(0, r, c).unwrap();
                    println!("diff ({r},{c}) {content:?}: incremental={a:?} full={b:?}");
                }
            }
            for r in 1..=ROWS as i32 {
                for c in 1..=COLS as i32 {
                    let content = model.get_cell_content(0, r, c).unwrap();
                    if !content.is_empty() {
                        println!("  ({r},{c}) {content} -> {}", model.get_formatted_cell_value(0, r, c).unwrap());
                    }
                }
            }
        }
        assert_eq!(
            after_incremental, after_full,
            "step {step}: {input:?} at ({row},{column})"
        );
    }
    let _ = incremental_runs;
}

/// Random edits without cycles (a formula reads only rows above it), so
/// that nearly every edit takes the incremental path; after each one the incremental result must
/// equal a full evaluation of the same model.
#[test]
fn ens_incremental_random_acyclic_edits_match_full() {
    let mut seed: u64 = std::env::var("ENS_SEED")
        .ok()
        .and_then(|s| s.parse().ok())
        .unwrap_or(0x2545_f491_4f6c_dd1d);
    let mut next = move |n: u64| {
        seed ^= seed << 13;
        seed ^= seed >> 7;
        seed ^= seed << 17;
        seed % n
    };
    const ROWS: u64 = 12;
    const COLS: u64 = 6;
    let col = |c: u64| char::from(b'A' + c as u8);
    let mut model = new_model();
    model.evaluate();
    let mut incremental_runs = 0;
    for step in 0..400 {
        let row = 2 + next(ROWS - 1) as i32;
        let column = 1 + next(COLS) as i32;
        let above = row as u64 - 1;
        let (r1, r2) = (1 + next(above), 1 + next(above));
        let (c1, c2) = (next(COLS), next(COLS));
        let (lo, hi) = (r1.min(r2), r1.max(r2));
        let input = match next(12) {
            0 | 1 | 2 => format!("{}", next(20)),
            3 => format!("K{}", next(5)),
            4 => format!("={}{}+1", col(c1), r1),
            5 => format!("=SUM({}{lo}:{}{hi})", col(c1), col(c1)),
            6 => format!("=VLOOKUP(\"K{}\",A1:F{above},{},0)", next(5), 1 + next(COLS)),
            7 => format!("=IF({}{}>10,\"big\",{}{})", col(c1), r1, col(c2), r2),
            8 => format!("=COUNTIF({}1:{}{above},\">5\")", col(c1), col(c1)),
            9 => format!("=INDEX(A1:F{above},MATCH({},B1:B{above},0),{})", next(20), 1 + next(COLS)),
            10 => format!("=SUBTOTAL(9,{}1:{}{hi})", col(c1.min(c2)), col(c1.max(c2))),
            _ => String::new(),
        };
        if input.is_empty() {
            model
                .range_clear_contents(&crate::expressions::types::Area {
                    sheet: 0,
                    row,
                    column,
                    width: 1,
                    height: 1,
                })
                .unwrap();
        } else {
            model.set_user_input(0, row, column, &input).unwrap();
        }
        if matches!(ran(&model), LastEvaluation::Incremental(_)) {
            incremental_runs += 1;
        } else if std::env::var("ENS_DEBUG").is_ok() {
            println!("full at step {step}: {input:?} valid={}", model.model.deps.valid);
        }
        let snapshot = |model: &UserModel| -> Vec<String> {
            let mut values = Vec::new();
            for r in 1..=ROWS as i32 {
                for c in 1..=COLS as i32 {
                    values.push(model.get_formatted_cell_value(0, r, c).unwrap());
                }
            }
            values
        };
        let after_incremental = snapshot(&model);
        model.model.evaluate();
        let after_full = snapshot(&model);
        if after_incremental != after_full {
            for (i, (a, b)) in after_incremental.iter().zip(after_full.iter()).enumerate() {
                if a != b {
                    let (r, c) = (i as i32 / COLS as i32 + 1, i as i32 % COLS as i32 + 1);
                    let content = model.get_cell_content(0, r, c).unwrap();
                    println!("diff ({r},{c}) {content:?}: incremental={a:?} full={b:?}");
                }
            }
            for r in 1..=ROWS as i32 {
                for c in 1..=COLS as i32 {
                    let content = model.get_cell_content(0, r, c).unwrap();
                    if !content.is_empty() {
                        println!("  ({r},{c}) {content} -> {}", model.get_formatted_cell_value(0, r, c).unwrap());
                    }
                }
            }
        }
        assert_eq!(
            after_incremental, after_full,
            "step {step}: {input:?} at ({row},{column})"
        );
    }
    // Most edits must have taken the incremental path.
    assert!(incremental_runs > 350, "only {incremental_runs} incremental runs");
}
