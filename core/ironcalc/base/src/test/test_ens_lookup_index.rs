#![allow(clippy::unwrap_used)]

//! ENS-Office: the per-pass index for exact VLOOKUP/HLOOKUP/MATCH must give
//! exactly the results of the linear scan. Every case is checked twice: in a
//! pass (index) and with the formula evaluated again after an edit.

use crate::test::util::new_empty_model;

#[test]
fn ens_lookup_text_case_insensitive_first_match() {
    let mut model = new_empty_model();
    model._set("A1", "apple");
    model._set("B1", "1");
    model._set("A2", "APPLE");
    model._set("B2", "2");
    model._set("A3", "Груша");
    model._set("B3", "3");
    model._set("A4", "ГРУША");
    model._set("B4", "4");
    model._set("C1", "=VLOOKUP(\"Apple\",A1:B4,2,FALSE)");
    model._set("C2", "=VLOOKUP(\"груша\",A1:B4,2,0)");
    model._set("C3", "=MATCH(\"гРуША\",A1:A4,0)");
    model._set("C4", "=VLOOKUP(\"нет\",A1:B4,2,0)");
    model.evaluate();
    assert_eq!(model._get_text("C1"), *"1");
    assert_eq!(model._get_text("C2"), *"3");
    assert_eq!(model._get_text("C3"), *"3");
    assert_eq!(model._get_text("C4"), *"#N/A");
}

#[test]
fn ens_lookup_wildcards_still_work() {
    let mut model = new_empty_model();
    model._set("A1", "К0001");
    model._set("B1", "10");
    model._set("A2", "К*");
    model._set("B2", "20");
    model._set("C1", "=VLOOKUP(\"К*\",A1:B2,2,0)");
    model._set("C2", "=VLOOKUP(\"К~*\",A1:B2,2,0)");
    model._set("C3", "=VLOOKUP(\"?0001\",A1:B2,2,0)");
    model.evaluate();
    assert_eq!(model._get_text("C1"), *"10");
    assert_eq!(model._get_text("C2"), *"20");
    assert_eq!(model._get_text("C3"), *"10");
}

#[test]
fn ens_lookup_numbers_and_types() {
    let mut model = new_empty_model();
    model._set("A1", "'5"); // text "5" must not match the number 5
    model._set("B1", "text");
    model._set("A2", "5");
    model._set("B2", "number");
    model._set("A3", "=0.1+0.2");
    model._set("B3", "small");
    model._set("A4", "=1.1+2.2");
    model._set("B4", "rounded");
    model._set("C1", "=VLOOKUP(5,A1:B4,2,0)");
    model._set("C2", "=VLOOKUP(\"5\",A1:B4,2,0)");
    model._set("C3", "=VLOOKUP(0.3,A1:B4,2,0)"); // < 1: original scan
    model._set("C4", "=VLOOKUP(3.3,A1:B4,2,0)"); // 3.3000000000000003 matches by 15 digits
    model._set("C5", "=MATCH(5,A1:A4,0)");
    model.evaluate();
    assert_eq!(model._get_text("C1"), *"number");
    assert_eq!(model._get_text("C2"), *"text");
    assert_eq!(model._get_text("C3"), *"small");
    assert_eq!(model._get_text("C4"), *"rounded");
    assert_eq!(model._get_text("C5"), *"2");
}

#[test]
fn ens_lookup_range_of_formulas_and_edits() {
    // The keys are formulas; after an edit the index must be rebuilt.
    let mut model = new_empty_model();
    for row in 1..=50 {
        model._set(&format!("A{row}"), &format!("=\"K\"&{row}*D1"));
        model._set(&format!("B{row}"), &format!("{}", row * 10));
    }
    model._set("D1", "1");
    model._set("C1", "=VLOOKUP(\"K7\",A1:B50,2,0)");
    model._set("C2", "=HLOOKUP(\"x\",E1:G2,2,0)");
    model._set("E1", "a");
    model._set("F1", "X");
    model._set("G1", "x");
    model._set("E2", "1");
    model._set("F2", "2");
    model._set("G2", "3");
    model.evaluate();
    assert_eq!(model._get_text("C1"), *"70");
    assert_eq!(model._get_text("C2"), *"2");

    model._set("D1", "7"); // now K7 is in row 1, K49 in row 7
    model.evaluate();
    assert_eq!(model._get_text("C1"), *"10");
}

#[test]
fn ens_lookup_circular_reference() {
    let mut model = new_empty_model();
    model._set("A1", "x");
    model._set("A2", "=VLOOKUP(\"x\",A1:B3,2,0)");
    model._set("B1", "=A2");
    model._set("A3", "y");
    model._set("B3", "=VLOOKUP(\"y\",A1:B3,2,0)");
    model.evaluate();
    assert_eq!(model._get_text("A2"), *"#CIRC!");
}

#[test]
fn ens_lookup_many_rows_is_fast() {
    // 2000 keys, 5000 exact lookups. With the scan this took seconds in a
    // debug build; with the index it must take well under one.
    let mut model = new_empty_model();
    for row in 1..=2000 {
        model._set(&format!("A{row}"), &format!("К{row:04}"));
        model._set(&format!("B{row}"), &format!("{row}"));
    }
    for row in 1..=5000 {
        let key = (row * 7) % 2000 + 1;
        model._set(
            &format!("D{row}"),
            &format!("=VLOOKUP(\"К{key:04}\",$A$1:$B$2000,2,0)"),
        );
    }
    let start = std::time::Instant::now();
    model.evaluate();
    let elapsed = start.elapsed();
    assert_eq!(model._get_text("D1"), *"8");
    assert_eq!(model._get_text("D5000"), *"1001");
    assert!(elapsed.as_secs_f64() < 3.0, "evaluate took {elapsed:?}");
    println!("5000 VLOOKUP over 2000 rows: {elapsed:?}");
}
