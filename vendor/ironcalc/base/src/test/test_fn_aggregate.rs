#![allow(clippy::unwrap_used)]

use crate::test::util::new_empty_model;

/// Microsoft's AGGREGATE example: column A has #DIV/0! and #NUM! among the
/// numbers; the documented results are 96, 72, #VALUE! and 68.
#[test]
fn test_aggregate_documented_example() {
    let mut model = new_empty_model();
    let a = ["=1/0", "72", "30", "=SQRT(-1)", "31", "96", "32", "81", "33", "53", "34"];
    let b = ["82", "65", "95", "63", "53", "71", "55", "83", "100", "91", "89"];
    for (i, (x, y)) in a.iter().zip(b.iter()).enumerate() {
        model._set(&format!("A{}", i + 1), x);
        model._set(&format!("B{}", i + 1), y);
    }
    model._set("C1", "=AGGREGATE(4, 6, A1:A11)");
    model._set("C2", "=AGGREGATE(14, 6, A1:A11, 3)");
    model._set("C3", "=AGGREGATE(15, 6, A1:A11)");
    model._set("C4", "=AGGREGATE(12, 6, A1:A11, B1:B11)");
    model._set("C5", "=MAX(A1:A2)");
    // Errors kept (options 0-1, 4-5): returned, except by COUNT and COUNTA.
    model._set("C6", "=AGGREGATE(9, 4, A1:A3)");
    model._set("C7", "=AGGREGATE(2, 4, A1:A11)");
    model._set("C8", "=AGGREGATE(3, 4, A1:A11)");
    model._set("C9", "=AGGREGATE(3, 6, A1:A11)");
    // Percentiles and quartiles of B (sorted: 53 55 63 65 71 82 83 89 91 95 100).
    model._set("D1", "=AGGREGATE(16, 6, B1:B11, 0.5)");
    model._set("D2", "=AGGREGATE(17, 6, B1:B11, 1)");
    model._set("D3", "=AGGREGATE(18, 6, B1:B11, 0.5)");
    model._set("D4", "=AGGREGATE(19, 6, B1:B11, 1)");
    model._set("D5", "=AGGREGATE(1, 6, B1:B3)");
    model._set("D6", "=AGGREGATE(13, 4, B1:B11, 53)");
    // Invalid arguments.
    model._set("E1", "=AGGREGATE(20, 6, B1:B11)");
    model._set("E2", "=AGGREGATE(9, 8, B1:B11)");
    model._set("E3", "=AGGREGATE(9, 6, B1:B3*2)");
    model._set("E4", "=AGGREGATE(14, 6, B1:B11, 12)");
    model.evaluate();

    assert_eq!(model._get_text("C1"), *"96");
    assert_eq!(model._get_text("C2"), *"72");
    assert_eq!(model._get_text("C3"), *"#VALUE!");
    assert_eq!(model._get_text("C4"), *"68");
    assert_eq!(model._get_text("C5"), *"#DIV/0!");
    assert_eq!(model._get_text("C6"), *"#DIV/0!");
    assert_eq!(model._get_text("C7"), *"9");
    assert_eq!(model._get_text("C8"), *"11");
    assert_eq!(model._get_text("C9"), *"9");
    assert_eq!(model._get_text("D1"), *"82");
    assert_eq!(model._get_text("D2"), *"64");
    assert_eq!(model._get_text("D3"), *"82");
    assert_eq!(model._get_text("D4"), *"63");
    assert_eq!(model._get_text("D5"), *"80.666666667");
    assert_eq!(model._get_text("D6"), *"53");
    assert_eq!(model._get_text("E1"), *"#VALUE!");
    assert_eq!(model._get_text("E2"), *"#VALUE!");
    assert_eq!(model._get_text("E3"), *"#VALUE!");
    assert_eq!(model._get_text("E4"), *"#NUM!");
}

/// The usual array idiom: the k-th row matching a condition, the errors from
/// dividing by FALSE skipped by option 6.
#[test]
fn test_aggregate_kth_match_idiom() {
    let mut model = new_empty_model();
    for (i, v) in ["x", "y", "x", "z", "x", "y"].iter().enumerate() {
        model._set(&format!("A{}", i + 1), v);
    }
    model._set("B1", "=AGGREGATE(15,6,ROW(A1:A6)/(A1:A6=\"x\"),2)");
    model._set("B2", "=AGGREGATE(14,6,ROW(A1:A6)/(A1:A6=\"x\"),1)");
    model._set("B3", "=INDEX(A1:A6,AGGREGATE(15,6,ROW(A1:A6)/(A1:A6=\"z\"),1))");
    model._set("B4", "=AGGREGATE(15,4,ROW(A1:A6)/(A1:A6=\"x\"),1)");
    model.evaluate();
    assert_eq!(model._get_text("B1"), *"3");
    assert_eq!(model._get_text("B2"), *"5");
    assert_eq!(model._get_text("B3"), *"z");
    assert_eq!(model._get_text("B4"), *"#DIV/0!");
}

/// Hidden rows (options 1, 3, 5, 7) and nested SUBTOTAL/AGGREGATE cells
/// (options 0-3).
#[test]
fn test_aggregate_hidden_rows_and_nested() {
    let mut model = new_empty_model();
    for (i, v) in ["1", "2", "3", "4"].iter().enumerate() {
        model._set(&format!("A{}", i + 1), v);
    }
    model.set_row_hidden(0, 2, true).unwrap();
    model._set("B1", "=AGGREGATE(9,5,A1:A4)");
    model._set("B2", "=AGGREGATE(9,4,A1:A4)");
    model._set("B3", "=SUBTOTAL(109,A1:A4)");
    model._set("B4", "=SUBTOTAL(110,A1:A4)");

    model._set("C1", "5");
    model._set("C2", "7");
    model._set("C3", "=AGGREGATE(9,4,C1:C2)");
    model._set("C4", "=SUBTOTAL(9,C1:C2)");
    model._set("D1", "=AGGREGATE(9,0,C1:C4)");
    model._set("D2", "=AGGREGATE(9,4,C1:C4)");
    model._set("D3", "=AGGREGATE(9,4,C:C)");
    model._set("D4", "=_xlfn.AGGREGATE(9,4,C1:C2)");
    model.evaluate();
    assert_eq!(model._get_text("B1"), *"8");
    assert_eq!(model._get_text("B2"), *"10");
    assert_eq!(model._get_text("B3"), *"8");
    // VAR.S of 1, 3, 4 (row 2 hidden).
    assert_eq!(model._get_text("B4"), *"2.333333333");
    assert_eq!(model._get_text("D1"), *"12");
    assert_eq!(model._get_text("D2"), *"36");
    assert_eq!(model._get_text("D3"), *"36");
    assert_eq!(model._get_text("D4"), *"12");
    assert_eq!(model.get_cell_formula(0, 4, 4).unwrap().unwrap(), "=AGGREGATE(9,4,C1:C2)");
}
