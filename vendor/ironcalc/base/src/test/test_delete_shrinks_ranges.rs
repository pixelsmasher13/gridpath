#![allow(clippy::unwrap_used)]

use crate::test::util::new_empty_model;

/// Deleting rows or columns that cover an end of a range shrinks the range,
/// as in Excel (`SUM(A1:A3)`, delete row 1 → `SUM(A1:A2)`); only a range
/// deleted entirely becomes #REF!.
#[test]
fn test_delete_rows_shrinks_ranges() {
    let mut model = new_empty_model();
    for (i, v) in ["1", "2", "3", "4"].iter().enumerate() {
        model._set(&format!("A{}", i + 1), v);
    }
    model._set("C11", "=SUM(A1:A3)");
    model._set("C12", "=SUM(A2:A4)");
    model._set("C13", "=SUM($A$1:$A$2)");
    model._set("C14", "=SUM(A1:A1)");
    model._set("C15", "=SUM(A1:A4)");
    model.evaluate();
    // Delete row 1: the first range loses its first row, A1:A1 is gone.
    // The formulas move up a row.
    model.delete_rows(0, 1, 1).unwrap();
    model.evaluate();
    let f = |m: &crate::Model, r: i32, c: i32| m.get_cell_formula(0, r, c).unwrap().unwrap();
    assert_eq!(f(&model, 10, 3), "=SUM(A1:A2)");
    assert_eq!(model._get_text("C10"), *"5");
    assert_eq!(f(&model, 11, 3), "=SUM(A1:A3)");
    assert_eq!(f(&model, 12, 3), "=SUM($A$1:$A$1)");
    assert_eq!(f(&model, 13, 3), "=SUM(#REF!)");
    assert_eq!(f(&model, 14, 3), "=SUM(A1:A3)");
    assert_eq!(model._get_text("C14"), *"9");
}

#[test]
fn test_delete_rows_at_range_end_and_columns() {
    let mut model = new_empty_model();
    for (i, v) in ["1", "2", "3", "4", "5"].iter().enumerate() {
        model._set(&format!("A{}", i + 1), v);
    }
    model._set("C1", "=SUM(A1:A3)");
    model._set("C2", "=AVERAGE(A2:A5)");
    model.evaluate();
    // Delete rows 3-5 (the formulas sit above them): C1 keeps A1:A2, C2 keeps A2:A2.
    model.delete_rows(0, 3, 3).unwrap();
    model.evaluate();
    assert_eq!(model.get_cell_formula(0, 1, 3).unwrap().unwrap(), "=SUM(A1:A2)");
    assert_eq!(model.get_cell_formula(0, 2, 3).unwrap().unwrap(), "=AVERAGE(A2:A2)");
    assert_eq!(model._get_text("C2"), *"2");

    let mut model = new_empty_model();
    model._set("A1", "1");
    model._set("B1", "2");
    model._set("C1", "3");
    model._set("A3", "=SUM(A1:C1)");
    model._set("A4", "=SUM(B1:B1)");
    model.evaluate();
    // Delete column B: A1:C1 → A1:B1, B1:B1 → #REF!.
    model.delete_columns(0, 2, 1).unwrap();
    model.evaluate();
    assert_eq!(model.get_cell_formula(0, 3, 1).unwrap().unwrap(), "=SUM(A1:B1)");
    assert_eq!(model._get_text("A3"), *"4");
    assert_eq!(model.get_cell_formula(0, 4, 1).unwrap().unwrap(), "=SUM(#REF!)");
}
