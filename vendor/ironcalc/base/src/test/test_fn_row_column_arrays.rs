#![allow(clippy::unwrap_used)]

use crate::test::util::new_empty_model;

/// ROW/COLUMN of a multi-cell reference are arrays in Excel. Ranking idioms
/// encode the row into a sort key and decode it back; with a scalar ROW every
/// key decoded to the same row (seen on a real large model).
#[test]
fn test_row_column_arrays() {
    let mut model = new_empty_model();
    for (i, (name, v)) in [("Acme", 10), ("Initech", 5), ("", 7), ("Acme", 3), ("Globex", 9), ("Initech", 1)].iter().enumerate() {
        let r = i + 1;
        if !name.is_empty() {
            model._set(&format!("A{r}"), name);
        }
        model._set(&format!("B{r}"), &v.to_string());
    }
    model._set("D1", "=SUM(ROW(A1:A6))");
    model._set("D2", "=MAX(ROW(A1:A6))");
    model._set("D3", "=SUM(COLUMN(A1:C1))");
    model._set("D4", "=SMALL(IF(A1:A6<>\"\",ROW(A1:A6)),2)");
    model._set("D5", "=LARGE(IF(A1:A6<>\"\",B1:B6*100000+ROW(A1:A6)),1)");
    model._set("D6", "=SUMPRODUCT(--(A1:A6<>\"\"),ROW(A1:A6))");
    model._set("D7", "=ROW(B3)");
    model._set("D8", "=COLUMN(C3)");
    model._set("D9", "=INDEX(A1:A6,MOD(LARGE(IF(A1:A6<>\"\",B1:B6*100000+ROW(A1:A6)),1),100000))");
    model._set("D10", "=TEXTJOIN(\",\",TRUE,IF(A1:A6=\"Initech\",B1:B6,\"\"))");
    model._set("D11", "=TEXTJOIN(\"-\",FALSE,IF(A1:A6=\"Initech\",B1:B6,\"\"))");
    // Unary minus idioms: `--(cond)` coerces booleans, `-(range)` negates element-wise.
    model._set("E1", "=SUMPRODUCT(--(A1:A6<>\"\"),B1:B6)");
    model._set("E2", "=SUM(--(A1:A6=\"Acme\"))");
    model._set("E3", "=SUM(-(B1:B6))");
    model._set("E4", "=SUMPRODUCT(--(A1:A6<>\"\"),ROW(A1:A6))");
    model._set("E5", "=SUM(-(A1:A6<>\"\"))");
    model.evaluate();
    assert_eq!(model._get_text("E1"), *"28");
    assert_eq!(model._get_text("E2"), *"2");
    assert_eq!(model._get_text("E3"), *"-35");
    assert_eq!(model._get_text("E4"), *"18");
    assert_eq!(model._get_text("E5"), *"-5");
    // Formula text round-trips: the double minus and the parentheses survive.
    assert_eq!(model.get_cell_formula(0, 1, 5).unwrap().unwrap(), "=SUMPRODUCT(--(A1:A6<>\"\"),B1:B6)");
    assert_eq!(model.get_cell_formula(0, 5, 5).unwrap().unwrap(), "=SUM(-(A1:A6<>\"\"))");
    assert_eq!(model._get_text("D1"), *"21");
    assert_eq!(model._get_text("D2"), *"6");
    assert_eq!(model._get_text("D3"), *"6");
    assert_eq!(model._get_text("D4"), *"2");
    assert_eq!(model._get_text("D5"), *"1000001"); // Acme, 10, row 1
    assert_eq!(model._get_text("D6"), *"18"); // rows 1+2+4+5+6
    assert_eq!(model._get_text("D7"), *"3");
    assert_eq!(model._get_text("D8"), *"3");
    assert_eq!(model._get_text("D9"), *"Acme");
    assert_eq!(model._get_text("D10"), *"5,1");
    assert_eq!(model._get_text("D11"), *"-5----1");
}
