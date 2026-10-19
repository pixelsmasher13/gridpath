#![allow(clippy::unwrap_used)]

use crate::test::util::new_empty_model;

// INDIRECT with a plain A1 string — the baseline behavior.
#[test]
fn indirect_a1_reference() {
    let mut model = new_empty_model();
    model._set("A1", "42");
    model._set("B1", r#"=INDIRECT("A1")"#);
    model.evaluate();
    assert_eq!(model._get_text("B1"), "42");
}

// INDIRECT resolving a workbook-scoped defined name (Excel-compatible).
#[test]
fn indirect_defined_name_cell() {
    let mut model = new_empty_model();
    model._set("A1", "42");
    model
        .new_defined_name("MyCell", None, "Sheet1!$A$1")
        .unwrap();
    model._set("B1", r#"=INDIRECT("MyCell")"#);
    model.evaluate();
    assert_eq!(model._get_text("B1"), "42");
}

// A defined name pointing at a range: functions that take references work
// on the result (this is the analyst-model pattern: ROW(INDIRECT("name"))).
#[test]
fn indirect_defined_name_range() {
    let mut model = new_empty_model();
    model
        .new_defined_name("MyRow", None, "Sheet1!$5:$5")
        .unwrap();
    model._set("B1", r#"=ROW(INDIRECT("MyRow"))"#);
    model._set("B2", r#"=SUM(INDIRECT("MyRow"))"#);
    model._set("A5", "10");
    model._set("C5", "32");
    model.evaluate();
    assert_eq!(model._get_text("B1"), "5");
    assert_eq!(model._get_text("B2"), "42");
}

// Name lookup is case-insensitive, like the rest of Excel.
#[test]
fn indirect_defined_name_case_insensitive() {
    let mut model = new_empty_model();
    model._set("A1", "42");
    model
        .new_defined_name("MyCell", None, "Sheet1!$A$1")
        .unwrap();
    model._set("B1", r#"=INDIRECT("mycell")"#);
    model.evaluate();
    assert_eq!(model._get_text("B1"), "42");
}

// Sheet-scoped names win over workbook-scoped ones from their own sheet.
#[test]
fn indirect_defined_name_sheet_scope_precedence() {
    let mut model = new_empty_model();
    model._set("A1", "1");
    model._set("A2", "2");
    model
        .new_defined_name("Val", None, "Sheet1!$A$1")
        .unwrap();
    model
        .new_defined_name("Val", Some(0), "Sheet1!$A$2")
        .unwrap();
    model._set("B1", r#"=INDIRECT("Val")"#);
    model.evaluate();
    // Sheet-local definition shadows the global one.
    assert_eq!(model._get_text("B1"), "2");
}

// A sheet-qualified string resolves that sheet's local names.
#[test]
fn indirect_defined_name_sheet_qualified() {
    let mut model = new_empty_model();
    model.add_sheet("Data").unwrap();
    model._set("A1", "7"); // Sheet1!A1
    model
        .new_defined_name("Local", Some(1), "Sheet1!$A$1")
        .unwrap();
    model._set("B1", r#"=INDIRECT("Data!Local")"#);
    model.evaluate();
    assert_eq!(model._get_text("B1"), "7");
}

// An unknown name is still a #REF! error, as before.
#[test]
fn indirect_unknown_name_is_ref_error() {
    let mut model = new_empty_model();
    model._set("B1", r#"=INDIRECT("NoSuchName")"#);
    model.evaluate();
    assert_eq!(model._get_text("B1"), "#REF!");
}

// Editing a cell inside the named range recalculates INDIRECT consumers
// (runtime dependency recording).
#[test]
fn indirect_defined_name_dependency_tracking() {
    let mut model = new_empty_model();
    model._set("A1", "1");
    model
        .new_defined_name("MyCell", None, "Sheet1!$A$1")
        .unwrap();
    model._set("B1", r#"=SUM(INDIRECT("MyCell"))"#);
    model.evaluate();
    assert_eq!(model._get_text("B1"), "1");
    model._set("A1", "9");
    model.evaluate();
    assert_eq!(model._get_text("B1"), "9");
}
