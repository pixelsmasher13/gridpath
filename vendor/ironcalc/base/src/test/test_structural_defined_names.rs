#![allow(clippy::unwrap_used)]

//! GRIDPATH: a defined name points at cells, so it has to follow them when
//! rows or columns are inserted, deleted or moved — like a reference in a
//! formula does. Upstream left the names where they were: after one row was
//! inserted above `Rate = Sheet1!$B$5`, every formula using `Rate` read
//! whatever had moved into B5.

use crate::test::util::new_empty_model;
use crate::types::DefinedName;
use crate::Model;

/// What the name refers to, as the name manager shows it.
fn refers_to(model: &Model, name: &str) -> String {
    model
        .get_defined_name_list()
        .into_iter()
        .find(|(n, _, _)| n == name)
        .map(|(_, _, formula)| formula)
        .unwrap()
}

#[test]
fn a_named_cell_follows_its_cell_through_row_and_column_edits() {
    let mut model = new_empty_model();
    model._set("B5", "10");
    model._set("B6", "999");
    model._set("C1", "=Rate*2");
    model.new_defined_name("Rate", None, "Sheet1!$B$5").unwrap();
    model.evaluate();
    assert_eq!(model._get_text("C1"), "20");

    model.insert_rows(0, 2, 1).unwrap();
    model.evaluate();
    assert_eq!(refers_to(&model, "Rate"), "Sheet1!$B$6");
    assert_eq!(model._get_text("C1"), "20");

    model.insert_columns(0, 1, 2).unwrap();
    model.evaluate();
    assert_eq!(refers_to(&model, "Rate"), "Sheet1!$D$6");
    assert_eq!(model._get_text("E1"), "20");

    model.delete_columns(0, 1, 2).unwrap();
    model.evaluate();
    assert_eq!(refers_to(&model, "Rate"), "Sheet1!$B$6");
    assert_eq!(model._get_text("C1"), "20");

    model.delete_rows(0, 2, 1).unwrap();
    model.evaluate();
    assert_eq!(refers_to(&model, "Rate"), "Sheet1!$B$5");
    assert_eq!(model._get_text("C1"), "20");

    // An edit below and to the right of the named cell leaves the name alone.
    model.insert_rows(0, 6, 3).unwrap();
    model.insert_columns(0, 3, 1).unwrap();
    model.evaluate();
    assert_eq!(refers_to(&model, "Rate"), "Sheet1!$B$5");
    assert_eq!(model._get_text("D1"), "20");
}

#[test]
fn a_named_range_grows_and_shrinks_with_its_rows() {
    let mut model = new_empty_model();
    model._set("B5", "1");
    model._set("B6", "2");
    model._set("B7", "3");
    model._set("D1", "=SUM(Rates)");
    model
        .new_defined_name("Rates", None, "Sheet1!$B$5:$B$7")
        .unwrap();
    model.evaluate();
    assert_eq!(model._get_text("D1"), "6");

    // A row inside the range: the range grows, and the new cell counts.
    model.insert_rows(0, 6, 1).unwrap();
    model.evaluate();
    assert_eq!(refers_to(&model, "Rates"), "Sheet1!$B$5:$B$8");
    assert_eq!(model._get_text("D1"), "6");
    model._set("B6", "10");
    model.evaluate();
    assert_eq!(model._get_text("D1"), "16");

    model.delete_rows(0, 6, 1).unwrap();
    model.evaluate();
    assert_eq!(refers_to(&model, "Rates"), "Sheet1!$B$5:$B$7");
    assert_eq!(model._get_text("D1"), "6");

    // Rows above the range: it moves as a whole.
    model.insert_rows(0, 1, 4).unwrap();
    model.evaluate();
    assert_eq!(refers_to(&model, "Rates"), "Sheet1!$B$9:$B$11");
    assert_eq!(model._get_text("D5"), "6");
}

#[test]
fn a_name_follows_a_moved_row_and_column() {
    let mut model = new_empty_model();
    model._set("B5", "10");
    model._set("B6", "999");
    model._set("A1", "=Rate*2");
    model.new_defined_name("Rate", None, "Sheet1!$B$5").unwrap();
    model.evaluate();

    // Row 5 goes down past rows 6 and 7.
    model.move_rows_action(0, 5, 1, 2).unwrap();
    model.evaluate();
    assert_eq!(refers_to(&model, "Rate"), "Sheet1!$B$7");
    assert_eq!(model._get_text("A1"), "20");

    // Column B goes right past column C.
    model.move_columns_action(0, 2, 1, 1).unwrap();
    model.evaluate();
    assert_eq!(refers_to(&model, "Rate"), "Sheet1!$C$7");
    assert_eq!(model._get_text("A1"), "20");
}

#[test]
fn names_on_other_sheets_and_names_of_one_sheet_only() {
    let mut model = new_empty_model();
    model.new_sheet();
    model._set("Sheet2!B5", "7");
    model._set("B5", "10");
    model._set("A1", "=Elsewhere+Local");
    model
        .new_defined_name("Elsewhere", None, "Sheet2!$B$5")
        .unwrap();
    // A name that exists on the first sheet only.
    model
        .new_defined_name("Local", Some(0), "Sheet1!$B$5")
        .unwrap();
    model.evaluate();
    assert_eq!(model._get_text("A1"), "17");

    // A row inserted on the first sheet moves the name that points there and
    // not the one that points at the second sheet.
    model.insert_rows(0, 3, 1).unwrap();
    model.evaluate();
    assert_eq!(refers_to(&model, "Elsewhere"), "Sheet2!$B$5");
    assert_eq!(refers_to(&model, "Local"), "Sheet1!$B$6");
    assert_eq!(model._get_text("A1"), "17");

    model.insert_rows(1, 1, 2).unwrap();
    model.evaluate();
    assert_eq!(refers_to(&model, "Elsewhere"), "Sheet2!$B$7");
    assert_eq!(refers_to(&model, "Local"), "Sheet1!$B$6");
    assert_eq!(model._get_text("A1"), "17");
}

#[test]
fn deleting_the_named_cell_leaves_an_error_not_the_next_cell() {
    let mut model = new_empty_model();
    model._set("B5", "10");
    model._set("B6", "999");
    model._set("A1", "=Rate*2");
    model.new_defined_name("Rate", None, "Sheet1!$B$5").unwrap();
    model.evaluate();

    model.delete_rows(0, 5, 1).unwrap();
    model.evaluate();
    assert_eq!(refers_to(&model, "Rate"), "#REF!");
    assert_eq!(model._get_text("A1"), "#REF!");

    // The way Excel writes such a name in a file.
    model.workbook.defined_names.push(DefinedName {
        name: "Gone".to_string(),
        formula: "Sheet1!#REF!".to_string(),
        sheet_id: None,
    });
    model.reset_parsed_structures();
    model._set("A2", "=Gone+1");
    model.evaluate();
    assert_eq!(model._get_text("A2"), "#REF!");
}

#[test]
fn a_named_function_is_left_as_it_is() {
    let mut model = new_empty_model();
    model._set("B5", "4");
    model._set("A1", "=Double(B5)");
    model
        .new_defined_name("Double", None, "=LAMBDA(x, x * 2)")
        .unwrap();
    model.evaluate();
    assert_eq!(model._get_text("A1"), "8");
    let before = refers_to(&model, "Double");
    let formula = model._get_formula("A1");

    model.insert_rows(0, 2, 1).unwrap();
    model.evaluate();
    assert_eq!(refers_to(&model, "Double"), before);
    assert_eq!(model._get_formula("A1"), formula.replace("B5", "B6"));
    assert_eq!(model._get_text("A1"), "8");
}
