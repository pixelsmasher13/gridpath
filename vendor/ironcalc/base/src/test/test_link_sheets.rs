#![allow(clippy::unwrap_used)]

//! GRIDPATH (workspace-links spike): a reference to another workbook,
//! `[1]Model!B23`, is evaluated against a hidden "link sheet" of that name
//! (`Model::add_link_sheet`) instead of staying frozen at its cached value.

use crate::cell::CellValue;
use crate::test::util::new_empty_model;
use crate::types::{FormulaValue, SheetState};
use crate::Model;

/// Rows 5 (period labels) and 23 (revenue) of the other workbook's sheet.
fn fill(model: &mut Model, sheet: u32) {
    for (column, label, revenue) in [(2, "Q4-2026", 100.0), (3, "FY2026", 250.0)] {
        model
            .seed_cell_literal(sheet, 5, column, CellValue::String(label.to_string()))
            .unwrap();
        model
            .seed_cell_literal(sheet, 23, column, CellValue::Number(revenue))
            .unwrap();
    }
}

fn model_with_links<'a>() -> (Model<'a>, u32) {
    let mut model = new_empty_model();
    model._set("B1", "10");
    // Quoted and unquoted spellings, a label-keyed lookup, a range, a mix
    // with a local cell, and the file-name spelling.
    model._set("A1", "='[1]Model'!B23");
    model._set("A2", "=[1]Model!C23");
    model._set(
        "A3",
        "=INDEX('[1]Model'!$23:$23,MATCH(\"FY2026\",'[1]Model'!$5:$5,0))",
    );
    model._set("A4", "=SUM([1]Model!B23:C23)");
    model._set("A5", "=[1]Model!B23*2+B1");
    model._set("A6", "=[NOW.xlsx]Model!B23+'[NOW.xlsx]Model'!C23");
    model.evaluate();
    let link = model.add_link_sheet("[1]Model").unwrap();
    fill(&mut model, link);
    let named = model.add_link_sheet("[NOW.xlsx]Model").unwrap();
    fill(&mut model, named);
    model.evaluate();
    (model, link)
}

#[test]
fn unresolved_link_is_ref_error() {
    let mut model = new_empty_model();
    model._set("A1", "=[1]Model!B23");
    model._set("A2", "='[1]Model'!B23");
    model.evaluate();
    assert_eq!(model._get_text("A1"), "#REF!");
    assert_eq!(model._get_text("A2"), "#REF!");
    // The formula text survives as typed (it used to come back unparseable).
    assert_eq!(model._get_formula("A1"), "=[1]Model!B23");
    assert_eq!(model._get_formula("A2"), "=[1]Model!B23");
}

#[test]
fn links_evaluate_against_the_link_sheet() {
    let (model, _) = model_with_links();
    assert_eq!(model._get_text("A1"), "100");
    assert_eq!(model._get_text("A2"), "250");
    assert_eq!(model._get_text("A3"), "250");
    assert_eq!(model._get_text("A4"), "350");
    assert_eq!(model._get_text("A5"), "210");
    assert_eq!(model._get_text("A6"), "350");
    assert_eq!(model._get_formula("A1"), "=[1]Model!B23");
    assert_eq!(
        model._get_formula("A3"),
        "=INDEX([1]Model!$23:$23,MATCH(\"FY2026\",[1]Model!$5:$5,0))"
    );
}

#[test]
fn link_sheet_is_appended_and_very_hidden() {
    let (model, link) = model_with_links();
    assert_eq!(
        model.workbook.get_worksheet_names(),
        ["Sheet1", "[1]Model", "[NOW.xlsx]Model"]
    );
    assert_eq!(link, 1);
    let state = &model.workbook.worksheet(link).unwrap().state;
    assert!(matches!(state, SheetState::VeryHidden));
}

#[test]
fn a_changed_source_value_recalculates_incrementally() {
    let (mut model, link) = model_with_links();
    // The other workbook's FY2026 revenue moves 250 -> 400.
    model
        .seed_cell_literal(link, 23, 3, CellValue::Number(400.0))
        .unwrap();
    let changed = model.evaluate_edited(&[(link, 23, 3)]);
    assert_eq!(model._get_text("A1"), "100");
    assert_eq!(model._get_text("A2"), "400");
    assert_eq!(model._get_text("A3"), "400");
    assert_eq!(model._get_text("A4"), "500");
    // Only the readers of that cell are reported as changed.
    assert!(changed.contains(&(0, 2, 1)));
    assert!(changed.contains(&(0, 3, 1)));
    assert!(changed.contains(&(0, 4, 1)));
    assert!(!changed.contains(&(0, 1, 1)));
    assert!(!changed.contains(&(0, 5, 1)));
    // And the incremental result is what a full recalculation gives.
    model.evaluate();
    assert_eq!(model._get_text("A3"), "400");
    assert_eq!(model._get_text("A4"), "500");
}

#[test]
fn links_survive_a_reparse_of_the_stored_formulas() {
    let (mut model, _) = model_with_links();
    // new_sheet() reparses every formula from its stored (R1C1) text.
    model.new_sheet();
    model.evaluate();
    assert_eq!(model._get_text("A1"), "100");
    assert_eq!(model._get_text("A3"), "250");
    assert_eq!(model._get_text("A6"), "350");
}

#[test]
fn a_workbook_that_is_not_loaded_keeps_its_cached_values() {
    let mut model = new_empty_model();
    model
        .seed_formula_with_cached_value(0, 1, 1, "=[1]Model!B23", FormulaValue::Number(111.0))
        .unwrap();
    model
        .seed_formula_with_cached_value(0, 2, 1, "=[2]Other!B23", FormulaValue::Number(222.0))
        .unwrap();
    model._set("A3", "=A1+A2");
    model.evaluate();
    // Neither workbook is loaded: both show what the file saved.
    assert_eq!(model._get_text("A1"), "111");
    assert_eq!(model._get_text("A2"), "222");
    assert_eq!(model._get_text("A3"), "333");

    let link = model.add_link_sheet("[1]Model").unwrap();
    fill(&mut model, link);
    model.evaluate();
    // [1] is live now; [2] still shows its cached value.
    assert_eq!(model._get_text("A1"), "100");
    assert_eq!(model._get_text("A2"), "222");
    assert_eq!(model._get_text("A3"), "322");
}

#[test]
fn link_sheet_names_are_validated() {
    let mut model = new_empty_model();
    assert!(model.add_link_sheet("Model").is_err());
    assert!(model.add_link_sheet("[1]").is_err());
    assert!(model.add_link_sheet("[]Model").is_err());
    assert!(model.add_link_sheet("[1]Mo:del").is_err());
    assert!(model.add_link_sheet("[1]Model").is_ok());
    assert!(model.add_link_sheet("[1]model").is_err());
    // The ordinary path still refuses brackets.
    assert!(model.add_sheet("[2]Model").is_err());
}

#[test]
fn brackets_that_are_not_a_workbook_prefix_are_unchanged() {
    let mut model = new_empty_model();
    // No table exists, so these are errors — but the same errors as before
    // the lexer learned the workbook prefix, not a reference to sheet "[Qty]".
    model._set("A1", "=SUM(Table1[Qty])");
    model._set("A2", "=[Qty]");
    model.evaluate();
    assert_eq!(model._get_formula("A1"), "=SUM(Table1[Qty])");
    assert_eq!(model._get_formula("A2"), "=[Qty]");
}

#[test]
fn unloaded_links_keep_their_value_through_a_structural_edit() {
    let mut model = new_empty_model();
    model
        .seed_formula_with_cached_value(0, 1, 1, "=[2]Other!B23", FormulaValue::Number(222.0))
        .unwrap();
    model
        .seed_formula_with_cached_value(
            0,
            2,
            1,
            "='[2]Other Sheet'!B23*2",
            FormulaValue::Number(333.0),
        )
        .unwrap();
    model._set("A3", "=A1+A2");
    model.evaluate();
    assert_eq!(model._get_text("A3"), "555");

    // A structural edit elsewhere used to turn both into #REF!: the saved
    // values were kept by coordinate, and the coordinates were thrown away.
    model.insert_rows(0, 50, 1).unwrap();
    model.new_sheet();
    model.evaluate();
    assert_eq!(model._get_text("A1"), "222");
    assert_eq!(model._get_text("A2"), "333");
    assert_eq!(model._get_text("A3"), "555");
    for row in [1, 2] {
        assert_eq!(
            model.imported_value_reason(0, row, 1),
            Some(crate::ImportedValueReason::ExternalReference)
        );
    }
}

#[test]
fn a_row_inserted_in_the_source_sheet_moves_the_links() {
    let (mut model, link) = model_with_links();
    // The other workbook gained a row above revenue; the workspace mirrors
    // the same edit on the link sheet and the links follow, as they do
    // between two sheets of one workbook.
    model.insert_rows(link, 10, 1).unwrap();
    model.evaluate();
    assert_eq!(model._get_formula("A1"), "=[1]Model!B24");
    assert_eq!(
        model._get_formula("A3"),
        "=INDEX([1]Model!$24:$24,MATCH(\"FY2026\",[1]Model!$5:$5,0))"
    );
    assert_eq!(model._get_text("A1"), "100");
    assert_eq!(model._get_text("A3"), "250");
    assert_eq!(model._get_text("A4"), "350");
}

#[test]
fn link_cells_take_typed_values_errors_and_blanks() {
    use crate::expressions::token::Error;
    use crate::new_empty::LinkValue;
    let (mut model, link) = model_with_links();
    model._set("A7", "=[1]Model!B30");
    model._set("A8", "=IFERROR([1]Model!B31,\"bad\")");
    model._set("A9", "=ISBLANK([1]Model!B23)");
    model.evaluate();

    // An error in the other workbook is an error here, not the text "#DIV/0!".
    model
        .set_link_cell(link, 30, 2, LinkValue::Text("007".to_string()))
        .unwrap();
    model
        .set_link_cell(link, 31, 2, LinkValue::Error(Error::DIV))
        .unwrap();
    model.evaluate_edited(&[(link, 30, 2), (link, 31, 2)]);
    assert_eq!(model._get_text("A7"), "007");
    assert_eq!(model._get_text("A8"), "bad");
    assert_eq!(
        model.link_value(link, 31, 2),
        LinkValue::Error(Error::DIV)
    );
    assert_eq!(
        model.link_value(link, 30, 2),
        LinkValue::Text("007".to_string())
    );

    // A cell the other workbook cleared is gone here too.
    assert_eq!(model._get_text("A9"), "FALSE");
    model.set_link_cell(link, 23, 2, LinkValue::Empty).unwrap();
    model.evaluate_edited(&[(link, 23, 2)]);
    assert_eq!(model.link_value(link, 23, 2), LinkValue::Empty);
    assert_eq!(model._get_text("A9"), "TRUE");
    assert_eq!(model._get_text("A1"), "0");

    // Ordinary sheets do not take link cells.
    assert!(model.set_link_cell(0, 1, 1, LinkValue::Number(1.0)).is_err());
    // What a formula cell holds is read typed as well.
    assert_eq!(model.link_value(0, 4, 1), LinkValue::Number(250.0));
    assert_eq!(model.link_value(0, 200, 1), LinkValue::Empty);
}

#[test]
fn a_renamed_source_sheet_renames_the_link_sheet_and_the_links() {
    let (mut model, _) = model_with_links();
    let named = model
        .workbook
        .get_worksheet_names()
        .iter()
        .position(|n| n == "[NOW.xlsx]Model")
        .unwrap() as u32;
    model
        .rename_link_sheet(named, "[NOW.xlsx]Income Statement")
        .unwrap();
    model.evaluate();
    assert_eq!(
        model._get_formula("A6"),
        "='[NOW.xlsx]Income Statement'!B23+'[NOW.xlsx]Income Statement'!C23"
    );
    assert_eq!(model._get_text("A6"), "350");
    // Only link sheets, only link names, and no name twice.
    assert!(model.rename_link_sheet(0, "[NOW.xlsx]Other").is_err());
    assert!(model.rename_link_sheet(named, "Plain").is_err());
    assert!(model.rename_link_sheet(named, "[1]Model").is_err());
    // The ordinary rename still refuses a link name.
    assert!(model.rename_sheet_by_index(0, "[NOW.xlsx]X").is_err());
}

#[test]
fn a_rename_leaves_ranges_on_sheets_that_are_not_here_alone() {
    // A range on a workbook that is not loaded, next to one that is. Renaming
    // any sheet used to give the unloaded range the new name too.
    let mut model = new_empty_model();
    model.add_sheet("Data").unwrap();
    model._set("A1", "=SUM([2]Other!A1:B2)+[2]Other!C3+SUM(Data!A1:A2)");
    model._set("A2", "=SUM([1]Model!B23:C23)+SUM('[2]Other Sheet'!A1:B2)");
    model.evaluate();
    let link = model.add_link_sheet("[1]Model").unwrap();
    fill(&mut model, link);
    model.evaluate();

    // An ordinary sheet.
    model.rename_sheet_by_index(1, "Inputs").unwrap();
    assert_eq!(
        model._get_formula("A1"),
        "=SUM([2]Other!A1:B2)+[2]Other!C3+SUM(Inputs!A1:A2)"
    );
    assert_eq!(
        model._get_formula("A2"),
        "=SUM([1]Model!B23:C23)+SUM('[2]Other Sheet'!A1:B2)"
    );

    // A link sheet.
    model.rename_link_sheet(link, "[NOW.xlsx]Model").unwrap();
    model.evaluate();
    assert_eq!(
        model._get_formula("A1"),
        "=SUM([2]Other!A1:B2)+[2]Other!C3+SUM(Inputs!A1:A2)"
    );
    assert_eq!(
        model._get_formula("A2"),
        "=SUM([NOW.xlsx]Model!B23:C23)+SUM('[2]Other Sheet'!A1:B2)"
    );
}

#[test]
fn a_deleted_link_sheet_leaves_ref_errors() {
    let (mut model, link) = model_with_links();
    model.delete_sheet(link).unwrap();
    model.evaluate();
    assert_eq!(model._get_text("A1"), "#REF!");
    // The other workbook's links are untouched.
    assert_eq!(model._get_text("A6"), "350");
}
