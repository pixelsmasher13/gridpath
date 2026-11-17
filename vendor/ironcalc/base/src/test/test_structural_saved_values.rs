#![allow(clippy::unwrap_used)]

//! GRIDPATH: inserting or deleting rows, columns and sheets MOVES cells. A
//! moved cell must hold exactly what it held before: the value a file saved
//! for a formula the engine cannot calculate (add-in calls, links to other
//! workbooks), text that only looks like a formula or a number, every digit
//! of a number, and whether a formula is a plain one or an array formula.
//!
//! Moves used to re-enter each cell as if a user had typed its text, which
//! lost all four.

use crate::cell::CellValue;
use crate::test::util::new_empty_model;
use crate::types::{ArrayKind, Cell, FormulaValue};
use crate::{ImportedValueReason, Model};

fn kind(model: &Model, cell: &str) -> &'static str {
    let reference = model._parse_reference(cell);
    let found = model
        .workbook
        .worksheet(reference.sheet)
        .unwrap()
        .cell(reference.row, reference.column);
    match found {
        None => "none",
        Some(Cell::EmptyCell { .. }) => "empty",
        Some(Cell::BooleanCell { .. }) => "boolean",
        Some(Cell::NumberCell { .. }) => "number",
        Some(Cell::ErrorCell { .. }) => "error",
        Some(Cell::SharedString { .. }) => "text",
        Some(Cell::CellFormula { .. }) => "formula",
        Some(Cell::ArrayFormula {
            kind: ArrayKind::Cse,
            ..
        }) => "cse",
        Some(Cell::ArrayFormula {
            kind: ArrayKind::Dynamic,
            ..
        }) => "dynamic",
        Some(Cell::SpillCell { .. }) => "spill",
    }
}

fn value(model: &Model, cell: &str) -> CellValue {
    let reference = model._parse_reference(cell);
    model
        .get_cell_value_by_index(reference.sheet, reference.row, reference.column)
        .unwrap()
}

fn number(model: &Model, cell: &str) -> f64 {
    match value(model, cell) {
        CellValue::Number(n) => n,
        other => panic!("{cell} is not a number: {other:?}"),
    }
}

fn style(model: &Model, cell: &str) -> i32 {
    let reference = model._parse_reference(cell);
    model
        .get_cell_style_index(reference.sheet, reference.row, reference.column)
        .unwrap()
}

fn reason(model: &Model, cell: &str) -> Option<ImportedValueReason> {
    let reference = model._parse_reference(cell);
    model.imported_value_reason(reference.sheet, reference.row, reference.column)
}

/// A formula cell as the xlsx import builds it from a plain `<f>`: an
/// ordinary (non-array) formula, whatever static analysis makes of its text,
/// carrying the value the file saved.
fn seed_plain(model: &mut Model, sheet: u32, row: i32, column: i32, formula: &str, saved: f64) {
    model
        .seed_formula_with_cached_value(sheet, row, column, formula, FormulaValue::Number(saved))
        .unwrap();
    let worksheet = model.workbook.worksheet_mut(sheet).unwrap();
    let cell = worksheet.cell(row, column).unwrap().clone();
    if let Cell::ArrayFormula { f, s, v, .. } = cell {
        worksheet
            .update_cell(row, column, Cell::CellFormula { f, s, v })
            .unwrap();
    }
}

/// Rows 5 and 6: an add-in call in A, a link in C (to another workbook in
/// row 5, to a sheet that does not exist in row 6), and their sum in D.
fn model_with_saved_values<'a>() -> Model<'a> {
    let mut model = new_empty_model();
    model._set("B5", "EIGI");
    model._set("B6", "WEB");
    seed_plain(&mut model, 0, 5, 1, "=_xll.FDS(B5,\"EV_SALES\")", 2.549606);
    seed_plain(&mut model, 0, 6, 1, "=_xll.FDS(B6,\"EV_SALES\")", 2.292617);
    seed_plain(&mut model, 0, 5, 3, "='[2]Other'!B5", 7.0);
    seed_plain(&mut model, 0, 6, 3, "=Missing!B6*2", 9.0);
    model._set("D5", "=A5+C5");
    model._set("D6", "=A6+C6");
    model.evaluate();
    model
}

/// The four saved values, their kind and the sums, with the block's first
/// row and the columns it now occupies.
fn assert_saved_values(model: &Model, row: i32, columns: [&str; 3]) {
    let [addin, link, sum] = columns;
    let (first, second) = (row, row + 1);
    assert_eq!(number(model, &format!("{addin}{first}")), 2.549606);
    assert_eq!(number(model, &format!("{addin}{second}")), 2.292617);
    assert_eq!(number(model, &format!("{link}{first}")), 7.0);
    assert_eq!(number(model, &format!("{link}{second}")), 9.0);
    assert_eq!(number(model, &format!("{sum}{first}")), 2.549606 + 7.0);
    assert_eq!(number(model, &format!("{sum}{second}")), 2.292617 + 9.0);
    for column in [addin, link] {
        for r in [first, second] {
            let cell = format!("{column}{r}");
            assert_eq!(kind(model, &cell), "formula", "{cell}");
            assert!(reason(model, &cell).is_some(), "{cell} is not kept as saved");
        }
    }
    assert!(reason(model, &format!("{sum}{first}")).is_none());
}

#[test]
fn saved_values_follow_their_cells_through_row_and_column_edits() {
    let mut model = model_with_saved_values();
    assert_saved_values(&model, 5, ["A", "C", "D"]);

    // A row above: everything moves down one, and the formulas follow.
    model.insert_rows(0, 2, 1).unwrap();
    model.evaluate();
    assert_saved_values(&model, 6, ["A", "C", "D"]);
    assert_eq!(
        model._get_formula("A6").to_uppercase(),
        "=_XLL.FDS(B6,\"EV_SALES\")"
    );
    assert_eq!(model._get_formula("D7"), "=A7+C7");
    assert_eq!(kind(&model, "A5"), "none");

    // Rows removed above: up two.
    model.delete_rows(0, 1, 2).unwrap();
    model.evaluate();
    assert_saved_values(&model, 4, ["A", "C", "D"]);

    // Columns in front: right two.
    model.insert_columns(0, 1, 2).unwrap();
    model.evaluate();
    assert_saved_values(&model, 4, ["C", "E", "F"]);
    assert_eq!(
        model._get_formula("C4").to_uppercase(),
        "=_XLL.FDS(D4,\"EV_SALES\")"
    );

    // A column removed in front: left one.
    model.delete_columns(0, 2, 1).unwrap();
    model.evaluate();
    assert_saved_values(&model, 4, ["B", "D", "E"]);

    // An edit that moves nothing leaves them alone too.
    model.insert_rows(0, 50, 3).unwrap();
    model.insert_columns(0, 20, 1).unwrap();
    model.evaluate();
    assert_saved_values(&model, 4, ["B", "D", "E"]);
}

#[test]
fn deleting_the_row_of_a_saved_value_does_not_hand_it_to_the_next_cell() {
    let mut model = model_with_saved_values();
    // Row 5 goes; row 6 becomes row 5 and keeps ITS values.
    model.delete_rows(0, 5, 1).unwrap();
    model._set("A6", "=1+1");
    model._set("C6", "=A6*10");
    model.evaluate();
    assert_eq!(number(&model, "A5"), 2.292617);
    assert_eq!(number(&model, "C5"), 9.0);
    assert_eq!(number(&model, "D5"), 2.292617 + 9.0);
    // The cells that took the old coordinates are ordinary formulas.
    assert_eq!(number(&model, "A6"), 2.0);
    assert_eq!(number(&model, "C6"), 20.0);
    assert!(reason(&model, "C6").is_none());
}

#[test]
fn saved_values_follow_a_moved_row_and_column() {
    let mut model = model_with_saved_values();
    // Row 5 goes below row 6.
    model.move_rows_action(0, 5, 1, 1).unwrap();
    model.evaluate();
    assert_eq!(number(&model, "A6"), 2.549606);
    assert_eq!(number(&model, "A5"), 2.292617);
    assert_eq!(number(&model, "C6"), 7.0);
    assert_eq!(number(&model, "C5"), 9.0);
    assert_eq!(number(&model, "D6"), 2.549606 + 7.0);
    assert_eq!(number(&model, "D5"), 2.292617 + 9.0);
    for cell in ["A5", "A6", "C5", "C6"] {
        assert_eq!(kind(&model, cell), "formula", "{cell}");
    }

    // Column C goes in front of column A.
    model.move_columns_action(0, 3, 1, -2).unwrap();
    model.evaluate();
    assert_eq!(number(&model, "A6"), 7.0);
    assert_eq!(number(&model, "A5"), 9.0);
    assert_eq!(number(&model, "B6"), 2.549606);
    assert_eq!(number(&model, "B5"), 2.292617);
    assert_eq!(number(&model, "D6"), 2.549606 + 7.0);
    for cell in ["A5", "A6", "B5", "B6"] {
        assert_eq!(kind(&model, cell), "formula", "{cell}");
        assert!(reason(&model, cell).is_some(), "{cell}");
    }
}

#[test]
fn text_is_not_read_again_as_if_typed() {
    let texts = [
        "007",
        "+1 212 526 7952",
        "=======",
        "--",
        "+19.2%",
        "+/- bps, y/y",
        "TRUE",
        "#N/A",
        "$100",
        "1e5",
        "'quoted",
        "",
    ];
    let mut model = new_empty_model();
    for (i, text) in texts.iter().enumerate() {
        model
            .seed_cell_literal(0, 3 + i as i32, 2, CellValue::String(text.to_string()))
            .unwrap();
    }
    let styles: Vec<i32> = (0..texts.len())
        .map(|i| style(&model, &format!("B{}", 3 + i)))
        .collect();
    model.evaluate();

    let check = |model: &Model, column: &str, first_row: usize| {
        for (i, text) in texts.iter().enumerate() {
            let cell = format!("{column}{}", first_row + i);
            assert_eq!(kind(model, &cell), "text", "{cell} ({text:?})");
            assert_eq!(value(model, &cell), CellValue::String(text.to_string()));
            assert_eq!(style(model, &cell), styles[i], "{cell} ({text:?})");
        }
    };

    model.insert_rows(0, 1, 2).unwrap();
    model.evaluate();
    check(&model, "B", 5);
    model.delete_rows(0, 2, 1).unwrap();
    model.evaluate();
    check(&model, "B", 4);
    model.insert_columns(0, 1, 1).unwrap();
    model.evaluate();
    check(&model, "C", 4);
    model.delete_columns(0, 1, 2).unwrap();
    model.evaluate();
    check(&model, "A", 4);
    model.move_rows_action(0, 4, 1, 20).unwrap();
    model.evaluate();
    assert_eq!(value(&model, "A24"), CellValue::String("007".to_string()));
    // The rows it passed moved up one.
    assert_eq!(
        value(&model, "A4"),
        CellValue::String("+1 212 526 7952".to_string())
    );
    assert_eq!(kind(&model, "A24"), "text");
}

#[test]
fn numbers_booleans_and_errors_move_untouched() {
    // Each of these has more digits than the 15 its text shows.
    let numbers = [0.1 + 0.2, 1.0 / 3.0, 123456789.12345679, 2.549606e-7 / 3.0];
    let mut model = new_empty_model();
    for (i, n) in numbers.iter().enumerate() {
        model
            .seed_cell_literal(0, 2 + i as i32, 1, CellValue::Number(*n))
            .unwrap();
    }
    model
        .seed_cell_literal(0, 10, 1, CellValue::Boolean(true))
        .unwrap();
    model._set("A11", "#N/A");
    model._set("C2", "=A2*3");
    model.evaluate();
    let tripled = number(&model, "C2");

    model.insert_rows(0, 1, 1).unwrap();
    model.insert_columns(0, 1, 1).unwrap();
    model.evaluate();
    for (i, n) in numbers.iter().enumerate() {
        let cell = format!("B{}", 3 + i);
        assert_eq!(number(&model, &cell).to_bits(), n.to_bits(), "{cell}");
    }
    assert_eq!(kind(&model, "B11"), "boolean");
    assert_eq!(value(&model, "B11"), CellValue::Boolean(true));
    assert_eq!(kind(&model, "B12"), "error");
    assert_eq!(model._get_text("B12"), "#N/A");
    assert_eq!(number(&model, "D3").to_bits(), tripled.to_bits());
}

#[test]
fn a_plain_formula_stays_plain() {
    let mut model = new_empty_model();
    for row in 1..=5 {
        model._set(&format!("A{row}"), &format!("{}", row * 10));
    }
    // A column reference in an ordinary cell: Excel (and the engine) take
    // the element on the cell's own row. As an array formula it would spill.
    seed_plain(&mut model, 0, 3, 3, "=A1:A5", 30.0);
    model.evaluate();
    assert_eq!(kind(&model, "C3"), "formula");
    assert_eq!(number(&model, "C3"), 30.0);

    model.insert_rows(0, 1, 1).unwrap();
    model.evaluate();
    assert_eq!(model._get_formula("C4"), "=A2:A6");
    assert_eq!(kind(&model, "C4"), "formula");
    assert_eq!(number(&model, "C4"), 30.0);
    assert_eq!(kind(&model, "C5"), "none");

    model.insert_columns(0, 2, 1).unwrap();
    model.evaluate();
    assert_eq!(kind(&model, "D4"), "formula");
    assert_eq!(number(&model, "D4"), 30.0);
    assert_eq!(kind(&model, "D5"), "none");

    // The reference shifts while the cell stays put.
    model.delete_rows(0, 6, 1).unwrap();
    model.evaluate();
    assert_eq!(model._get_formula("D4"), "=A2:A5");
    assert_eq!(kind(&model, "D4"), "formula");
    assert_eq!(number(&model, "D4"), 30.0);
}

#[test]
fn an_array_formula_stays_an_array_formula() {
    let mut model = new_empty_model();
    model._set("B3", "4");
    // The file marked this one as a dynamic array formula; its text alone
    // does not say so.
    model._set("A3", "=B3+1");
    let worksheet = model.workbook.worksheet_mut(0).unwrap();
    if let Some(Cell::CellFormula { f, s, .. }) = worksheet.cell(3, 1).cloned() {
        let cell = Cell::ArrayFormula {
            f,
            s,
            r: (1, 1),
            kind: ArrayKind::Dynamic,
            v: FormulaValue::Number(5.0),
        };
        worksheet.update_cell(3, 1, cell).unwrap();
    }
    model.evaluate();
    assert_eq!(kind(&model, "A3"), "dynamic");

    model.insert_rows(0, 1, 1).unwrap();
    model.evaluate();
    assert_eq!(kind(&model, "A4"), "dynamic");
    assert_eq!(number(&model, "A4"), 5.0);
    model.delete_columns(0, 1, 1).unwrap();
    model.evaluate();
    // Column A went, taking the formula with it; B4 (the 4) is A4 now.
    assert_eq!(number(&model, "A4"), 4.0);
}

#[test]
fn a_dynamic_array_spills_again_where_it_lands() {
    let mut model = new_empty_model();
    model._set("A2", "=SEQUENCE(3)");
    model._set("C1", "=SUM(A2:A4)");
    model.evaluate();
    assert_eq!(number(&model, "C1"), 6.0);

    model.insert_rows(0, 2, 1).unwrap();
    model.evaluate();
    assert_eq!(kind(&model, "A3"), "dynamic");
    assert_eq!(kind(&model, "A4"), "spill");
    assert_eq!(number(&model, "A3"), 1.0);
    assert_eq!(number(&model, "A5"), 3.0);
    assert_eq!(number(&model, "C1"), 6.0);

    // A row through the middle of the spill: the array spills over it.
    model.insert_rows(0, 4, 1).unwrap();
    model.evaluate();
    assert_eq!(number(&model, "A3"), 1.0);
    assert_eq!(number(&model, "A4"), 2.0);
    assert_eq!(number(&model, "A5"), 3.0);
    assert_eq!(model._get_text("A6"), "");
}

#[test]
fn a_cse_array_moves_as_a_block() {
    let mut model = new_empty_model();
    model._set("B3", "1");
    model._set("B4", "2");
    model._set("B5", "3");
    model
        .set_user_array_formula(0, 3, 1, 1, 3, "=B3:B5*2")
        .unwrap();
    model.evaluate();
    let check = |model: &Model, column: &str, first_row: i32| {
        for (i, expected) in [2.0, 4.0, 6.0].iter().enumerate() {
            let cell = format!("{column}{}", first_row + i as i32);
            assert_eq!(number(model, &cell), *expected, "{cell}");
            assert_eq!(kind(model, &cell), if i == 0 { "cse" } else { "spill" }, "{cell}");
        }
    };
    check(&model, "A", 3);

    model.insert_rows(0, 1, 1).unwrap();
    model.evaluate();
    check(&model, "A", 4);
    assert_eq!(model._get_formula("A4"), "=B4:B6*2");

    // The block lands on rows it occupied a moment ago.
    model.delete_rows(0, 2, 1).unwrap();
    model.evaluate();
    check(&model, "A", 3);

    model.insert_columns(0, 1, 1).unwrap();
    model.evaluate();
    check(&model, "B", 3);
    assert_eq!(model._get_formula("B3"), "=C3:C5*2");

    model.delete_columns(0, 1, 1).unwrap();
    model.evaluate();
    check(&model, "A", 3);

    // And it still refuses to be split.
    assert!(model.insert_rows(0, 4, 1).is_err());
    check(&model, "A", 3);
}

#[test]
fn an_overwritten_link_cell_is_calculated_again() {
    let mut model = new_empty_model();
    seed_plain(&mut model, 0, 1, 2, "=Missing!A1", 5.0);
    model.evaluate();
    assert_eq!(number(&model, "B1"), 5.0);
    assert!(reason(&model, "B1").is_some());

    // The user replaces the dead link with a working formula.
    model._set("A1", "3");
    model._set("B1", "=A1*2");
    model.evaluate();
    assert_eq!(number(&model, "B1"), 6.0);
    // It follows its input from then on, like any formula.
    model._set("A1", "10");
    model.evaluate();
    assert_eq!(number(&model, "B1"), 20.0);
    assert!(reason(&model, "B1").is_none());
}

#[test]
fn saved_values_follow_their_cells_when_sheets_move() {
    let mut model = new_empty_model();
    model.new_sheet();
    seed_plain(&mut model, 0, 4, 2, "=Missing!A1+1", 43.0);
    seed_plain(&mut model, 1, 4, 2, "=Missing!A1", 42.0);
    model._set("Sheet1!C4", "=B4*2");
    model._set("Sheet2!C4", "=B4*2");
    model.evaluate();
    let check = |model: &Model| {
        assert_eq!(number(model, "Sheet1!B4"), 43.0);
        assert_eq!(number(model, "Sheet2!B4"), 42.0);
        assert_eq!(number(model, "Sheet1!C4"), 86.0);
        assert_eq!(number(model, "Sheet2!C4"), 84.0);
        assert!(reason(model, "Sheet1!B4").is_some());
        assert!(reason(model, "Sheet2!B4").is_some());
        assert!(reason(model, "Sheet1!C4").is_none());
    };
    check(&model);

    // A sheet in front: both sheets are one index further along.
    model.insert_sheet("First", 0, None).unwrap();
    model.evaluate();
    assert_eq!(
        model.workbook.get_worksheet_names(),
        ["First", "Sheet1", "Sheet2"]
    );
    check(&model);

    // The last sheet goes to the front.
    model.move_sheet(2, 0).unwrap();
    model.evaluate();
    assert_eq!(
        model.workbook.get_worksheet_names(),
        ["Sheet2", "First", "Sheet1"]
    );
    check(&model);

    // And back behind the others.
    model.move_sheet(0, 2).unwrap();
    model.evaluate();
    check(&model);
    model.move_sheet(2, 0).unwrap();
    model.evaluate();

    // The sheet between them goes.
    model.delete_sheet(1).unwrap();
    model.evaluate();
    assert_eq!(model.workbook.get_worksheet_names(), ["Sheet2", "Sheet1"]);
    check(&model);

    // A copy of a sheet keeps the saved values of the original, and the
    // sheet behind the copy keeps its own.
    let (_, copy) = model.duplicate_sheet(0).unwrap();
    model.evaluate();
    assert_eq!(copy, 1);
    assert_eq!(
        model.get_cell_value_by_index(copy, 4, 2).unwrap(),
        CellValue::Number(42.0)
    );
    assert_eq!(
        model.get_cell_value_by_index(copy, 4, 3).unwrap(),
        CellValue::Number(84.0)
    );
    assert!(model.imported_value_reason(copy, 4, 2).is_some());
    check(&model);
    model.delete_sheet(copy).unwrap();
    model.evaluate();
    check(&model);

    // A reference that breaks in THIS session is an error, not a kept value.
    model._set("Sheet1!D1", "=Sheet2!C4");
    model.evaluate();
    assert_eq!(number(&model, "Sheet1!D1"), 84.0);
    model.delete_sheet(0).unwrap();
    model.evaluate();
    assert_eq!(model._get_text("Sheet1!D1"), "#REF!");
    assert!(reason(&model, "Sheet1!D1").is_none());
    assert_eq!(number(&model, "Sheet1!B4"), 43.0);
    assert_eq!(number(&model, "Sheet1!C4"), 86.0);
}
