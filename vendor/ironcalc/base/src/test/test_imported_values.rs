#![allow(clippy::unwrap_used)]

use crate::model::ImportedValueReason;
use crate::test::util::new_empty_model;
use crate::types::{Cell, FormulaValue};

/// The formula cell as an imported file has it: the formula plus the value
/// Excel saved for it.
fn with_saved_value(model: &mut crate::Model, row: i32, column: i32, value: f64) {
    let cell = model.workbook.worksheets[0].sheet_data.get_mut(&row).unwrap().get_mut(&column).unwrap();
    match cell {
        Cell::CellFormula { v, .. } | Cell::ArrayFormula { v, .. } => *v = FormulaValue::Number(value),
        other => panic!("not a formula: {other:?}"),
    }
}

/// An unknown function with a saved value keeps it (dependents calculate
/// from it) and is reported with the function's name; one with nothing saved
/// calculates to its error and is not.
#[test]
fn unknown_functions_keep_the_saved_value_and_say_why() {
    let mut model = new_empty_model();
    model._set("A1", "=_xll.FDS(\"x\")");
    model._set("B1", "=A1*2");
    model._set("C1", "=NOSUCHFN(1)");
    model._set("D1", "=1+1");
    with_saved_value(&mut model, 1, 1, 42.0);
    model.evaluate();
    assert_eq!(model._get_text("A1"), *"42");
    assert_eq!(model._get_text("B1"), *"84");
    assert_eq!(model._get_text("C1"), *"#NAME?");
    let cells = model.imported_value_cells(10);
    assert_eq!(cells.len(), 1, "{cells:?}");
    assert_eq!((cells[0].0, cells[0].1, cells[0].2), (0, 1, 1));
    assert!(
        matches!(&cells[0].3, ImportedValueReason::UnknownFunction(name) if name.eq_ignore_ascii_case("_xll.FDS")),
        "{:?}",
        cells[0].3
    );
    assert_eq!(model.imported_value_reason(0, 1, 1), Some(cells[0].3.clone()));
    assert_eq!(model.imported_value_reason(0, 1, 2), None);
    assert_eq!(model.imported_value_reason(0, 1, 3), None);
    assert_eq!(model.imported_value_reason(0, 1, 4), None);
    assert!(model.imported_value_cells(0).is_empty());
}
