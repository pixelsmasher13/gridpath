#![allow(clippy::unwrap_used)]

use crate::cell::CellValue;
use crate::test::util::new_empty_model;
use crate::types::IterationSettings;

fn number(model: &crate::Model, row: i32, column: i32) -> f64 {
    match model.get_cell_value_by_index(0, row, column).unwrap() {
        CellValue::Number(n) => n,
        other => panic!("not a number at R{row}C{column}: {other:?}"),
    }
}

fn excel_defaults() -> Option<IterationSettings> {
    Some(IterationSettings { max_iterations: 100, max_change: 0.001 })
}

/// Off (the default), a circular reference is #CIRC!.
#[test]
fn circular_reference_without_iteration_is_circ() {
    let mut model = new_empty_model();
    model._set("A1", "=A1/2+1");
    model.evaluate();
    assert_eq!(model._get_text("A1"), *"#CIRC!");
}

/// On, a converging circle settles within max_change of its fixed point,
/// and cells reading the circle see the settled value.
#[test]
fn converging_circles() {
    let mut model = new_empty_model();
    model.set_iteration(excel_defaults());
    model._set("A1", "=A1/2+1"); // → 2
    model._set("B1", "=C1*0.5+10"); // B1 and C1 → 20
    model._set("C1", "=B1");
    model._set("D1", "=B1*2"); // outside the circle → 40
    model.evaluate();
    assert!((number(&model, 1, 1) - 2.0).abs() < 0.002, "{}", number(&model, 1, 1));
    assert!((number(&model, 1, 2) - 20.0).abs() < 0.01, "{}", number(&model, 1, 2));
    assert!((number(&model, 1, 3) - 20.0).abs() < 0.01, "{}", number(&model, 1, 3));
    assert!((number(&model, 1, 4) - number(&model, 1, 2) * 2.0).abs() < 1e-9);
}

/// A circle that never settles stops after max_iterations rounds, with its
/// last values (Excel: =A1+1 with 100 iterations shows 100).
#[test]
fn diverging_circle_stops_at_max_iterations() {
    let mut model = new_empty_model();
    model.set_iteration(Some(IterationSettings { max_iterations: 100, max_change: 0.001 }));
    model._set("A1", "=A1+1");
    model.evaluate();
    assert_eq!(number(&model, 1, 1), 100.0);
}

/// An edit feeding the circle re-converges in the incremental pass.
#[test]
fn edit_into_a_circle_reconverges() {
    let mut model = new_empty_model();
    model.set_iteration(excel_defaults());
    model._set("E1", "10");
    model._set("B1", "=C1*0.5+E1");
    model._set("C1", "=B1");
    model.evaluate();
    assert!((number(&model, 1, 2) - 20.0).abs() < 0.01);
    model.set_user_input(0, 1, 5, "30".to_string()).unwrap();
    let changed = model.evaluate_edited(&[(0, 1, 5)]);
    assert!((number(&model, 1, 2) - 60.0).abs() < 0.01, "{}", number(&model, 1, 2));
    assert!((number(&model, 1, 3) - 60.0).abs() < 0.01);
    assert!(changed.contains(&(0, 1, 2)) && changed.contains(&(0, 1, 3)), "{changed:?}");
}

/// TRIM collapses runs of spaces inside the text and drops them at the ends;
/// other whitespace stays (Excel).
#[test]
fn trim_collapses_internal_spaces() {
    let mut model = new_empty_model();
    model._set("A1", "=TRIM(\"  a   b  \")");
    model._set("A2", "=TRIM(\"a\tb\")");
    model._set("A3", "=TRIM(\"one  two   three\")");
    model._set("A4", "=TRIM(\"   \")");
    model._set("A5", "=LEN(TRIM(CHAR(160)&\" x \"))");
    model.evaluate();
    assert_eq!(model._get_text("A1"), *"a b");
    assert_eq!(model._get_text("A2"), *"a\tb");
    assert_eq!(model._get_text("A3"), *"one two three");
    assert_eq!(model._get_text("A4"), *"");
    assert_eq!(model._get_text("A5"), *"3");
}
