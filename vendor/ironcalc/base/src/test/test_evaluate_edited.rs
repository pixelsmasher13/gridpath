#![allow(clippy::unwrap_used)]

//! `Model::evaluate_edited` (gridpath incremental recalculation) must agree
//! with a full `evaluate` and report exactly the cells whose values moved.

use crate::test::util::new_empty_model;

const ROWS: i32 = 60;

/// A supply-style sheet: type in A, quantity in B, and in C a SUMIFS over the
/// SAME absolute ranges on every row (the shape that made the reader index
/// rescan its column lists once per dirty cell).
fn supply_model<'a>() -> crate::model::Model<'a> {
    let mut model = new_empty_model();
    for r in 1..=ROWS {
        model._set(&format!("A{r}"), ["red", "green", "blue"][(r % 3) as usize]);
        model._set(&format!("B{r}"), &r.to_string());
        model._set(&format!("C{r}"), &format!("=SUMIFS($B$1:$B${ROWS},$A$1:$A${ROWS},A{r})"));
    }
    model._set("D1", &format!("=SUM(C1:C{ROWS})"));
    model._set("E1", "=D1*2");
    model.evaluate();
    model
}

fn texts(model: &crate::model::Model) -> Vec<String> {
    let mut out = Vec::new();
    for r in 1..=ROWS {
        for c in 1..=5 {
            out.push(model._get_text_at(0, r, c));
        }
    }
    out
}

/// Edit, recalculate incrementally, and compare with a full recalculation
/// of the same edits.
fn check(edits: &[(&str, &str)]) {
    let mut inc = supply_model();
    let mut full = supply_model();
    for (cell, value) in edits {
        let before = texts(&inc);
        inc._set(cell, value);
        let (row, column) = {
            let col: String = cell.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
            let row: i32 = cell[col.len()..].parse().unwrap();
            (row, col.bytes().fold(0, |n, b| n * 26 + (b - b'A' + 1) as i32))
        };
        let changed = inc.evaluate_edited(&[(0, row, column)]);
        full._set(cell, value);
        full.evaluate();
        assert_eq!(texts(&inc), texts(&full), "after setting {cell} to {value}");
        // Every cell that moved is reported; nothing unchanged is.
        let after = texts(&inc);
        for r in 1..=ROWS {
            for c in 1..=5 {
                let i = ((r - 1) * 5 + (c - 1)) as usize;
                let moved = before[i] != after[i];
                let reported = changed.contains(&(0, r, c));
                if (r, c) == (row, column) {
                    continue;
                }
                assert_eq!(moved, reported, "cell ({r},{c}) moved={moved} reported={reported} after {cell}={value}");
            }
        }
    }
}

#[test]
fn shared_range_sumifs_matches_full_evaluate() {
    check(&[("B5", "100")]);
}

#[test]
fn successive_edits_keep_the_index_consistent() {
    // The second and third edits rely on the edges the first re-recorded.
    check(&[("B5", "100"), ("B6", "7"), ("A7", "red"), ("B5", "5")]);
}

#[test]
fn criteria_change_moves_only_affected_groups() {
    // Changing a type moves the totals of the two types involved, not the third.
    check(&[("A3", "green")]);
}

#[test]
fn edit_with_no_dependents_reports_nothing() {
    let mut model = supply_model();
    model._set("H40", "1");
    assert!(model.evaluate_edited(&[(0, 40, 8)]).is_empty());
}

#[test]
fn value_that_does_not_change_results_is_not_reported() {
    // B4 is "4"; writing 4 again re-evaluates its readers but moves nothing.
    let mut model = supply_model();
    model._set("B4", "4");
    let changed = model.evaluate_edited(&[(0, 4, 2)]);
    assert!(changed.is_empty(), "reported {changed:?}");
}
