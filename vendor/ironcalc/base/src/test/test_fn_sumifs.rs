#![allow(clippy::unwrap_used)]

use crate::model::Model;
use crate::test::util::new_empty_model;

#[test]
fn test_fn_sumifs_arguments() {
    let mut model = new_empty_model();

    // Incorrect number of arguments
    model._set("A1", "=SUMIFS()");
    model._set("A2", "=SUMIFS(B2:B9)");
    model._set("A3", "=SUMIFS(B2:B9,C2:C9)");
    model._set("A4", "=SUMIFS(B2:B9,C2:C9,\"=A*\",D2:D9)");

    // Correct (Sum everything in column 'B' if column 'C' starts with "A")
    model._set("A5", "=SUMIFS(B2:B9,C2:C9,\"=A*\")");

    // Data
    model._set("B2", "5");
    model._set("B3", "4");
    model._set("B4", "15");
    model._set("B5", "22");
    model._set("B6", "=NA()");
    model._set("C2", "Apples");
    model._set("C3", "Bananas");
    model._set("C4", "Almonds");
    model._set("C5", "Yoni");
    model._set("C6", "Mandarin");

    model.evaluate();

    // Error (Incorrect number of arguments)
    assert_eq!(model._get_text("A1"), *"#ERROR!");
    assert_eq!(model._get_text("A2"), *"#ERROR!");
    assert_eq!(model._get_text("A3"), *"#ERROR!");
    assert_eq!(model._get_text("A4"), *"#ERROR!");

    // Correct
    assert_eq!(model._get_text("A5"), *"20");
}

// Date-string criteria reach build_criteria via two distinct paths:
//   * fn_countifs has its own criteria loop (covered by test_fn_count.rs)
//   * apply_ifs is the shared path behind SUMIF, SUMIFS, AVERAGEIF,
//     AVERAGEIFS, MAXIFS, MINIFS — exercised here.
//
// Layout:
//   B2..B8  = 45131..45137 (2023-07-24 .. 2023-07-30, all "<7/31/2023")
//   B9      = 45138        (2023-07-31, the criterion target)
//   B10..B11= 45139, 45200 (after 7/31/2023)
//   C2..C8  = 1..7         (matches "<7/31/2023")
//   C9      = 100          (matches "7/31/2023")
//   C10..C11= 99, 50       (matches ">7/31/2023")
fn populate_date_apply_ifs_data(model: &mut Model) {
    let date_serials = [45131, 45132, 45133, 45134, 45135, 45136, 45137];
    let values = [1, 2, 3, 4, 5, 6, 7];
    for (idx, (serial, value)) in date_serials.iter().zip(values.iter()).enumerate() {
        let row = idx + 2;
        model._set(&format!("B{row}"), &serial.to_string());
        model._set(&format!("C{row}"), &value.to_string());
    }
    model._set("B9", "45138");
    model._set("C9", "100");
    model._set("B10", "45139");
    model._set("C10", "99");
    model._set("B11", "45200");
    model._set("C11", "50");
}

#[test]
fn test_apply_ifs_date_criterion_covers_all_consumers() {
    // One worksheet, one criterion shape, every apply_ifs consumer probed.
    // If a future change drops the locale at the apply_ifs call site, every
    // assertion below collapses to 0 / DIV-by-0 and the test fails loudly.
    let mut model = new_empty_model();
    populate_date_apply_ifs_data(&mut model);

    // SUMIF / SUMIFS — sum of C where B "<7/31/2023" is 1+2+3+4+5+6+7 = 28
    model._set("A1", "=SUMIF(B2:B11, \"<7/31/2023\", C2:C11)");
    model._set("A2", "=SUMIFS(C2:C11, B2:B11, \"<7/31/2023\")");
    // AVERAGEIF / AVERAGEIFS — 28 / 7 = 4
    model._set("A3", "=AVERAGEIF(B2:B11, \"<7/31/2023\", C2:C11)");
    model._set("A4", "=AVERAGEIFS(C2:C11, B2:B11, \"<7/31/2023\")");
    // MAXIFS / MINIFS over the same range
    model._set("A5", "=MAXIFS(C2:C11, B2:B11, \"<7/31/2023\")");
    model._set("A6", "=MINIFS(C2:C11, B2:B11, \"<7/31/2023\")");

    // Use a different operator (>=) on SUMIFS / MAXIFS / MINIFS to exercise
    // a second branch of build_criteria through the apply_ifs path.
    // C cells where B >= 7/31/2023 are C9=100, C10=99, C11=50 -> sum 249
    model._set("A7", "=SUMIFS(C2:C11, B2:B11, \">=7/31/2023\")");
    model._set("A8", "=MAXIFS(C2:C11, B2:B11, \">=7/31/2023\")");
    model._set("A9", "=MINIFS(C2:C11, B2:B11, \">=7/31/2023\")");

    model.evaluate();

    assert_eq!(model._get_text("A1"), *"28");
    assert_eq!(model._get_text("A2"), *"28");
    assert_eq!(model._get_text("A3"), *"4");
    assert_eq!(model._get_text("A4"), *"4");
    assert_eq!(model._get_text("A5"), *"7");
    assert_eq!(model._get_text("A6"), *"1");
    assert_eq!(model._get_text("A7"), *"249");
    assert_eq!(model._get_text("A8"), *"100");
    assert_eq!(model._get_text("A9"), *"50");
}

#[test]
fn test_apply_ifs_date_criterion_respects_locale() {
    // Build the same workbook under en-GB (D/M/Y). "31/7/2023" is only a
    // valid date under a D/M/Y locale; if apply_ifs ever stops forwarding
    // self.locale to build_criteria, en's M/D/Y rules would reject this
    // string and SUMIFS would collapse to 0.
    let mut model = Model::new_empty("model", "en-GB", "UTC", "en").unwrap();
    populate_date_apply_ifs_data(&mut model);

    model._set("A1", "=SUMIFS(C2:C11, B2:B11, \"<31/7/2023\")");
    model._set("A2", "=AVERAGEIFS(C2:C11, B2:B11, \"<31/7/2023\")");
    model._set("A3", "=MAXIFS(C2:C11, B2:B11, \"<31/7/2023\")");
    model._set("A4", "=MINIFS(C2:C11, B2:B11, \"<31/7/2023\")");

    model.evaluate();

    assert_eq!(model._get_text("A1"), *"28");
    assert_eq!(model._get_text("A2"), *"4");
    assert_eq!(model._get_text("A3"), *"7");
    assert_eq!(model._get_text("A4"), *"1");
}

/// The per-pass criteria index (`Model::ifs_cache`) must be invisible:
/// many SUMIFS sharing one criteria column, multi-criteria intersections,
/// criteria ranges offset from the sum range, errors only where a MATCHED
/// sum cell errors, and criteria edits picked up by the next pass.
#[test]
fn test_sumifs_shared_criteria_index_matches_plain_walk() {
    let mut model = new_empty_model();
    // Data: B = amount, C = type, D = region. Row 8's amount is an error but
    // its type never matches, so it must not poison the colocation totals.
    let rows = [
        ("5", "colocation", "east"),
        ("4", "hyperscaler", "west"),
        ("15", "colocation", "west"),
        ("22", "colocation", "east"),
        ("=NA()", "other", "east"),
        ("7", "hyperscaler", "east"),
        ("", "colocation", "west"),
    ];
    for (i, (b, c, d)) in rows.iter().enumerate() {
        let r = i + 2;
        if !b.is_empty() {
            model._set(&format!("B{r}"), b);
        }
        model._set(&format!("C{r}"), c);
        model._set(&format!("D{r}"), d);
    }
    // Forty formulas over the same criteria column, several criteria.
    for r in 20..40 {
        model._set(&format!("A{r}"), "=SUMIFS($B$2:$B$8,$C$2:$C$8,\"colocation\")");
        model._set(&format!("E{r}"), "=SUMIFS($B$2:$B$8,$C$2:$C$8,\"hyperscaler\")");
    }
    model._set("F1", "=SUMIFS(B2:B8,C2:C8,\"colocation\",D2:D8,\"east\")");
    model._set("F2", "=SUMIFS(B2:B8,C2:C8,\"colocation\",D2:D8,\"west\")");
    // Criteria range offset from the sum range (rows 3..9 against 2..8).
    model._set("F3", "=SUMIFS(B2:B8,C3:C9,\"colocation\")");
    // A matched error must propagate.
    model._set("F4", "=SUMIFS(B2:B8,C2:C8,\"other\")");
    // Whole-column ranges.
    model._set("F5", "=SUMIFS(B:B,C:C,\"colocation\",D:D,\"east\")");
    // Numeric and comparison criteria share the index keyed by criterion.
    model._set("F6", "=SUMIFS(B2:B8,B2:B8,\">10\")");
    model._set("F7", "=SUMIFS(B2:B8,B2:B8,15)");

    model.evaluate();

    assert_eq!(model._get_text("A20"), *"42");
    assert_eq!(model._get_text("A39"), *"42");
    assert_eq!(model._get_text("E20"), *"11");
    assert_eq!(model._get_text("F1"), *"27");
    assert_eq!(model._get_text("F2"), *"15");
    // Offset criteria: C3..C9 = hyper, colo, colo, other, hyper, colo, "" → B3,B4,B7 = 4+15+7
    assert_eq!(model._get_text("F3"), *"26");
    assert_eq!(model._get_text("F4"), *"#N/A");
    assert_eq!(model._get_text("F5"), *"27");
    assert_eq!(model._get_text("F6"), *"37");
    assert_eq!(model._get_text("F7"), *"15");

    // Editing a criteria cell invalidates the index on the next pass.
    model._set("C3", "colocation");
    model.evaluate();
    assert_eq!(model._get_text("A20"), *"46");
    assert_eq!(model._get_text("E39"), *"7");
    assert_eq!(model._get_text("F2"), *"19");
}
