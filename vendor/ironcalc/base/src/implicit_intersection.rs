use crate::{calc_result::Range, expressions::types::CellReferenceIndex};

/// It returns the closest cell from cell_reference to range in the same column/row
/// Examples
///  * i_i(B5, A2:A9) -> A5
///  * i_i(B5, A7:A9) -> None
///  * i_i(B5, A2:D2) -> B2
///
/// Implicit intersection works across sheets: the consuming cell contributes
/// its row/column position, but the intersected cell lives on the RANGE's
/// sheet (`@Model!A:A` in ModelSummary!B5 reads Model!A5). A single-cell
/// range needs no intersection at all — Excel's `@` passes it through
/// regardless of position or sheet (the Canalyst pattern:
/// `@INDIRECT("Model!"&ADDRESS(row, col))` consumed from a summary sheet).
pub(crate) fn implicit_intersection(
    cell_reference: &CellReferenceIndex,
    range: &Range,
) -> Option<CellReferenceIndex> {
    let left = &range.left;
    let right = &range.right;
    // Ranges are single-sheet; the result must point into that sheet.
    let sheet = left.sheet;
    if left.row == right.row && left.column == right.column {
        // Single cell: nothing to intersect.
        return Some(CellReferenceIndex {
            sheet,
            row: left.row,
            column: left.column,
        });
    }
    let row = cell_reference.row;
    let column = cell_reference.column;
    if row >= left.row && row <= right.row {
        if left.column != right.column {
            return None;
        }
        return Some(CellReferenceIndex {
            sheet,
            row,
            column: left.column,
        });
    }
    if column >= left.column && column <= right.column {
        if left.row != right.row {
            return None;
        }
        return Some(CellReferenceIndex {
            sheet,
            row: left.row,
            column,
        });
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cell(sheet: u32, row: i32, column: i32) -> CellReferenceIndex {
        CellReferenceIndex { sheet, row, column }
    }

    #[test]
    fn single_cell_passes_through_regardless_of_position_and_sheet() {
        // @Model!$BD$125 consumed from ModelSummary!J6
        let r = Range {
            left: cell(1, 125, 56),
            right: cell(1, 125, 56),
        };
        assert_eq!(implicit_intersection(&cell(0, 6, 10), &r), Some(cell(1, 125, 56)));
    }

    #[test]
    fn column_vector_intersects_by_row_across_sheets() {
        // @Model!A2:A9 consumed from Summary!B5 -> Model!A5
        let r = Range {
            left: cell(1, 2, 1),
            right: cell(1, 9, 1),
        };
        assert_eq!(implicit_intersection(&cell(0, 5, 2), &r), Some(cell(1, 5, 1)));
    }

    #[test]
    fn row_vector_intersects_by_column() {
        let r = Range {
            left: cell(0, 2, 1),
            right: cell(0, 2, 4),
        };
        assert_eq!(implicit_intersection(&cell(0, 5, 2), &r), Some(cell(0, 2, 2)));
    }

    #[test]
    fn no_intersection_outside_span() {
        let r = Range {
            left: cell(0, 7, 1),
            right: cell(0, 9, 1),
        };
        assert_eq!(implicit_intersection(&cell(0, 5, 2), &r), None);
    }
}
