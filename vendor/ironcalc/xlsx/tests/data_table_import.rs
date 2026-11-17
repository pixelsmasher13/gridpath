#![allow(clippy::unwrap_used)]
#![allow(clippy::panic)]

// GRIDPATH: workbooks with a what-if data table (Data > What-If Analysis >
// Data Table) used to be refused whole ("data table formulas" not
// implemented). They open now: the cells of the table read
// =TABLE(row input, column input), as Excel shows them, and keep the
// values the file was saved with, since this engine does not compute data
// tables.

use std::io::{Cursor, Read, Write};

use ironcalc::base::cell::CellValue;
use ironcalc::base::Model;
use ironcalc::export::save_xlsx_to_writer;
use ironcalc::import::load_from_xlsx_bytes;

/// A workbook saved by the exporter, with `cells` put in place of sheet 1's
/// rows: the way to get a file that holds XML the exporter never writes.
fn workbook_with_rows(rows: &str) -> Vec<u8> {
    let mut model = Model::new_empty("book", "en", "UTC", "en").unwrap();
    model.set_user_input(0, 1, 1, "1".to_string()).unwrap();
    model.evaluate();
    let saved = save_xlsx_to_writer(&model, Cursor::new(Vec::new())).unwrap().into_inner();

    let mut archive = zip::ZipArchive::new(Cursor::new(saved)).unwrap();
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let name = file.name().to_string();
        let mut data = Vec::new();
        file.read_to_end(&mut data).unwrap();
        if name == "xl/worksheets/sheet1.xml" {
            let xml = String::from_utf8(data).unwrap();
            let start = xml.find("<sheetData>").unwrap();
            let end = xml.find("</sheetData>").unwrap();
            data = format!("{}<sheetData>{rows}{}", &xml[..start], &xml[end..]).into_bytes();
        }
        out.start_file(name, zip::write::FileOptions::default()).unwrap();
        out.write_all(&data).unwrap();
    }
    out.finish().unwrap().into_inner()
}

fn load(bytes: &[u8]) -> Model<'static> {
    let workbook = load_from_xlsx_bytes(bytes, "book", "en", "UTC").unwrap();
    let mut model = Model::from_workbook(workbook, "en").unwrap();
    model.evaluate();
    model
}

fn number(model: &Model, row: i32, column: i32) -> f64 {
    match model.get_cell_value_by_index(0, row, column).unwrap() {
        CellValue::Number(n) => n,
        other => panic!("r{row}c{column}: {other:?}"),
    }
}

#[test]
fn a_two_way_data_table_opens_with_its_saved_values() {
    // Inputs in B1 (row input) and B2 (column input); the table is C5:D6.
    let bytes = workbook_with_rows(
        r#"<row r="1"><c r="B1"><v>10</v></c></row><row r="2"><c r="B2"><v>0.3</v></c></row><row r="4"><c r="B4"><f>B1*B2</f><v>3</v></c></row><row r="5"><c r="C5"><f t="dataTable" ref="C5:D6" dt2D="1" dtr="1" r1="B1" r2="B2"/><v>11</v></c><c r="D5"><v>12</v></c></row><row r="6"><c r="C6"><v>21</v></c><c r="D6"><v>22</v></c></row><row r="7"><c r="C7"><v>99</v></c></row>"#,
    );
    let mut model = load(&bytes);

    // Every cell of the block reads as the table, with the values of the file.
    for (row, column, value) in [(5, 3, 11.0), (5, 4, 12.0), (6, 3, 21.0), (6, 4, 22.0)] {
        assert_eq!(model.get_cell_formula(0, row, column).unwrap(), Some("=TABLE($B$1,$B$2)".to_string()));
        assert_eq!(number(&model, row, column), value);
    }
    // The cells around it are ordinary cells.
    assert_eq!(model.get_cell_formula(0, 7, 3).unwrap(), None);
    assert_eq!(model.get_cell_formula(0, 4, 2).unwrap(), Some("=B1*B2".to_string()));
    assert_eq!(number(&model, 4, 2), 3.0);

    // An input changes: ordinary formulas follow, the table keeps what the
    // file had (it is not computed here) and never turns into an error.
    model.set_user_input(0, 1, 2, "20".to_string()).unwrap();
    model.evaluate();
    assert_eq!(number(&model, 4, 2), 6.0);
    assert_eq!(number(&model, 5, 3), 11.0);
    assert_eq!(number(&model, 6, 4), 22.0);

    // Rows inserted above move the inputs the table names, as for any formula.
    model.insert_rows(0, 1, 2).unwrap();
    model.evaluate();
    assert_eq!(model.get_cell_formula(0, 7, 3).unwrap(), Some("=TABLE($B$3,$B$4)".to_string()));
    assert_eq!(number(&model, 7, 3), 11.0);
}

#[test]
fn one_way_data_tables_name_their_single_input() {
    // A column-oriented table (the input varies down a column) and a
    // row-oriented one; an input that was deleted reads #REF!.
    let bytes = workbook_with_rows(
        r#"<row r="1"><c r="B1"><v>10</v></c></row><row r="5"><c r="C5"><f t="dataTable" ref="C5:C6" dt2D="0" dtr="0" r1="B1"/><v>1</v></c><c r="E5"><f t="dataTable" ref="E5:F5" dt2D="0" dtr="1" r1="B1"/><v>3</v></c><c r="F5"><v>4</v></c><c r="H5"><f t="dataTable" ref="H5:H5" dt2D="1" dtr="1" r1="B1" r2="B2" del2="1"/><v>5</v></c></row><row r="6"><c r="C6"><v>2</v></c></row>"#,
    );
    let model = load(&bytes);
    assert_eq!(model.get_cell_formula(0, 6, 3).unwrap(), Some("=TABLE(,$B$1)".to_string()));
    assert_eq!(model.get_cell_formula(0, 5, 6).unwrap(), Some("=TABLE($B$1,)".to_string()));
    assert_eq!(model.get_cell_formula(0, 5, 8).unwrap(), Some("=TABLE($B$1,#REF!)".to_string()));
    assert_eq!(number(&model, 6, 3), 2.0);
    assert_eq!(number(&model, 5, 6), 4.0);
    assert_eq!(number(&model, 5, 8), 5.0);
}
