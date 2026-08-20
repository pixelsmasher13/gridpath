//! Tests for the surgical save pipeline.
//!
//! The core invariant everywhere: entries the patch has no reason to touch
//! come out byte-identical, in the original order.

use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Read, Write};

use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use super::apply_patch_json;

// ---------------------------------------------------------------------------
// fixture
// ---------------------------------------------------------------------------

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;

const WORKBOOK: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Alpha" sheetId="1" r:id="rId1"/><sheet name="Beta &amp; Co" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="OLD_NAME">Alpha!$A$1</definedName></definedNames><calcPr calcId="191029"/></workbook>"#;

const WORKBOOK_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>"#;

const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>"#;

/// Sheet1 "Alpha": values, a formula with cached v, a shared-formula group
/// (master B2, members C2/D2), a styled cell on row 4, a gap at row 3.
const SHEET1: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:D4"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData><row r="1" spans="1:4"><c r="A1"><v>1</v></c><c r="B1"><f>A1*10</f><v>10</v></c><c r="D1" t="s"><v>0</v></c></row><row r="2"><c r="A2"><v>5</v></c><c r="B2"><f t="shared" ref="B2:D2" si="0">A2*2</f><v>10</v></c><c r="C2"><f t="shared" si="0"/><v>20</v></c><c r="D2"><f t="shared" si="0"/><v>30</v></c></row><row r="4" ht="20" customHeight="1"><c r="A4" s="1"><v>99</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"#;

/// Sheet2 "Beta & Co": existing cols, merges, autofilter.
const SHEET2: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C3"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><cols><col min="1" max="3" width="12" style="1" customWidth="1"/></cols><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><autoFilter ref="A1:C3"/><mergeCells count="1"><mergeCell ref="B2:C2"/></mergeCells><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"#;

const CALC_CHAIN: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B1" i="1"/></calcChain>"#;

pub(crate) fn fixture() -> Vec<u8> {
    let mut w = ZipWriter::new(Cursor::new(Vec::new()));
    let opt = SimpleFileOptions::default();
    let mut add = |name: &str, data: &[u8]| {
        w.start_file(name, opt).unwrap();
        w.write_all(data).unwrap();
    };
    add("[Content_Types].xml", CONTENT_TYPES.as_bytes());
    add("_rels/.rels", ROOT_RELS.as_bytes());
    add("xl/workbook.xml", WORKBOOK.as_bytes());
    add("xl/_rels/workbook.xml.rels", WORKBOOK_RELS.as_bytes());
    add("xl/styles.xml", STYLES.as_bytes());
    add("xl/worksheets/sheet1.xml", SHEET1.as_bytes());
    add("xl/worksheets/sheet2.xml", SHEET2.as_bytes());
    add("xl/calcChain.xml", CALC_CHAIN.as_bytes());
    // Binary blob standing in for media/customProperty parts that must
    // survive untouched.
    add("xl/media/blob.bin", &[0u8, 159, 146, 150, 255, 0, 1, 2]);
    add("docProps/custom.xml", b"<props><p name=\"secret\"/></props>");
    w.finish().unwrap().into_inner()
}

fn entries(bytes: &[u8]) -> Vec<(String, Vec<u8>)> {
    let mut ar = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut out = Vec::new();
    for i in 0..ar.len() {
        let mut f = ar.by_index(i).unwrap();
        let mut data = Vec::new();
        f.read_to_end(&mut data).unwrap();
        out.push((f.name().to_string(), data));
    }
    out
}

fn entry(bytes: &[u8], name: &str) -> Vec<u8> {
    entries(bytes)
        .into_iter()
        .find(|(n, _)| n == name)
        .unwrap_or_else(|| panic!("entry {name} missing"))
        .1
}

fn entry_str(bytes: &[u8], name: &str) -> String {
    String::from_utf8(entry(bytes, name)).unwrap()
}

fn patch(json: &str) -> String {
    format!(r#"{{"version":1,{json}}}"#)
}

/// Assert every entry except `touched` is byte-identical and order-preserved.
fn assert_untouched_identical(before: &[u8], after: &[u8], touched: &[&str], removed: &[&str]) {
    let b = entries(before);
    let a = entries(after);
    let a_map: HashMap<&str, &Vec<u8>> = a.iter().map(|(n, d)| (n.as_str(), d)).collect();
    for (name, data) in &b {
        if removed.contains(&name.as_str()) {
            assert!(!a_map.contains_key(name.as_str()), "{name} should be removed");
            continue;
        }
        let out = a_map
            .get(name.as_str())
            .unwrap_or_else(|| panic!("{name} missing from output"));
        if touched.contains(&name.as_str()) {
            continue;
        }
        assert_eq!(&data, out, "{name} must be byte-identical");
    }
    // Order of surviving entries is preserved (brand-new entries — sheet
    // creates — append after them and are ignored here).
    let b_order: Vec<&String> = b
        .iter()
        .map(|(n, _)| n)
        .filter(|n| !removed.contains(&n.as_str()))
        .collect();
    let a_order: Vec<&String> = a
        .iter()
        .map(|(n, _)| n)
        .filter(|n| b.iter().any(|(bn, _)| &bn == n))
        .collect();
    assert_eq!(b_order, a_order, "entry order must be preserved");
}

// ---------------------------------------------------------------------------
// tests
// ---------------------------------------------------------------------------

#[test]
fn empty_patch_returns_input_bytes() {
    let base = fixture();
    let out = apply_patch_json(&base, r#"{"version":1}"#).unwrap();
    assert_eq!(base, out);
}

#[test]
fn unknown_version_rejected() {
    let base = fixture();
    assert!(apply_patch_json(&base, r#"{"version":99}"#).is_err());
}

#[test]
fn value_edit_touches_only_expected_parts() {
    let base = fixture();
    let p = patch(
        r#""sheets":[{"name":"Alpha","cells":[{"r":0,"c":0,"v":{"t":"n","n":42.5}}]}]"#,
    );
    let out = apply_patch_json(&base, &p).unwrap();

    assert_untouched_identical(
        &base,
        &out,
        &[
            "xl/worksheets/sheet1.xml",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "[Content_Types].xml",
        ],
        &["xl/calcChain.xml"],
    );

    let sheet = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(sheet.contains(r#"<c r="A1"><v>42.5</v></c>"#), "sheet: {sheet}");

    let wb = entry_str(&out, "xl/workbook.xml");
    assert!(wb.contains(r#"fullCalcOnLoad="1""#), "workbook: {wb}");
    // Existing defined name untouched.
    assert!(wb.contains("OLD_NAME"));

    let ct = entry_str(&out, "[Content_Types].xml");
    assert!(!ct.contains("calcChain"), "content types keeps calcChain: {ct}");
    let rels = entry_str(&out, "xl/_rels/workbook.xml.rels");
    assert!(!rels.contains("calcChain"), "rels keeps calcChain: {rels}");
}

#[test]
fn formula_written_without_cached_value() {
    let base = fixture();
    let p = patch(
        r#""sheets":[{"name":"Alpha","cells":[{"r":0,"c":1,"f":"=SUM(A1:A2)"}]}]"#,
    );
    let out = apply_patch_json(&base, &p).unwrap();
    let sheet = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(
        sheet.contains(r#"<c r="B1"><f>SUM(A1:A2)</f></c>"#),
        "sheet: {sheet}"
    );
    assert!(!sheet.contains(r#"<c r="B1"><f>SUM(A1:A2)</f><v>"#));
}

/// `f` + `v` in one cell patch = formula with its evaluated result cached,
/// typed per OOXML (`t="str"` strings, `t="b"` bools, `t="e"` errors,
/// numbers untyped) so reopening in-app skips the full recalculation.
#[test]
fn formula_with_cached_value_written_and_typed() {
    let base = fixture();
    let p = patch(
        r##""sheets":[{"name":"Alpha","cells":[
            {"r":0,"c":1,"f":"=SUM(A1:A2)","v":{"t":"n","n":6}},
            {"r":4,"c":0,"f":"=CONCAT(D1,\"!\")","v":{"t":"s","s":"hello!"}},
            {"r":4,"c":1,"f":"=A1>0","v":{"t":"b","b":true}},
            {"r":4,"c":2,"f":"=1/0","v":{"t":"e","e":"#DIV/0!"}}
        ]}]"##,
    );
    let out = apply_patch_json(&base, &p).unwrap();
    let sheet = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(
        sheet.contains(r#"<c r="B1"><f>SUM(A1:A2)</f><v>6</v></c>"#),
        "number cache: {sheet}"
    );
    assert!(
        sheet.contains(r#"<c r="A5" t="str"><f>CONCAT(D1,&quot;!&quot;)</f><v>hello!</v></c>"#),
        "string cache: {sheet}"
    );
    assert!(
        sheet.contains(r#"<c r="B5" t="b"><f>A1&gt;0</f><v>1</v></c>"#),
        "bool cache: {sheet}"
    );
    assert!(
        sheet.contains(r#"<c r="C5" t="e"><f>1/0</f><v>#DIV/0!</v></c>"#),
        "error cache: {sheet}"
    );
}

#[test]
fn new_rows_and_cells_land_in_order() {
    let base = fixture();
    let p = patch(
        r#""sheets":[{"name":"Alpha","cells":[
            {"r":2,"c":0,"v":{"t":"n","n":3}},
            {"r":9,"c":3,"v":{"t":"s","s":"tail"}},
            {"r":0,"c":2,"v":{"t":"n","n":7}}
        ]}]"#,
    );
    let out = apply_patch_json(&base, &p).unwrap();
    let sheet = entry_str(&out, "xl/worksheets/sheet1.xml");

    // Row 3 spliced between rows 2 and 4; row 10 appended at the end.
    let pos = |needle: &str| sheet.find(needle).unwrap_or_else(|| panic!("missing {needle}"));
    assert!(pos(r#"<row r="2""#) < pos(r#"<row r="3">"#));
    assert!(pos(r#"<row r="3">"#) < pos(r#"<row r="4""#));
    assert!(pos(r#"<row r="4""#) < pos(r#"<row r="10">"#));
    assert!(sheet.contains(r#"<c r="A3"><v>3</v></c>"#));
    // C1 slots between B1 and D1 within row 1.
    assert!(pos(r#"<c r="B1">"#) < pos(r#"<c r="C1">"#));
    assert!(pos(r#"<c r="C1">"#) < pos(r#"<c r="D1""#));
    // Inline string cell.
    assert!(sheet.contains(r#"<c r="D10" t="inlineStr"><is><t xml:space="preserve">tail</t></is></c>"#));
    // Row 1 originally had spans; modified rows drop it rather than lie.
    assert!(!sheet.contains("spans"));
}

#[test]
fn inline_strings_escape_xml_and_control_chars() {
    let base = fixture();
    let p = patch(
        r#""sheets":[{"name":"Alpha","cells":[{"r":0,"c":0,"v":{"t":"s","s":"a<b&c \u0001 _x0001_ ok"}}]}]"#,
    );
    let out = apply_patch_json(&base, &p).unwrap();
    let sheet = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(
        sheet.contains("a&lt;b&amp;c _x0001_ _x005F_x0001_ ok"),
        "sheet: {sheet}"
    );
}

#[test]
fn clear_keeps_style() {
    let base = fixture();
    let p = patch(r#""sheets":[{"name":"Alpha","cells":[{"r":3,"c":0,"clear":true}]}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    let sheet = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(sheet.contains(r#"<c r="A4" s="1"/>"#), "sheet: {sheet}");
    assert!(!sheet.contains(r#"<c r="A4" s="1"><v>99</v>"#));
}

#[test]
fn shared_formula_group_materializes_when_member_edited() {
    let base = fixture();
    // Overwrite member C2 with a plain value.
    let p = patch(
        r#""sheets":[{"name":"Alpha","cells":[{"r":1,"c":2,"v":{"t":"n","n":123}}]}]"#,
    );
    let out = apply_patch_json(&base, &p).unwrap();
    let sheet = entry_str(&out, "xl/worksheets/sheet1.xml");

    // Master B2: shared stub expanded to plain formula, cached value kept.
    assert!(sheet.contains(r#"<c r="B2"><f>A2*2</f><v>10</v></c>"#), "sheet: {sheet}");
    // C2: replaced by the value.
    assert!(sheet.contains(r#"<c r="C2"><v>123</v></c>"#));
    // D2: translated (+2 cols from master), cached value kept.
    assert!(sheet.contains(r#"<c r="D2"><f>C2*2</f><v>30</v></c>"#));
    assert!(!sheet.contains("t=\"shared\""));
}

#[test]
fn style_patch_appends_without_renumbering() {
    let base = fixture();
    let p = patch(
        r##""sheets":[{"name":"Alpha","styles":[{"r":0,"c":0,"bold":true,"backgroundColor":"#1F4E79"}]}]"##,
    );
    let out = apply_patch_json(&base, &p).unwrap();

    let styles = entry_str(&out, "xl/styles.xml");
    // Bold font already exists (index 1) → deduped, not appended. The solid
    // fill and the composed xf are new.
    assert!(styles.contains(r#"<fonts count="2">"#), "styles: {styles}");
    assert!(styles.contains(r#"<fills count="3">"#));
    assert!(styles.contains(r#"<cellXfs count="3">"#));
    assert!(styles.contains(r#"<fgColor rgb="FF1F4E79"/>"#));

    let sheet = entry_str(&out, "xl/worksheets/sheet1.xml");
    // A1 got the new xf (index 2) and kept its value.
    assert!(sheet.contains(r#"<c r="A1" s="2"><v>1</v></c>"#), "sheet: {sheet}");
}

#[test]
fn style_patch_composes_from_existing_cell_style() {
    let base = fixture();
    // A4 has s=1 (bold font). Adding italic must produce a bold+italic font.
    let p = patch(r#""sheets":[{"name":"Alpha","styles":[{"r":3,"c":0,"italic":true}]}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    let styles = entry_str(&out, "xl/styles.xml");
    assert!(
        styles.contains(r#"<font><i/><b/><sz val="11"/><name val="Calibri"/></font>"#)
            || styles.contains(r#"<font><b/><i/>"#),
        "derived font must keep bold: {styles}"
    );
}

#[test]
fn merge_and_unmerge() {
    let base = fixture();
    let p = patch(
        r#""sheets":[{"name":"Beta & Co","merges":[{"range":"B2:C2","merge":false},{"range":"A1:A2","merge":true}]}]"#,
    );
    let out = apply_patch_json(&base, &p).unwrap();
    let sheet = entry_str(&out, "xl/worksheets/sheet2.xml");
    assert!(sheet.contains(r#"<mergeCells count="1"><mergeCell ref="A1:A2"/></mergeCells>"#), "sheet: {sheet}");
    // Replaced element stays at its original position (after autoFilter).
    let pos = |needle: &str| sheet.find(needle).unwrap();
    assert!(pos("autoFilter") < pos("mergeCells"));
    assert!(pos("mergeCells") < pos("pageMargins"));
}

#[test]
fn overlapping_merge_is_rejected() {
    let base = fixture();
    // B2:C2 already merged in sheet2; A2:B2 overlaps it.
    let p = patch(
        r#""sheets":[{"name":"Beta & Co","merges":[{"range":"A2:B2","merge":true}]}]"#,
    );
    assert!(apply_patch_json(&base, &p).is_err());
}

#[test]
fn col_width_splits_existing_range() {
    let base = fixture();
    let p = patch(r#""sheets":[{"name":"Beta & Co","colWidths":[{"c":1,"chars":30}]}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    let sheet = entry_str(&out, "xl/worksheets/sheet2.xml");
    // 1..3 splits into 1, 2 (patched), 3 — style preserved on all parts.
    assert!(sheet.contains(r#"<col min="1" max="1" width="12" style="1" customWidth="1"/>"#), "sheet: {sheet}");
    assert!(sheet.contains(r#"<col min="2" max="2" style="1" width="30" customWidth="1"/>"#));
    assert!(sheet.contains(r#"<col min="3" max="3" width="12" style="1" customWidth="1"/>"#));
}

#[test]
fn freeze_panes_written() {
    let base = fixture();
    let p = patch(r#""sheets":[{"name":"Alpha","freeze":{"rows":1,"cols":2}}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    let sheet = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(
        sheet.contains(r#"<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>"#),
        "sheet: {sheet}"
    );
}

#[test]
fn freeze_on_new_sheet_lands_in_schema_order() {
    // A created sheet has no <sheetViews>; a freeze must insert it before
    // cols/sheetData (CT_Worksheet sequence), not at worksheet close —
    // Excel repairs sheets with trailing sheetViews.
    let base = structural::fixture();
    let p = patch(
        r##""sheetOps":[{"op":"create","name":"Notes"}],"sheets":[{"name":"Notes","freeze":{"rows":4,"cols":1},"colWidths":[{"c":0,"chars":33}],"cells":[{"r":0,"c":0,"v":{"t":"s","s":"title"}}]}]"##,
    );
    let out = apply_patch_json(&base, &p).unwrap();
    let ws = entry_str(&out, "xl/worksheets/sheet3.xml");
    let views = ws.find("<sheetViews>").expect("sheetViews present");
    let cols = ws.find("<cols>").expect("cols present");
    let data = ws.find("<sheetData").expect("sheetData present");
    assert!(views < cols && cols < data, "bad element order: {ws}");
    assert!(ws.contains(
        r#"<pane xSplit="1" ySplit="4" topLeftCell="B5" activePane="bottomRight" state="frozen"/>"#
    ));
}

#[test]
fn auto_filter_set_and_clear() {
    let base = fixture();
    let set = patch(r#""sheets":[{"name":"Beta & Co","autoFilter":"A1:B9"}]"#);
    let out = apply_patch_json(&base, &set).unwrap();
    let sheet = entry_str(&out, "xl/worksheets/sheet2.xml");
    assert!(sheet.contains(r#"<autoFilter ref="A1:B9"/>"#));
    assert!(!sheet.contains(r#"<autoFilter ref="A1:C3"/>"#));

    let clear = patch(r#""sheets":[{"name":"Beta & Co","autoFilter":null}]"#);
    let out = apply_patch_json(&base, &clear).unwrap();
    let sheet = entry_str(&out, "xl/worksheets/sheet2.xml");
    assert!(!sheet.contains("autoFilter"));
}

#[test]
fn defined_names_upsert() {
    let base = fixture();
    let p = patch(
        r#""definedNames":[{"name":"OLD_NAME","ref":"Alpha!$B$9"},{"name":"WACC","ref":"'Beta &amp; Co'!$C$4"}]"#,
    );
    let out = apply_patch_json(&base, &p).unwrap();
    let wb = entry_str(&out, "xl/workbook.xml");
    assert!(wb.contains(r#"<definedName name="OLD_NAME">Alpha!$B$9</definedName>"#), "wb: {wb}");
    assert!(wb.contains(r#"<definedName name="WACC">"#));
    // No content changed → calcChain untouched, no fullCalcOnLoad.
    assert!(!wb.contains("fullCalcOnLoad"));
    let names: HashSet<String> = entries(&out).into_iter().map(|(n, _)| n).collect();
    assert!(names.contains("xl/calcChain.xml"));
}

#[test]
fn missing_sheet_is_an_error() {
    let base = fixture();
    let p = patch(r#""sheets":[{"name":"Nope","cells":[{"r":0,"c":0,"v":{"t":"n","n":1}}]}]"#);
    assert!(apply_patch_json(&base, &p).is_err());
}

// ---------------------------------------------------------------------------
// structural ops: sheet create / rename / delete, row-col shifts
// ---------------------------------------------------------------------------

/// A workbook with the carriers the structural engine must protect:
/// cross-sheet formulas, defined names, a chart, comments + VML, a drawing,
/// merges, CF, DV, and a hyperlink.
mod structural {
    use super::*;

    const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#;

    const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;

    const WORKBOOK: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="1"/></bookViews><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Calc" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="DATA_RANGE">Data!$A$2:$A$6</definedName></definedNames><calcPr calcId="191029"/></workbook>"#;

    const WORKBOOK_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;

    const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>"#;

    /// "Data": 6 rows of values, a shared group on B2:B4, a merge, CF, DV,
    /// a hyperlink, and rels to comments/VML/drawing.
    const SHEET1: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:D6"/><sheetViews><sheetView workbookViewId="0"><selection activeCell="A4" sqref="A4"/></sheetView></sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="C1"><f>A1*2</f><v>2</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" ref="B2:B4" si="0">A2*10</f><v>20</v></c></row><row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>30</v></c></row><row r="4"><c r="A4"><v>4</v></c><c r="B4"><f t="shared" si="0"/><v>40</v></c></row><row r="5"><c r="A5"><v>5</v></c></row><row r="6"><c r="A6"><v>6</v></c></row></sheetData><mergeCells count="2"><mergeCell ref="C2:D2"/><mergeCell ref="C5:D5"/></mergeCells><conditionalFormatting sqref="A1:A6"><cfRule type="expression" dxfId="0" priority="1"><formula>A1&gt;3</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="list" sqref="D1:D6"><formula1>"a,b"</formula1></dataValidation></dataValidations><hyperlinks><hyperlink ref="A5" r:id="rId4"/></hyperlinks><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/><legacyDrawing r:id="rId2"/></worksheet>"#;

    const SHEET1_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>"#;

    /// "Calc": cross-sheet formulas into Data.
    const SHEET2: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A3"/><sheetData><row r="1"><c r="A1"><f>Data!B3+1</f><v>31</v></c></row><row r="2"><c r="A2"><f>'Data'!$A$5</f><v>5</v></c></row><row r="3"><c r="A3"><f>SUM(Data!A2:A5)</f><v>14</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"#;

    const COMMENTS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>GP</author></authors><commentList><comment ref="A1" authorId="0"><text><t>keep me</t></text></comment><comment ref="A4" authorId="0"><text><t>note on A4</t></text></comment></commentList></comments>"#;

    const VML: &str = r##"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="_x0000_s1026" type="#_x0000_t202" style="position:absolute"><x:ClientData ObjectType="Note"><x:Anchor>1, 15, 0, 2, 3, 15, 1, 4</x:Anchor><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape><v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute"><x:ClientData ObjectType="Note"><x:Anchor>1, 15, 2, 10, 3, 15, 5, 4</x:Anchor><x:Row>3</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>"##;

    const DRAWING: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;

    const CHART: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:cat><c:strRef><c:f>Data!$A$2:$A$6</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$6</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;

    pub fn fixture() -> Vec<u8> {
        let mut w = ZipWriter::new(Cursor::new(Vec::new()));
        let opt = SimpleFileOptions::default();
        let mut add = |name: &str, data: &[u8]| {
            w.start_file(name, opt).unwrap();
            w.write_all(data).unwrap();
        };
        add("[Content_Types].xml", CONTENT_TYPES.as_bytes());
        add("_rels/.rels", ROOT_RELS.as_bytes());
        add("xl/workbook.xml", WORKBOOK.as_bytes());
        add("xl/_rels/workbook.xml.rels", WORKBOOK_RELS.as_bytes());
        add("xl/styles.xml", STYLES.as_bytes());
        add("xl/worksheets/sheet1.xml", SHEET1.as_bytes());
        add("xl/worksheets/_rels/sheet1.xml.rels", SHEET1_RELS.as_bytes());
        add("xl/worksheets/sheet2.xml", SHEET2.as_bytes());
        add("xl/comments1.xml", COMMENTS.as_bytes());
        add("xl/drawings/drawing1.xml", DRAWING.as_bytes());
        add("xl/drawings/vmlDrawing1.vml", VML.as_bytes());
        add("xl/charts/chart1.xml", CHART.as_bytes());
        add("xl/media/blob.bin", &[7u8, 0, 255, 3]);
        w.finish().unwrap().into_inner()
    }

    pub fn fixture_with(extra: &[(&str, &[u8])]) -> Vec<u8> {
        let base = fixture();
        let mut ar = ZipArchive::new(Cursor::new(base.as_slice())).unwrap();
        let mut w = ZipWriter::new(Cursor::new(Vec::new()));
        for i in 0..ar.len() {
            let raw = ar.by_index_raw(i).unwrap();
            w.raw_copy_file(raw).unwrap();
        }
        for (name, data) in extra {
            w.start_file(*name, SimpleFileOptions::default()).unwrap();
            w.write_all(data).unwrap();
        }
        w.finish().unwrap().into_inner()
    }
}

#[test]
fn sheet_create_is_additive_and_patchable() {
    let base = structural::fixture();
    let p = patch(
        r##""sheetOps":[{"op":"create","name":"Notes","tabColor":"#3366CC"}],"sheets":[{"name":"Notes","cells":[{"r":0,"c":0,"v":{"t":"s","s":"hello"}}]}]"##,
    );
    let out = apply_patch_json(&base, &p).unwrap();

    let wb = entry_str(&out, "xl/workbook.xml");
    assert!(wb.contains(r#"name="Notes""#), "workbook.xml: {wb}");
    assert!(wb.contains(r#"sheetId="3""#));

    let rels = entry_str(&out, "xl/_rels/workbook.xml.rels");
    assert!(rels.contains("worksheets/sheet3.xml"), "rels: {rels}");

    let ct = entry_str(&out, "[Content_Types].xml");
    assert!(ct.contains("/xl/worksheets/sheet3.xml"), "ct: {ct}");

    let ws = entry_str(&out, "xl/worksheets/sheet3.xml");
    assert!(ws.contains("hello"), "new sheet content: {ws}");
    assert!(ws.contains(r#"tabColor rgb="FF3366CC""#), "tab color: {ws}");

    // Everything pre-existing survives byte-identical except the three
    // registry parts (a create has no content changes → no calc scrub).
    assert_untouched_identical(
        &base,
        &out,
        &["xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"],
        &[],
    );
}

#[test]
fn sheet_create_is_idempotent() {
    let base = structural::fixture();
    let p = patch(r#""sheetOps":[{"op":"create","name":"Data"}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    assert_untouched_identical(&base, &out, &[], &[]);
}

#[test]
fn sheet_rename_rewrites_cross_sheet_refs() {
    let base = structural::fixture();
    let p = patch(r#""sheetOps":[{"op":"rename","oldName":"Data","newName":"My Data"}]"#);
    let out = apply_patch_json(&base, &p).unwrap();

    let wb = entry_str(&out, "xl/workbook.xml");
    assert!(wb.contains(r#"name="My Data""#), "{wb}");
    assert!(wb.contains("'My Data'!$A$2:$A$6"), "defined name: {wb}");

    let s2 = entry_str(&out, "xl/worksheets/sheet2.xml");
    assert!(s2.contains("'My Data'!B3+1"), "{s2}");
    assert!(s2.contains("'My Data'!$A$5"), "{s2}");
    assert!(s2.contains("SUM('My Data'!A2:A5)"), "{s2}");

    let chart = entry_str(&out, "xl/charts/chart1.xml");
    assert!(chart.contains("'My Data'!$B$2:$B$6"), "{chart}");

    // Comments/VML/drawing/styles/media untouched.
    assert_untouched_identical(
        &base,
        &out,
        &[
            "xl/workbook.xml",
            "xl/worksheets/sheet2.xml",
            "xl/charts/chart1.xml",
        ],
        &[],
    );
}

#[test]
fn sheet_delete_fails_closed_when_referenced() {
    let base = structural::fixture();
    let p = patch(r#""sheetOps":[{"op":"delete","name":"Data"}]"#);
    let err = apply_patch_json(&base, &p).unwrap_err();
    assert!(
        err.to_string().contains("unsupported"),
        "expected guarded failure, got: {err}"
    );
}

#[test]
fn sheet_delete_removes_unreferenced_sheet() {
    let base = structural::fixture();
    let p = patch(r#""sheetOps":[{"op":"delete","name":"Calc"}]"#);
    let out = apply_patch_json(&base, &p).unwrap();

    let wb = entry_str(&out, "xl/workbook.xml");
    assert!(!wb.contains(r#"name="Calc""#), "{wb}");
    // activeTab pointed at the deleted index → clamped.
    assert!(wb.contains(r#"activeTab="0""#), "{wb}");

    let rels = entry_str(&out, "xl/_rels/workbook.xml.rels");
    assert!(!rels.contains("sheet2.xml"), "{rels}");

    let ct = entry_str(&out, "[Content_Types].xml");
    assert!(!ct.contains("/xl/worksheets/sheet2.xml"), "{ct}");

    assert_untouched_identical(
        &base,
        &out,
        &["xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"],
        &["xl/worksheets/sheet2.xml"],
    );
}

#[test]
fn insert_rows_shifts_package_wide() {
    let base = structural::fixture();
    // Two rows before row 3 (0-based index 2).
    let p = patch(r#""rowColOps":[{"op":"insertRows","sheet":"Data","before":2,"count":2}]"#);
    let out = apply_patch_json(&base, &p).unwrap();

    let s1 = entry_str(&out, "xl/worksheets/sheet1.xml");
    // Rows 3..6 renumbered to 5..8; rows 1-2 untouched.
    assert!(s1.contains(r#"<row r="5"><c r="A5"><v>3</v></c>"#), "{s1}");
    assert!(s1.contains(r#"<row r="8"><c r="A8"><v>6</v></c>"#), "{s1}");
    assert!(s1.contains(r#"<c r="A2"><v>2</v></c>"#), "{s1}");
    // Shared group materialized: every member has a plain formula now.
    assert!(!s1.contains("t=\"shared\""), "{s1}");
    assert!(s1.contains("<f>A2*10</f>"), "master formula: {s1}");
    assert!(s1.contains("<f>A5*10</f>"), "member at old row 3: {s1}");
    assert!(s1.contains("<f>A6*10</f>"), "member at old row 4: {s1}");
    // Ranges: dimension, merge below the insert, CF, DV, selection.
    assert!(s1.contains(r#"dimension ref="A1:D8""#), "{s1}");
    assert!(s1.contains(r#"mergeCell ref="C2:D2""#), "{s1}");
    assert!(s1.contains(r#"mergeCell ref="C7:D7""#), "{s1}");
    assert!(s1.contains(r#"sqref="A1:A8""#), "CF sqref: {s1}");
    assert!(s1.contains(r#"sqref="D1:D8""#), "DV sqref: {s1}");
    assert!(s1.contains(r#"hyperlink ref="A7""#), "{s1}");
    assert!(s1.contains(r#"activeCell="A6""#), "{s1}");

    // Cross-sheet formulas on Calc.
    let s2 = entry_str(&out, "xl/worksheets/sheet2.xml");
    assert!(s2.contains("Data!B5+1"), "{s2}");
    assert!(s2.contains("'Data'!$A$7"), "{s2}");
    assert!(s2.contains("SUM(Data!A2:A7)"), "{s2}");

    // Defined name + chart series.
    let wb = entry_str(&out, "xl/workbook.xml");
    assert!(wb.contains("Data!$A$2:$A$8"), "{wb}");
    assert!(wb.contains("fullCalcOnLoad"), "{wb}");
    let chart = entry_str(&out, "xl/charts/chart1.xml");
    assert!(chart.contains("Data!$A$2:$A$8"), "{chart}");
    assert!(chart.contains("Data!$B$2:$B$8"), "{chart}");

    // Comment + VML + drawing anchors.
    let comments = entry_str(&out, "xl/comments1.xml");
    assert!(comments.contains(r#"ref="A6""#), "{comments}");
    let vml = entry_str(&out, "xl/drawings/vmlDrawing1.vml");
    assert!(vml.contains("<x:Row>5</x:Row>"), "{vml}");
    assert!(vml.contains("1, 15, 4, 10, 3, 15, 7, 4"), "{vml}");
    let drawing = entry_str(&out, "xl/drawings/drawing1.xml");
    assert!(drawing.contains("<xdr:row>6</xdr:row>"), "{drawing}");
    assert!(drawing.contains("<xdr:row>14</xdr:row>"), "{drawing}");

    // Styles and media must be byte-identical.
    let touched = [
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/sheet2.xml",
        "xl/workbook.xml",
        "xl/charts/chart1.xml",
        "xl/comments1.xml",
        "xl/drawings/vmlDrawing1.vml",
        "xl/drawings/drawing1.xml",
    ];
    assert_untouched_identical(&base, &out, &touched, &[]);
}

#[test]
fn delete_rows_shrinks_and_refs_deleted_cells() {
    let base = structural::fixture();
    // Delete rows 2-3 (0-based start 1, count 2).
    let p = patch(r#""rowColOps":[{"op":"deleteRows","sheet":"Data","start":1,"count":2}]"#);
    let out = apply_patch_json(&base, &p).unwrap();

    let s1 = entry_str(&out, "xl/worksheets/sheet1.xml");
    // Old rows 4,5,6 → 2,3,4. Old rows 2,3 gone (their values 2 and 3).
    assert!(s1.contains(r#"<row r="2"><c r="A2"><v>4</v></c>"#), "{s1}");
    assert!(!s1.contains("<v>3</v>"), "deleted row content gone: {s1}");
    // Merge C2:D2 was inside the deleted block → dropped; C5:D5 → C3:D3.
    assert!(!s1.contains("C2:D2"), "{s1}");
    assert!(s1.contains(r#"mergeCell ref="C3:D3""#), "{s1}");
    assert!(s1.contains(r#"mergeCells count="1""#), "{s1}");

    let s2 = entry_str(&out, "xl/worksheets/sheet2.xml");
    // Data!B3 pointed into the deleted block.
    assert!(s2.contains("Data!#REF!+1"), "{s2}");
    assert!(s2.contains("'Data'!$A$3"), "{s2}");
    assert!(s2.contains("SUM(Data!A2:A3)"), "{s2}");

    // The comment sat on A4 (old) → A2 now.
    let comments = entry_str(&out, "xl/comments1.xml");
    assert!(comments.contains(r#"ref="A2""#), "{comments}");
}

#[test]
fn delete_rows_removes_comment_and_vml_note_in_range() {
    let base = structural::fixture();
    // Rows 4-5 (0-based 3..5) include the comment on A4 — Excel semantics:
    // the comment and its VML note box are deleted with the rows.
    let p = patch(r#""rowColOps":[{"op":"deleteRows","sheet":"Data","start":3,"count":2}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    let comments = entry_str(&out, "xl/comments1.xml");
    assert!(!comments.contains("note on A4"), "deleted comment gone: {comments}");
    assert!(comments.contains(r#"ref="A1""#), "untouched comment survives: {comments}");
    assert!(comments.contains("keep me"), "{comments}");
    let vml = entry_str(&out, "xl/drawings/vmlDrawing1.vml");
    assert!(!vml.contains("_x0000_s1025"), "deleted note shape gone: {vml}");
    assert!(vml.contains("_x0000_s1026"), "untouched note shape survives: {vml}");
    assert!(vml.contains("<x:Row>0</x:Row>"), "{vml}");
}

// ---------------------------------------------------------------------------
// Save-time validity gate
// ---------------------------------------------------------------------------
// Each case is a bug class that shipped and made Excel offer to "repair" a
// file. The gate must reject all of them before bytes reach disk.

#[test]
fn validator_rejects_worksheet_schema_violations() {
    use super::validate::validate_worksheet;
    let ok = br#"<worksheet><sheetPr/><sheetViews><sheetView workbookViewId="0"/></sheetViews><cols><col min="1" max="1" width="10"/></cols><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row></sheetData></worksheet>"#;
    assert!(validate_worksheet("ws", ok).is_ok());

    // The Regional ARPU bug: sheetViews after sheetData.
    let bad = br#"<worksheet><sheetPr/><cols><col min="1" max="1" width="10"/></cols><sheetData/><sheetViews><sheetView workbookViewId="0"/></sheetViews></worksheet>"#;
    let e = validate_worksheet("ws", bad).unwrap_err().to_string();
    assert!(e.contains("must come before"), "{e}");

    // Duplicate singleton container.
    let bad = br#"<worksheet><sheetData/><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><mergeCells count="1"><mergeCell ref="C1:D1"/></mergeCells></worksheet>"#;
    let e = validate_worksheet("ws", bad).unwrap_err().to_string();
    assert!(e.contains("duplicate <mergeCells>"), "{e}");

    // Rows out of order.
    let bad = br#"<worksheet><sheetData><row r="5"><c r="A5"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData></worksheet>"#;
    let e = validate_worksheet("ws", bad).unwrap_err().to_string();
    assert!(e.contains("out of order"), "{e}");

    // Cells out of order within a row.
    let bad = br#"<worksheet><sheetData><row r="1"><c r="C1"/><c r="A1"/></row></sheetData></worksheet>"#;
    let e = validate_worksheet("ws", bad).unwrap_err().to_string();
    assert!(e.contains("out of order"), "{e}");

    // Cell ref pointing outside its row.
    let bad = br#"<worksheet><sheetData><row r="1"><c r="A2"/></row></sheetData></worksheet>"#;
    let e = validate_worksheet("ws", bad).unwrap_err().to_string();
    assert!(e.contains("inside row"), "{e}");

    // The ExcelJS cfRule bug: operator on containsErrors.
    let bad = br#"<worksheet><sheetData/><conditionalFormatting sqref="A1"><cfRule type="containsErrors" operator="containsErrors" priority="1"/></conditionalFormatting></worksheet>"#;
    let e = validate_worksheet("ws", bad).unwrap_err().to_string();
    assert!(e.contains("must not carry an operator"), "{e}");

    // Unbalanced tags.
    let bad = br#"<worksheet><sheetData><row r="1"></sheetData></worksheet>"#;
    assert!(validate_worksheet("ws", bad).is_err());
}

#[test]
fn validator_rejects_workbook_violations() {
    use super::validate::validate_workbook;
    let ok = br#"<workbook><sheets><sheet name="A" sheetId="1" id="rId1"/></sheets><definedNames><definedName name="X">A!$B$1</definedName></definedNames><calcPr calcId="1"/></workbook>"#;
    assert!(validate_workbook("wb", ok).is_ok());

    // The ExcelJS gutted-constant bug: empty definedName.
    let bad = br#"<workbook><sheets><sheet name="A" sheetId="1" id="rId1"/></sheets><definedNames><definedName name="FLAG"/></definedNames></workbook>"#;
    let e = validate_workbook("wb", bad).unwrap_err().to_string();
    assert!(e.contains("empty definedName"), "{e}");
    let bad = br#"<workbook><sheets><sheet name="A" sheetId="1" id="rId1"/></sheets><definedNames><definedName name="FLAG">  </definedName></definedNames></workbook>"#;
    assert!(validate_workbook("wb", bad).is_err());

    // definedNames after calcPr (schema sequence).
    let bad = br#"<workbook><sheets><sheet name="A" sheetId="1" id="rId1"/></sheets><calcPr calcId="1"/><definedNames><definedName name="X">A!$B$1</definedName></definedNames></workbook>"#;
    let e = validate_workbook("wb", bad).unwrap_err().to_string();
    assert!(e.contains("must come before"), "{e}");

    // Duplicate sheet name (case-insensitive) and duplicate sheetId.
    let bad = br#"<workbook><sheets><sheet name="A" sheetId="1" id="rId1"/><sheet name="a" sheetId="2" id="rId2"/></sheets></workbook>"#;
    assert!(validate_workbook("wb", bad).is_err());
    let bad = br#"<workbook><sheets><sheet name="A" sheetId="1" id="rId1"/><sheet name="B" sheetId="1" id="rId2"/></sheets></workbook>"#;
    assert!(validate_workbook("wb", bad).is_err());
}

#[test]
fn validator_rejects_rels_violations() {
    use super::validate::validate_rels;
    let ok = br#"<Relationships><Relationship Id="rId1" Type="t" Target="a.xml"/></Relationships>"#;
    assert!(validate_rels("rels", ok).is_ok());
    let bad = br#"<Relationships><Relationship Id="rId1" Type="t" Target="a.xml"/><Relationship Id="rId1" Type="t" Target="b.xml"/></Relationships>"#;
    let e = validate_rels("rels", bad).unwrap_err().to_string();
    assert!(e.contains("duplicate relationship Id"), "{e}");
}

#[test]
fn gate_blocks_dangling_sheet_relationship() {
    // Package-level wiring check: a sheet whose r:id resolves to a part
    // that doesn't exist must fail the whole save.
    let base = fixture();
    let p = patch(r#""sheets":[{"name":"Alpha","cells":[{"r":0,"c":0,"v":{"t":"n","n":1}}]}]"#);

    // Sanity: the honest patch passes the gate.
    assert!(apply_patch_json(&base, &p).is_ok());

    // Corrupt the base's workbook rels to point rId1 at a missing part,
    // then make a patch that rewrites workbook.xml (defined names) so the
    // gate's package checks run against dirty registry state.
    let mut ar = ZipArchive::new(Cursor::new(base.as_slice())).unwrap();
    let mut w = ZipWriter::new(Cursor::new(Vec::new()));
    for i in 0..ar.len() {
        let name = ar.by_index_raw(i).unwrap().name().to_string();
        if name == "xl/_rels/workbook.xml.rels" {
            let mut xml = String::new();
            ar.by_index(i).unwrap().read_to_string(&mut xml).unwrap();
            let xml = xml.replace("worksheets/sheet1.xml", "worksheets/ghost.xml");
            w.start_file(name, SimpleFileOptions::default()).unwrap();
            w.write_all(xml.as_bytes()).unwrap();
        } else {
            w.raw_copy_file(ar.by_index_raw(i).unwrap()).unwrap();
        }
    }
    let corrupted = w.finish().unwrap().into_inner();
    let p = patch(r#""definedNames":[{"name":"N","ref":"'Beta &amp; Co'!$A$1"}]"#);
    let e = apply_patch_json(&corrupted, &p).unwrap_err().to_string();
    assert!(e.contains("missing part"), "{e}");
}

#[test]
fn validate_package_accepts_fixture_and_rejects_corruption() {
    use super::validate::validate_package;
    assert!(validate_package(&fixture()).is_ok());

    // Move sheetViews after sheetData inside sheet1 — the exact shape that
    // made Excel offer to repair NFLXv1.
    let base = fixture();
    let mut ar = ZipArchive::new(Cursor::new(base.as_slice())).unwrap();
    let mut w = ZipWriter::new(Cursor::new(Vec::new()));
    for i in 0..ar.len() {
        let name = ar.by_index_raw(i).unwrap().name().to_string();
        if name == "xl/worksheets/sheet1.xml" {
            let mut xml = String::new();
            ar.by_index(i).unwrap().read_to_string(&mut xml).unwrap();
            let views = r#"<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>"#;
            let moved = xml.replace(views, "").replace("</worksheet>", &format!("{views}</worksheet>"));
            w.start_file(name, SimpleFileOptions::default()).unwrap();
            w.write_all(moved.as_bytes()).unwrap();
        } else {
            w.raw_copy_file(ar.by_index_raw(i).unwrap()).unwrap();
        }
    }
    let corrupted = w.finish().unwrap().into_inner();
    let e = validate_package(&corrupted).unwrap_err().to_string();
    assert!(e.contains("must come before"), "{e}");
}

// Local-only smoke test against a real Canalyst model. Run with:
//   cargo test nflx_local -- --ignored
#[test]
#[ignore]
fn nflx_local_fallback_output_passes_package_validation() {
    // NFLX.xlsx is Excel-native output; NFLXv1.xlsx is a real ExcelJS
    // fallback export (post sanitize/repair). The package validator must
    // not false-positive on either — a failure here means saves would be
    // blocked on legitimate files.
    let home = std::env::var("HOME").unwrap();
    for f in ["NFLX.xlsx", "NFLXv1.xlsx"] {
        let path = format!("{home}/Downloads/{f}");
        let bytes = match std::fs::read(&path) {
            Ok(b) => b,
            Err(_) => continue,
        };
        if let Err(e) = super::validate::validate_package(&bytes) {
            panic!("{f}: {e}");
        }
    }
}

#[test]
#[ignore]
fn nflx_local_delete_comment_rows() {
    let path = std::env::var("HOME").unwrap() + "/Downloads/NFLX.xlsx";
    let base = match std::fs::read(&path) {
        Ok(b) => b,
        Err(_) => return,
    };
    let p = patch(r#""rowColOps":[{"op":"deleteRows","sheet":"Model","start":345,"count":44}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    std::fs::write("/tmp/nflx-comment-delete.xlsx", &out).unwrap();
}

#[test]
fn threaded_comments_shift_and_delete_with_rows() {
    use super::refs::RowColShift;
    let xml = br#"<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="A2" id="{1}" personId="{p}"><text>parent</text></threadedComment><threadedComment ref="A2" id="{2}" parentId="{1}" personId="{p}"><text>reply</text></threadedComment><threadedComment ref="A9" id="{3}" personId="{p}"><text>below</text></threadedComment></ThreadedComments>"#;
    // Delete rows 2-3 (0-based 1..3): the A2 thread (parent + reply) goes,
    // A9 shifts up to A7.
    let shift = RowColShift { rows: true, start: 1, count: 2, insert: false };
    let out = super::row_col::shift_comments(xml, &shift).unwrap();
    let out = String::from_utf8(out).unwrap();
    assert!(!out.contains("parent"), "{out}");
    assert!(!out.contains("reply"), "{out}");
    assert!(out.contains(r#"ref="A7""#), "{out}");
    assert!(out.contains("below"), "{out}");
}

#[test]
fn insert_and_delete_columns_renumber_cells() {
    let base = structural::fixture();
    // Insert one column before B (0-based index 1) on Data.
    let p = patch(r#""rowColOps":[{"op":"insertColumns","sheet":"Data","before":1,"count":1}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    let s1 = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(s1.contains(r#"<c r="A1"><v>1</v></c>"#), "{s1}");
    assert!(s1.contains(r#"<c r="D1">"#), "C1 → D1: {s1}");
    assert!(s1.contains("<f>A2*10</f>"), "unqualified col A ref stays: {s1}");
    assert!(s1.contains(r#"mergeCell ref="D2:E2""#), "{s1}");
    let s2 = entry_str(&out, "xl/worksheets/sheet2.xml");
    assert!(s2.contains("Data!C3+1"), "{s2}");
    let chart = entry_str(&out, "xl/charts/chart1.xml");
    assert!(chart.contains("Data!$C$2:$C$6"), "{chart}");
    let vml = entry_str(&out, "xl/drawings/vmlDrawing1.vml");
    assert!(vml.contains("<x:Column>0</x:Column>"), "{vml}");
    assert!(vml.contains("2, 15, 2, 10, 4, 15, 5, 4"), "{vml}");

    // Delete column B: the shared-formula column disappears; refs into it
    // become #REF!.
    let p = patch(r#""rowColOps":[{"op":"deleteColumns","sheet":"Data","start":1,"count":1}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    let s1 = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(s1.contains(r#"<c r="B1"><f>A1*2</f>"#), "C1 → B1: {s1}");
    assert!(!s1.contains("*10"), "shared column removed: {s1}");
    assert!(s1.contains(r#"mergeCell ref="B2:C2""#), "{s1}");
    let s2 = entry_str(&out, "xl/worksheets/sheet2.xml");
    assert!(s2.contains("Data!#REF!+1"), "{s2}");
    assert!(s2.contains("'Data'!$A$5"), "col A ref unaffected: {s2}");
    let chart = entry_str(&out, "xl/charts/chart1.xml");
    assert!(chart.contains("Data!$A$2:$A$6"), "{chart}");
    assert!(chart.contains("#REF!"), "series into deleted col: {chart}");
    let vml = entry_str(&out, "xl/drawings/vmlDrawing1.vml");
    assert!(vml.contains("1, 15, 2, 10, 2, 15, 5, 4"), "{vml}");

    // Deleting column A removes the comment on A4 and its VML note box
    // (Excel semantics — comments go with their cells).
    let p = patch(r#""rowColOps":[{"op":"deleteColumns","sheet":"Data","start":0,"count":1}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    let comments = entry_str(&out, "xl/comments1.xml");
    assert!(!comments.contains("<comment ref"), "{comments}");
    let vml = entry_str(&out, "xl/drawings/vmlDrawing1.vml");
    assert!(!vml.contains("<v:shape "), "{vml}");
}

#[test]
fn row_col_op_composes_with_cell_patch() {
    let base = structural::fixture();
    // Insert a row at the top, then write into the (post-shift) new row 1.
    let p = patch(
        r#""rowColOps":[{"op":"insertRows","sheet":"Data","before":0,"count":1}],"sheets":[{"name":"Data","cells":[{"r":0,"c":0,"v":{"t":"n","n":42}}]}]"#,
    );
    let out = apply_patch_json(&base, &p).unwrap();
    let s1 = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(s1.contains(r#"<c r="A1"><v>42</v></c>"#), "{s1}");
    // The old A1 landed on A2.
    assert!(s1.contains(r#"<c r="A2"><v>1</v></c>"#), "{s1}");
}

#[test]
fn structural_ops_fail_closed_on_pivots_and_tables() {
    for part in ["xl/pivotTables/pivotTable1.xml", "xl/tables/table1.xml"] {
        let base = structural::fixture_with(&[(part, b"<x/>")]);
        let p = patch(r#""rowColOps":[{"op":"insertRows","sheet":"Data","before":0,"count":1}]"#);
        let err = apply_patch_json(&base, &p).unwrap_err();
        assert!(
            err.to_string().contains("unsupported"),
            "{part}: expected fail-closed, got {err}"
        );
        // Sheet creation stays allowed — it's purely additive.
        let p = patch(r#""sheetOps":[{"op":"create","name":"Notes"}]"#);
        apply_patch_json(&base, &p).unwrap();
    }
}

/// End-to-end bridge for driving the patcher from outside cargo (the
/// frontend patch-builder test): applies GRIDPATH_PATCH_FILE to
/// GRIDPATH_PATCH_BASE and writes GRIDPATH_PATCH_OUT.
#[test]
fn apply_patch_from_env() {
    let (Ok(base), Ok(pfile), Ok(out)) = (
        std::env::var("GRIDPATH_PATCH_BASE"),
        std::env::var("GRIDPATH_PATCH_FILE"),
        std::env::var("GRIDPATH_PATCH_OUT"),
    ) else {
        return;
    };
    let base_bytes = std::fs::read(&base).unwrap();
    let patch_json = std::fs::read_to_string(&pfile).unwrap();
    let result = apply_patch_json(&base_bytes, &patch_json).expect("apply patch");
    std::fs::write(&out, result).unwrap();
}

/// Corpus harness: set GRIDPATH_XLSX_CORPUS to a directory of real .xlsx
/// files. For each, applies a far-away cell edit and asserts every other
/// part survives byte-identical.
#[test]
fn corpus_untouched_parts_survive() {
    let Ok(dir) = std::env::var("GRIDPATH_XLSX_CORPUS") else {
        eprintln!("GRIDPATH_XLSX_CORPUS not set; skipping corpus test");
        return;
    };
    let mut checked = 0;
    for f in std::fs::read_dir(&dir).expect("corpus dir") {
        let path = f.unwrap().path();
        if path.extension().and_then(|e| e.to_str()) != Some("xlsx") {
            continue;
        }
        let base = std::fs::read(&path).unwrap();

        // Empty patch: exact bytes back.
        let out = apply_patch_json(&base, r#"{"version":1}"#).unwrap();
        assert_eq!(base, out, "{path:?} empty patch");

        // Find the first sheet name from the package itself.
        let mut pkg = super::package::Package::open(&base).unwrap();
        let wb = pkg.read_part("xl/workbook.xml").unwrap();
        let rels = pkg.read_part("xl/_rels/workbook.xml.rels").unwrap();
        let sheets = super::workbook_xml::sheet_part_paths(&wb, &rels).unwrap();
        let (name, part) = sheets.first().expect("has sheets").clone();

        let p = format!(
            r#"{{"version":1,"sheets":[{{"name":{},"cells":[{{"r":99998,"c":0,"v":{{"t":"s","s":"gridpath-corpus-probe"}}}}]}}]}}"#,
            serde_json::to_string(&name).unwrap()
        );
        let out = apply_patch_json(&base, &p).unwrap_or_else(|e| panic!("{path:?}: {e}"));

        // Optional: dump patched output for manual Excel/LibreOffice checks.
        if let Ok(dump_dir) = std::env::var("GRIDPATH_CORPUS_DUMP") {
            let name = path.file_stem().unwrap().to_string_lossy();
            let dump = std::path::Path::new(&dump_dir).join(format!("{name}.patched.xlsx"));
            std::fs::write(&dump, &out).unwrap();
            eprintln!("dumped {dump:?}");
        }

        let before = entries(&base);
        let after: HashMap<String, Vec<u8>> = entries(&out).into_iter().collect();
        let mut calc_chain_removed = 0;
        for (n, data) in &before {
            if n == &part || n == "xl/workbook.xml" || n == "[Content_Types].xml" || n == "xl/_rels/workbook.xml.rels" {
                continue;
            }
            if n.ends_with("calcChain.xml") {
                calc_chain_removed += 1;
                assert!(!after.contains_key(n), "{path:?}: calcChain should be gone");
                continue;
            }
            assert_eq!(
                Some(data),
                after.get(n),
                "{path:?}: {n} must survive byte-identical"
            );
        }
        let _ = calc_chain_removed;
        checked += 1;
    }
    assert!(checked > 0, "no .xlsx files in {dir}");
}

// ---------------------------------------------------------------------------
// full-export package validation
// ---------------------------------------------------------------------------

/// ExcelJS emits zip DIRECTORY entries ("xl/") on every write — its internal
/// JSZip creates folder objects. OPC directories are not parts and Excel
/// ignores them, so full-export validation must skip them instead of failing
/// with "new part xl/ has no content type" (which blocked every save of an
/// untitled workbook).
#[test]
fn validate_package_ignores_zip_directory_entries() {
    let base = fixture();
    // Baseline: the fixture itself is a valid package.
    super::validate::validate_package(&base).unwrap();

    // Re-zip with explicit directory entries interleaved, as JSZip writes.
    let mut ar = ZipArchive::new(Cursor::new(base.as_slice())).unwrap();
    let mut w = ZipWriter::new(Cursor::new(Vec::new()));
    w.add_directory("xl", SimpleFileOptions::default()).unwrap();
    w.add_directory("xl/worksheets", SimpleFileOptions::default()).unwrap();
    for i in 0..ar.len() {
        let mut f = ar.by_index(i).unwrap();
        let mut data = Vec::new();
        f.read_to_end(&mut data).unwrap();
        w.start_file(f.name(), SimpleFileOptions::default()).unwrap();
        w.write_all(&data).unwrap();
    }
    let with_dirs = w.finish().unwrap().into_inner();

    super::validate::validate_package(&with_dirs)
        .expect("directory entries must not fail content-type validation");
}
