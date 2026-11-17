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
    // Formulas are written as dynamic-array formulas (see metadata.rs).
    assert!(
        sheet.contains(r#"<c r="B1" cm="1"><f t="array" ref="B1">SUM(A1:A2)</f></c>"#),
        "sheet: {sheet}"
    );
    assert!(!sheet.contains(r#">SUM(A1:A2)</f><v>"#));
}

/// Every formula the patch writes is a dynamic-array formula: `cm` points at
/// the XLDAPR record in xl/metadata.xml and the `<f>` is `t="array"` over
/// the cell itself, so Excel evaluates arrays the way the engine does
/// (no implicit intersection — the legacy mode that blanked every
/// `MATCH(1,(rng="x")*(rng2<>"y"),0)` row-finder). The part is created once
/// and reused by later saves; literal values never trigger it.
#[test]
fn formulas_are_dynamic_array_formulas_with_shared_metadata() {
    let base = fixture();
    assert!(!entries(&base).iter().any(|(n, _)| n == "xl/metadata.xml"));
    // A value-only patch leaves the package without a metadata part.
    let values_only = apply_patch_json(&base, &patch(r#""sheets":[{"name":"Alpha","cells":[{"r":9,"c":0,"v":{"t":"n","n":1}}]}]"#)).unwrap();
    assert!(!entries(&values_only).iter().any(|(n, _)| n == "xl/metadata.xml"));

    let p = patch(r##""sheets":[{"name":"Alpha","cells":[
        {"r":0,"c":1,"f":"=IF($D1=\"\",\"\",INDEX(Beta!A:A,$D1))","v":{"t":"n","n":7}},
        {"r":0,"c":3,"f":"=IFERROR(MATCH(1,(Beta!A1:A9=\"x\")*(Beta!B1:B9<>\"y\"),0),\"\")","v":{"t":"n","n":3}}
    ]}]"##);
    let out = apply_patch_json(&base, &p).unwrap();
    let sheet = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(sheet.contains(r#"<c r="B1" cm="1"><f t="array" ref="B1">IF($D1=&quot;&quot;,&quot;&quot;,INDEX(Beta!A:A,$D1))</f><v>7</v></c>"#), "{sheet}");
    assert!(sheet.contains(r#"<c r="D1" cm="1"><f t="array" ref="D1">"#), "{sheet}");
    let meta = entry_str(&out, "xl/metadata.xml");
    assert!(meta.contains(r#"<metadataType name="XLDAPR""#) && meta.contains(r#"xda:dynamicArrayProperties fDynamic="1""#) && meta.contains(r#"<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>"#), "{meta}");
    let ct = entry_str(&out, "[Content_Types].xml");
    assert!(ct.contains(r#"PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml""#), "{ct}");
    let rels = entry_str(&out, "xl/_rels/workbook.xml.rels");
    assert!(rels.contains(r#"relationships/sheetMetadata" Target="metadata.xml""#), "{rels}");
    // Untouched parts stay byte-identical; only the expected ones change.
    assert_untouched_identical(&base, &out, &["xl/worksheets/sheet1.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels", "xl/workbook.xml"], &["xl/calcChain.xml"]);

    // Second save on the result: the existing record is reused, nothing duplicated.
    let p2 = patch(r#""sheets":[{"name":"Alpha","cells":[{"r":2,"c":1,"f":"=B1*2"}]}]"#);
    let out2 = apply_patch_json(&out, &p2).unwrap();
    let sheet2 = entry_str(&out2, "xl/worksheets/sheet1.xml");
    assert!(sheet2.contains(r#"<c r="B3" cm="1"><f t="array" ref="B3">B1*2</f></c>"#), "{sheet2}");
    assert_eq!(entry_str(&out2, "xl/metadata.xml"), meta);
    assert_eq!(entry_str(&out2, "xl/_rels/workbook.xml.rels").matches("sheetMetadata").count(), 1);
    assert_eq!(entry_str(&out2, "[Content_Types].xml").matches("metadata.xml").count(), 1);
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
        sheet.contains(r#"<c r="B1" cm="1"><f t="array" ref="B1">SUM(A1:A2)</f><v>6</v></c>"#),
        "number cache: {sheet}"
    );
    assert!(
        sheet.contains(r#"<c r="A5" t="str" cm="1"><f t="array" ref="A5">CONCAT(D1,&quot;!&quot;)</f><v>hello!</v></c>"#),
        "string cache: {sheet}"
    );
    assert!(
        sheet.contains(r#"<c r="B5" t="b" cm="1"><f t="array" ref="B5">A1&gt;0</f><v>1</v></c>"#),
        "bool cache: {sheet}"
    );
    assert!(
        sheet.contains(r#"<c r="C5" t="e" cm="1"><f t="array" ref="C5">1/0</f><v>#DIV/0!</v></c>"#),
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
            // A part named in `extra` replaces the fixture's copy.
            if extra.iter().any(|(n, _)| *n == raw.name()) {
                continue;
            }
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
    // made Excel offer to repair a real fallback export.
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

// Local-only smoke tests against private real models, named by environment
// variables (skipped when unset). Run with:
//   GRIDPATH_LOCAL_XLSX=… GRIDPATH_LOCAL_FALLBACK_XLSX=… cargo test local_model -- --ignored
#[test]
#[ignore]
fn local_model_fallback_output_passes_package_validation() {
    // GRIDPATH_LOCAL_XLSX is Excel-native output; GRIDPATH_LOCAL_FALLBACK_XLSX
    // a real ExcelJS fallback export (post sanitize/repair). The package validator must
    // not false-positive on either — a failure here means saves would be
    // blocked on legitimate files.
    for var in ["GRIDPATH_LOCAL_XLSX", "GRIDPATH_LOCAL_FALLBACK_XLSX"] {
        let Ok(path) = std::env::var(var) else { continue };
        let f = var;
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
fn local_model_delete_comment_rows() {
    let Ok(path) = std::env::var("GRIDPATH_LOCAL_XLSX") else { return };
    let base = match std::fs::read(&path) {
        Ok(b) => b,
        Err(_) => return,
    };
    let p = patch(r#""rowColOps":[{"op":"deleteRows","sheet":"Model","start":345,"count":44}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    std::fs::write(std::env::temp_dir().join("local-model-comment-delete.xlsx"), &out).unwrap();
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
fn insert_columns_keeps_what_ends_on_the_last_column() {
    // Real models carry a column format that runs to the last column and
    // names like `Model!$C$1510:$XFD$1533`. Inserting a column used to stop
    // the save with RefOutOfRange; Excel leaves the far end where it is.
    let plain = structural::fixture();
    let sheet1 = entry_str(&plain, "xl/worksheets/sheet1.xml").replace(
        "<sheetData>",
        r#"<cols><col min="1" max="1" width="20" customWidth="1"/><col min="5" max="16384" width="9" style="0"/></cols><sheetData>"#,
    );
    let workbook = entry_str(&plain, "xl/workbook.xml").replace(
        "<definedName name=\"DATA_RANGE\">Data!$A$2:$A$6</definedName>",
        "<definedName name=\"DATA_RANGE\">Data!$A$2:$A$6</definedName><definedName name=\"TO_EDGE\">Data!$C$2:$XFD$6</definedName><definedName name=\"ALL_ROWS\">Data!$A$2:$A$1048576</definedName>",
    );
    assert!(sheet1.contains("16384") && workbook.contains("TO_EDGE"), "fixture not as expected");
    let base = structural::fixture_with(&[
        ("xl/worksheets/sheet1.xml", sheet1.as_bytes()),
        ("xl/workbook.xml", workbook.as_bytes()),
    ]);

    let p = patch(r#""rowColOps":[{"op":"insertColumns","sheet":"Data","before":1,"count":1}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    let s1 = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(s1.contains(r#"<col min="6" max="16384" width="9" style="0"/>"#), "{s1}");
    assert!(s1.contains(r#"<c r="D1">"#), "C1 → D1: {s1}");
    let wb = entry_str(&out, "xl/workbook.xml");
    assert!(wb.contains("Data!$D$2:$XFD$6"), "{wb}");

    // Rows: a name that runs to the last row keeps its end.
    let p = patch(r#""rowColOps":[{"op":"insertRows","sheet":"Data","before":3,"count":2}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    let wb = entry_str(&out, "xl/workbook.xml");
    assert!(wb.contains("Data!$A$2:$A$1048576"), "{wb}");
    assert!(wb.contains("Data!$C$2:$XFD$8"), "{wb}");
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

#[test]
fn hiding_a_grouped_row_keeps_its_outline_level() {
    // Outline groups (Excel's +/- bars) live on <row outlineLevel> / <col outlineLevel>;
    // the viewport grid collapses a group by hiding its rows, which must not drop them.
    let sheet = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A3"/><sheetFormatPr defaultRowHeight="15" outlineLevelRow="1" outlineLevelCol="1"/><cols><col min="2" max="3" width="9" outlineLevel="1"/></cols><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2" outlineLevel="1"><c r="A2"><v>2</v></c></row><row r="3" outlineLevel="1" hidden="1"><c r="A3"><v>3</v></c></row></sheetData></worksheet>"#;
    let base = structural::fixture_with(&[("xl/worksheets/sheet1.xml", sheet)]);
    let wb = entry_str(&base, "xl/workbook.xml");
    let first = crate::structure::list_sheets(wb.as_bytes()).unwrap().into_iter().next().unwrap().0;
    let p = patch(&format!(r#""sheets":[{{"name":"{first}","hiddenRows":[{{"r":1,"hidden":true}},{{"r":2,"hidden":false}}],"hiddenCols":[{{"c":1,"hidden":true}}]}}]"#));
    let out = apply_patch_json(&base, &p).unwrap();
    let xml = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(xml.contains(r#"<row r="2" outlineLevel="1" hidden="1">"#), "{xml}");
    assert!(xml.contains(r#"<row r="3" outlineLevel="1">"#), "{xml}");
    let col_b = xml.find(r#"min="2" max="2""#).map(|i| &xml[i.saturating_sub(5)..(i + 60).min(xml.len())]).unwrap_or("");
    assert!(col_b.contains("outlineLevel=\"1\"") && col_b.contains("hidden=\"1\""), "{xml}");
}

#[test]
fn auto_filter_with_criteria_round_trips() {
    let base = structural::fixture();
    let names: Vec<String> = crate::structure::list_sheets(entry_str(&base, "xl/workbook.xml").as_bytes()).unwrap().into_iter().map(|(n, _, _)| n).collect();
    let first = &names[0];
    let p = patch(&format!(r#""sheets":[{{"name":"{first}","autoFilter":{{"ref":"A1:C9","columns":[{{"colId":0,"values":["North","South"],"blank":true}},{{"colId":2,"custom":{{"and":true,"items":[{{"operator":"greaterThan","val":"5"}},{{"operator":"lessThanOrEqual","val":"10"}}]}}}}]}}}}]"#));
    let out = apply_patch_json(&base, &p).unwrap();
    let parts = crate::workbook_xml::sheet_part_paths(entry_str(&base, "xl/workbook.xml").as_bytes(), entry_str(&base, "xl/_rels/workbook.xml.rels").as_bytes()).unwrap();
    let part = parts.iter().find(|(n, _)| n == first).map(|(_, p)| if p.starts_with("xl/") { p.clone() } else { format!("xl/{p}") }).unwrap();
    let xml = entry_str(&out, &part);
    assert!(xml.contains(r#"<autoFilter ref="A1:C9"><filterColumn colId="0"><filters blank="1"><filter val="North"/><filter val="South"/></filters></filterColumn><filterColumn colId="2"><customFilters and="1"><customFilter operator="greaterThan" val="5"/><customFilter operator="lessThanOrEqual" val="10"/></customFilters></filterColumn></autoFilter>"#), "{xml}");
    // It sits before mergeCells / pageMargins and after sheetData.
    let pos = |t: &str| xml.find(t).unwrap_or(usize::MAX);
    assert!(pos("</sheetData>") < pos("<autoFilter") && pos("</autoFilter>") < pos("<pageMargins"), "{xml}");
    // A bare range still works and replaces the criteria; null clears.
    let p2 = patch(&format!(r#""sheets":[{{"name":"{first}","autoFilter":"A1:C9"}}]"#));
    let out2 = apply_patch_json(&out, &p2).unwrap();
    let xml2 = entry_str(&out2, &part);
    assert!(xml2.contains(r#"<autoFilter ref="A1:C9"/>"#) && !xml2.contains("filterColumn"), "{xml2}");
    let p3 = patch(&format!(r#""sheets":[{{"name":"{first}","autoFilter":null}}]"#));
    let out3 = apply_patch_json(&out2, &p3).unwrap();
    assert!(!entry_str(&out3, &part).contains("<autoFilter"));
}

#[test]
fn duplicate_create_in_one_patch_is_one_sheet() {
    let base = structural::fixture();
    let p = patch(r#""sheetOps":[{"op":"create","name":"Notes"},{"op":"create","name":"Notes"}],"sheets":[{"name":"Notes","cells":[{"r":0,"c":0,"v":{"t":"s","s":"hi"}}]}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    let wb = entry_str(&out, "xl/workbook.xml");
    assert_eq!(wb.matches(r#"name="Notes""#).count(), 1, "{wb}");
    // Creating a sheet the file already has is idempotent too (a patch may
    // be re-applied after a save that did not refresh client state).
    let p2 = patch(r#""sheetOps":[{"op":"create","name":"Notes"}]"#);
    let out2 = apply_patch_json(&out, &p2).unwrap();
    assert_eq!(entry_str(&out2, "xl/workbook.xml").matches(r#"name="Notes""#).count(), 1);
}

/// A workbook that declares the relationships prefix on each `<sheet>`
/// and not on its root (what openpyxl writes), or binds it to another
/// prefix: the `r:id` of a new sheet must be a declared prefix there too,
/// or the file does not open.
#[test]
fn a_new_sheet_declares_its_relationship_prefix_where_the_workbook_does_not() {
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let inline = format!(
        r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet xmlns:r="{REL}" name="Alpha" sheetId="1" r:id="rId1"/><sheet xmlns:r="{REL}" name="Beta &amp; Co" sheetId="2" r:id="rId2"/></sheets><calcPr calcId="1"/></workbook>"#
    );
    let p = patch(r#""sheetOps":[{"op":"create","name":"Charts"}]"#);
    let out = apply_patch_json(&fixture_with(&[("xl/workbook.xml", &inline)], &[]), &p).unwrap();
    crate::validate::validate_package(&out).unwrap();
    let wb = entry_str(&out, "xl/workbook.xml");
    assert!(wb.contains(&format!(r#"<sheet name="Charts" sheetId="3" xmlns:r="{REL}" r:id="rId5"/>"#)), "{wb}");

    let other = format!(
        r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:rel="{REL}"><sheets><sheet name="Alpha" sheetId="1" rel:id="rId1"/><sheet name="Beta &amp; Co" sheetId="2" rel:id="rId2"/></sheets><calcPr calcId="1"/></workbook>"#
    );
    let out = apply_patch_json(&fixture_with(&[("xl/workbook.xml", &other)], &[]), &p).unwrap();
    crate::validate::validate_package(&out).unwrap();
    let wb = entry_str(&out, "xl/workbook.xml");
    assert!(wb.contains(r#"<sheet name="Charts" sheetId="3" rel:id="rId5"/>"#), "{wb}");

    // The gate itself: a prefix nobody declared is refused, wherever it is;
    // one declared on the element or above it is fine.
    let undeclared = crate::validate::prefixes_declared("p.xml", br#"<a xmlns:x="u"><b><x:c y:id="1"/></b></a>"#).unwrap_err();
    assert!(undeclared.to_string().contains(r#"namespace prefix "y" on y:id is not declared"#), "{undeclared}");
    assert!(crate::validate::prefixes_declared("p.xml", br#"<a><b xmlns:x="u"/><x:c/></a>"#).is_err());
    crate::validate::prefixes_declared("p.xml", br#"<a xmlns:x="u"><b xmlns:y="v" y:id="1"><x:c xml:space="preserve" ref="A1:B2"/></b><d xmlns:z="w" z:k="1"/></a>"#).unwrap();
}

#[test]
fn sheet_tab_color_sets_replaces_and_clears() {
    let base = structural::fixture();
    let wb = entry_str(&base, "xl/workbook.xml");
    let rels = entry_str(&base, "xl/_rels/workbook.xml.rels");
    let parts = crate::workbook_xml::sheet_part_paths(wb.as_bytes(), rels.as_bytes()).unwrap();
    let (_, part) = parts.iter().find(|(n, _)| n == "Data").unwrap();
    let part = if part.starts_with("xl/") { part.clone() } else { format!("xl/{part}") };
    let p = patch(r##""sheetOps":[{"op":"setTabColor","name":"Data","color":"#ff8800"}]"##);
    let out = apply_patch_json(&base, &p).unwrap();
    let xml = entry_str(&out, &part);
    assert!(xml.contains(r#"<sheetPr><tabColor rgb="FFFF8800"/></sheetPr>"#), "{xml}");
    let pr = xml.find("<sheetPr").unwrap();
    let dim = xml.find("<dimension").or_else(|| xml.find("<sheetData")).unwrap();
    assert!(pr < dim, "sheetPr must precede the sheet body: {xml}");
    // Replacing keeps a single tabColor.
    let p2 = patch(r##""sheetOps":[{"op":"setTabColor","name":"Data","color":"#0000ff"}]"##);
    let out2 = apply_patch_json(&out, &p2).unwrap();
    let xml2 = entry_str(&out2, &part);
    assert_eq!(xml2.matches("<tabColor").count(), 1, "{xml2}");
    assert!(xml2.contains(r#"rgb="FF0000FF""#), "{xml2}");
    // Clearing removes it.
    let p3 = patch(r#""sheetOps":[{"op":"setTabColor","name":"Data","color":null}]"#);
    let out3 = apply_patch_json(&out2, &p3).unwrap();
    let xml3 = entry_str(&out3, &part);
    assert!(!xml3.contains("<tabColor"), "{xml3}");
    // Bad colours are rejected.
    let p4 = patch(r#""sheetOps":[{"op":"setTabColor","name":"Data","color":"orange"}]"#);
    assert!(apply_patch_json(&base, &p4).is_err());
}

#[test]
fn sheet_visibility_sets_and_clears_state_and_keeps_one_visible() {
    let base = structural::fixture();
    let p = patch(r#""sheetOps":[{"op":"setVisibility","name":"Data","hidden":true}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    let wb = entry_str(&out, "xl/workbook.xml");
    assert!(wb.contains(r#"name="Data""#) && wb.contains(r#"state="hidden""#), "{wb}");
    // Showing it again removes the attribute.
    let p2 = patch(r#""sheetOps":[{"op":"setVisibility","name":"Data","hidden":false}]"#);
    let out2 = apply_patch_json(&out, &p2).unwrap();
    let wb2 = entry_str(&out2, "xl/workbook.xml");
    assert!(!wb2.contains(r#"state="hidden""#), "{wb2}");
    // The last visible sheet cannot be hidden.
    let names: Vec<String> = crate::structure::list_sheets(wb.as_bytes()).unwrap().into_iter().map(|(n, _, _)| n).collect();
    let mut ops = String::new();
    for n in &names {
        if !ops.is_empty() { ops.push(','); }
        ops.push_str(&format!(r#"{{"op":"setVisibility","name":"{n}","hidden":true}}"#));
    }
    let p3 = patch(&format!(r#""sheetOps":[{ops}]"#));
    assert!(apply_patch_json(&base, &p3).is_err());
}

/// `cache` alone refreshes a formula's cached result (a dependent that
/// recalculated): the `<f>` stays byte-for-byte, only `<v>` and `t` change.
/// A literal cell is left alone, a missing cell is not created, and a
/// shared-formula group is not materialized by a refresh.
#[test]
fn cache_refresh_replaces_only_the_cached_value() {
    let base = fixture();
    let p = patch(
        r#""sheets":[{"name":"Alpha","cells":[
            {"r":0,"c":1,"cache":{"t":"n","n":42}},
            {"r":1,"c":1,"cache":{"t":"n","n":12}},
            {"r":1,"c":2,"cache":{"t":"s","s":"x<y"}},
            {"r":0,"c":0,"cache":{"t":"n","n":7}},
            {"r":6,"c":0,"cache":{"t":"n","n":7}}
        ]}]"#,
    );
    let out = apply_patch_json(&base, &p).unwrap();
    assert_untouched_identical(
        &base,
        &out,
        &["xl/worksheets/sheet1.xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"],
        &["xl/calcChain.xml"],
    );
    let sheet = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(sheet.contains(r#"<c r="B1"><f>A1*10</f><v>42</v></c>"#), "plain formula: {sheet}");
    assert!(sheet.contains(r#"<c r="B2"><f t="shared" ref="B2:D2" si="0">A2*2</f><v>12</v></c>"#), "shared master: {sheet}");
    assert!(sheet.contains(r#"<c r="C2" t="str"><f t="shared" si="0"/><v>x&lt;y</v></c>"#), "shared member, string: {sheet}");
    assert!(sheet.contains(r#"<c r="D2"><f t="shared" si="0"/><v>30</v></c>"#), "untouched member: {sheet}");
    assert!(sheet.contains(r#"<c r="A1"><v>1</v></c>"#), "literal kept: {sheet}");
    assert!(!sheet.contains(r#"r="A7""#) && !sheet.contains(r#"<row r="7""#), "nothing created: {sheet}");
}

/// Prefixed SpreadsheetML is valid even without a default namespace. A
/// save must preserve matching row tags and qualify inserted cells/caches.
#[test]
fn cell_edits_preserve_prefixed_spreadsheet_namespace() {
    const XML: &str = r#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c><x:c r="B1"><x:f>A1*2</x:f><x:v>2</x:v></x:c></x:row><x:row r="2"/></x:sheetData></x:worksheet>"#;
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    for (name, data) in entries(&fixture()) {
        zip.start_file(&name, SimpleFileOptions::default()).unwrap();
        zip.write_all(if name == "xl/worksheets/sheet1.xml" { XML.as_bytes() } else { &data }).unwrap();
    }
    let base = zip.finish().unwrap().into_inner();
    let p = patch(r#""sheets":[{"name":"Alpha","cells":[
        {"r":0,"c":0,"v":{"t":"n","n":5}},
        {"r":0,"c":1,"cache":{"t":"n","n":10}},
        {"r":1,"c":0,"v":{"t":"s","s":"new text"}},
        {"r":2,"c":0,"f":"A1+1","v":{"t":"n","n":6}}
    ]}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    let xml = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(xml.contains("<x:row r=\"1\">"), "{xml}");
    assert!(xml.contains("<x:v>10</x:v>"), "{xml}");
    assert!(xml.contains("<x:row r=\"2\">"), "{xml}");
    assert!(xml.contains("<x:row r=\"3\">"), "{xml}");
    let mut reader = quick_xml::NsReader::from_str(&xml);
    loop {
        let (ns, event) = reader.read_resolved_event().unwrap();
        match event {
            quick_xml::events::Event::Start(_) | quick_xml::events::Event::Empty(_) => {
                assert_eq!(ns, quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(b"http://schemas.openxmlformats.org/spreadsheetml/2006/main")));
            }
            quick_xml::events::Event::Eof => break,
            _ => {}
        }
    }
    assert_untouched_identical(&base, &out, &["xl/worksheets/sheet1.xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"], &["xl/calcChain.xml"]);
}

#[test]
fn prefixed_workbook_inserts_names_and_calc_properties_in_namespace() {
    let source = br#"<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheets/></x:workbook>"#;
    let names = [super::patch::DefinedName { name: "Discount".into(), r#ref: "Summary!$B$5".into() }];
    let xml = super::workbook_xml::patch_workbook_xml(source, true, &names).unwrap();
    let text = String::from_utf8(xml).unwrap();
    assert!(text.contains("<x:calcPr fullCalcOnLoad=\"1\"/>"), "{text}");
    assert!(text.contains("<x:definedNames><x:definedName name=\"Discount\">Summary!$B$5</x:definedName></x:definedNames>"), "{text}");
}

/// A refreshed member of a group that must materialize (another member was
/// edited) is written whole: the translated formula with its new result.
#[test]
fn cache_refresh_on_a_materialized_shared_member() {
    let base = fixture();
    let p = patch(
        r#""sheets":[{"name":"Alpha","cells":[
            {"r":1,"c":2,"v":{"t":"n","n":123}},
            {"r":1,"c":3,"cache":{"t":"b","b":true}}
        ]}]"#,
    );
    let out = apply_patch_json(&base, &p).unwrap();
    let sheet = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(sheet.contains(r#"<c r="B2"><f>A2*2</f><v>10</v></c>"#), "master materialized: {sheet}");
    assert!(sheet.contains(r#"<c r="C2"><v>123</v></c>"#), "edited member: {sheet}");
    assert!(sheet.contains(r#"<c r="D2" t="b"><f>C2*2</f><v>1</v></c>"#), "refreshed member: {sheet}");
    assert!(!sheet.contains("t=\"shared\""), "{sheet}");
}

// ---------------------------------------------------------------------------
// links to other workbooks
// ---------------------------------------------------------------------------

/// The fixture with `replace`d entries swapped and `extra` ones added.
fn fixture_with(replace: &[(&str, &str)], extra: &[(&str, &str)]) -> Vec<u8> {
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    for (name, data) in entries(&fixture()) {
        zip.start_file(&name, SimpleFileOptions::default()).unwrap();
        match replace.iter().find(|(n, _)| *n == name) {
            Some((_, with)) => zip.write_all(with.as_bytes()).unwrap(),
            None => zip.write_all(&data).unwrap(),
        }
    }
    for (name, data) in extra {
        zip.start_file(*name, SimpleFileOptions::default()).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

const LINK_OVERRIDE: &str = r#"<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>"#;
const LINK_REL: &str = r#"<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>"#;
const LINK_PATH_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///C:\Models\NOW%20FY26.xlsx" TargetMode="External"/></Relationships>"#;

/// A link part the way Excel writes one: an alternate-URL block first, the
/// source's full list of sheets, a defined name, kept values.
const EXCEL_LINK_HEAD: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14 xxl21" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xxl21="http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021"><externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"><mc:AlternateContent><mc:Choice Requires="xxl21"><xxl21:alternateUrls><xxl21:absoluteUrl r:id="rId2"/></xxl21:alternateUrls></mc:Choice></mc:AlternateContent><sheetNames><sheetName val="Cover"/><sheetName val="Model"/><sheetName val="P&amp;L"/>"#;
const EXCEL_LINK_NAMES_TAIL: &str = r#"</sheetNames><definedNames><definedName name="Rate" refersTo="=Model!$B$2"/></definedNames><sheetDataSet><sheetData sheetId="0"/>"#;
const EXCEL_LINK_DATA: &str = r#"<sheetData sheetId="1"><row r="2"><cell r="B2"><v>0.25</v></cell><cell r="C2" t="str"><v>FY25</v></cell></row><row r="23"><cell r="B23"><v>100</v></cell><cell r="C23"><v>200</v></cell></row></sheetData><sheetData sheetId="2" refreshError="1"/>"#;
const EXCEL_LINK_TAIL: &str = "</sheetDataSet></externalBook></externalLink>";

/// The fixture, already linked to one other workbook.
fn linked_fixture() -> Vec<u8> {
    let workbook = WORKBOOK.replace(
        "</sheets>",
        r#"</sheets><externalReferences><externalReference r:id="rId5"/></externalReferences>"#,
    );
    let rels = WORKBOOK_RELS.replace("</Relationships>", &format!("{LINK_REL}</Relationships>"));
    let types = CONTENT_TYPES.replace("</Types>", &format!("{LINK_OVERRIDE}</Types>"));
    let link = format!("{EXCEL_LINK_HEAD}{EXCEL_LINK_NAMES_TAIL}{EXCEL_LINK_DATA}{EXCEL_LINK_TAIL}");
    fixture_with(
        &[
            ("xl/workbook.xml", &workbook),
            ("xl/_rels/workbook.xml.rels", &rels),
            ("[Content_Types].xml", &types),
        ],
        &[
            ("xl/externalLinks/externalLink1.xml", &link),
            ("xl/externalLinks/_rels/externalLink1.xml.rels", LINK_PATH_RELS),
        ],
    )
}

#[test]
fn link_to_another_workbook_is_created() {
    use super::external_links::{read, ExternalLinkInfo};
    let base = fixture();
    assert_eq!(read(&base).unwrap(), vec![]);
    let p = patch(
        r##""sheets":[{"name":"Alpha","cells":[{"r":5,"c":0,"f":"[1]Model!B23+'[1]Income Statement'!A1","v":{"t":"n","n":1250.5}}]}],
        "externalLinks":[{"index":1,"target":"NOW%202026.xlsx","sheets":[
            {"name":"Model","areas":[[22,1,22,1],[1,0,2,1]],"cells":[
                {"r":22,"c":1,"v":{"t":"n","n":1250}},
                {"r":1,"c":0,"v":{"t":"s","s":"R&D <net>"}},
                {"r":2,"c":1,"v":{"t":"e","e":"#DIV/0!"}},
                {"r":1,"c":1,"v":{"t":"b","b":false}}]},
            {"name":"Income Statement","areas":[[0,0,0,0]],"cells":[{"r":0,"c":0,"v":{"t":"n","n":0.5}}]},
            {"name":"Empty","areas":[[0,0,0,0]]}]}]"##,
    );
    let out = apply_patch_json(&base, &p).unwrap();

    assert_eq!(
        entry_str(&out, "xl/externalLinks/externalLink1.xml"),
        concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
            r#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
            r#"<externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1">"#,
            r#"<sheetNames><sheetName val="Model"/><sheetName val="Income Statement"/><sheetName val="Empty"/></sheetNames>"#,
            r#"<sheetDataSet><sheetData sheetId="0">"#,
            r#"<row r="2"><cell r="A2" t="str"><v>R&amp;D &lt;net&gt;</v></cell><cell r="B2" t="b"><v>0</v></cell></row>"#,
            r##"<row r="3"><cell r="B3" t="e"><v>#DIV/0!</v></cell></row>"##,
            r#"<row r="23"><cell r="B23"><v>1250</v></cell></row></sheetData>"#,
            r#"<sheetData sheetId="1"><row r="1"><cell r="A1"><v>0.5</v></cell></row></sheetData>"#,
            r#"<sheetData sheetId="2"/></sheetDataSet></externalBook></externalLink>"#,
        )
    );
    let link_rels = entry_str(&out, "xl/externalLinks/_rels/externalLink1.xml.rels");
    assert!(
        link_rels.contains(r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="NOW%202026.xlsx" TargetMode="External"/>"#),
        "{link_rels}"
    );

    // The workbook lists it where the schema wants the list, and reaches the
    // part through a relationship of its own.
    // (rId5 went to the metadata part the formula needs.)
    let wb = entry_str(&out, "xl/workbook.xml");
    assert!(
        wb.contains(r#"</sheets><externalReferences><externalReference r:id="rId6"/></externalReferences><definedNames>"#),
        "{wb}"
    );
    let rels = entry_str(&out, "xl/_rels/workbook.xml.rels");
    assert!(rels.contains(&LINK_REL.replace("rId5", "rId6")), "{rels}");
    let types = entry_str(&out, "[Content_Types].xml");
    assert!(types.contains(LINK_OVERRIDE), "{types}");
    let sheet = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(sheet.contains("[1]Model!B23+&apos;[1]Income Statement&apos;!A1</f><v>1250.5</v>"), "{sheet}");

    assert_untouched_identical(
        &base,
        &out,
        &["xl/worksheets/sheet1.xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"],
        &["xl/calcChain.xml"],
    );
    super::validate::validate_package(&out).unwrap();
    assert_eq!(
        read(&out).unwrap(),
        vec![ExternalLinkInfo {
            index: 1,
            target: Some("NOW%202026.xlsx".into()),
            sheets: vec!["Model".into(), "Income Statement".into(), "Empty".into()],
        }]
    );
}

#[test]
fn links_are_numbered_in_the_order_they_are_added() {
    use super::external_links::read;
    let base = fixture();
    // Given out of order: 1 is still created first.
    let p = patch(
        r#""externalLinks":[
            {"index":2,"target":"/Users/me/Models/ANT.xlsx","sheets":[{"name":"Summary"}]},
            {"index":1,"target":"NOW.xlsx","sheets":[{"name":"Model"}]}]"#,
    );
    let out = apply_patch_json(&base, &p).unwrap();
    let wb = entry_str(&out, "xl/workbook.xml");
    assert!(
        wb.contains(r#"<externalReferences><externalReference r:id="rId5"/><externalReference r:id="rId6"/></externalReferences>"#),
        "{wb}"
    );
    // Nothing else of the workbook part changed: no formula was written.
    assert!(!wb.contains("fullCalcOnLoad"), "{wb}");
    let links = read(&out).unwrap();
    assert_eq!(links.len(), 2);
    assert_eq!(links[0].target.as_deref(), Some("NOW.xlsx"));
    assert_eq!(links[1].target.as_deref(), Some("/Users/me/Models/ANT.xlsx"));
    assert_eq!(links[1].sheets, vec!["Summary".to_string()]);
    super::validate::validate_package(&out).unwrap();

    // A third on top of the saved two.
    let p = patch(r#""externalLinks":[{"index":3,"target":"C.xlsx","sheets":[{"name":"S"}]}]"#);
    let again = apply_patch_json(&out, &p).unwrap();
    assert_eq!(read(&again).unwrap().len(), 3);
    assert!(entries(&again).iter().any(|(n, _)| n == "xl/externalLinks/externalLink3.xml"));
    assert_untouched_identical(
        &out,
        &again,
        &["xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"],
        &[],
    );

    for bad in [
        // A gap: the workbook has none, so the first is 1.
        r#""externalLinks":[{"index":2,"target":"NOW.xlsx"}]"#,
        r#""externalLinks":[{"index":0,"target":"NOW.xlsx"}]"#,
        // No file named.
        r#""externalLinks":[{"index":1,"sheets":[{"name":"Model"}]}]"#,
        // The same link twice, the same sheet twice, a cell off the sheet.
        r#""externalLinks":[{"index":1,"target":"A.xlsx"},{"index":1,"target":"B.xlsx"}]"#,
        r#""externalLinks":[{"index":1,"target":"A.xlsx","sheets":[{"name":"Model"},{"name":"MODEL"}]}]"#,
        r#""externalLinks":[{"index":1,"target":"A.xlsx","sheets":[{"name":"Model","cells":[{"r":1048576,"c":0,"v":{"t":"n","n":1}}]}]}]"#,
        r#""externalLinks":[{"index":1,"target":"A.xlsx","sheets":[{"name":"Model","areas":[[2,0,1,0]]}]}]"#,
    ] {
        let e = apply_patch_json(&base, &patch(bad)).unwrap_err();
        assert!(matches!(e, super::PatchError::BadValue(_)), "{bad}: {e}");
    }
}

#[test]
fn existing_link_is_brought_up_to_date_in_place() {
    use super::external_links::read;
    let base = linked_fixture();
    super::validate::validate_package(&base).unwrap();
    let before = read(&base).unwrap();
    assert_eq!(before.len(), 1);
    assert_eq!(before[0].target.as_deref(), Some(r"file:///C:\Models\NOW%20FY26.xlsx"));
    assert_eq!(before[0].sheets, vec!["Cover".to_string(), "Model".into(), "P&L".into()]);

    // Only the link: B23 has a new value, C23 is now empty, P&L!A1 is read
    // for the first time, and so is a sheet the link did not list.
    let p = patch(
        r##""externalLinks":[{"index":1,"target":"ignored.xlsx","sheets":[
            {"name":"model","areas":[[22,1,22,2]],"cells":[{"r":22,"c":1,"v":{"t":"n","n":150}}]},
            {"name":"P&L","areas":[[0,0,0,0]],"cells":[{"r":0,"c":0,"v":{"t":"s","s":"a<b"}}]},
            {"name":"New Sheet","areas":[[4,4,5,5]],"cells":[
                {"r":5,"c":5,"v":{"t":"e","e":"#N/A"}},
                {"r":4,"c":4,"v":{"t":"b","b":true}}]}]}]"##,
    );
    let out = apply_patch_json(&base, &p).unwrap();
    let expected = [
        EXCEL_LINK_HEAD,
        r#"<sheetName val="New Sheet"/>"#,
        EXCEL_LINK_NAMES_TAIL,
        // Values outside the patch's areas keep their bytes.
        r#"<sheetData sheetId="1"><row r="2"><cell r="B2"><v>0.25</v></cell><cell r="C2" t="str"><v>FY25</v></cell></row><row r="23"><cell r="B23"><v>150</v></cell></row></sheetData>"#,
        r#"<sheetData sheetId="2" refreshError="1"><row r="1"><cell r="A1" t="str"><v>a&lt;b</v></cell></row></sheetData>"#,
        r##"<sheetData sheetId="3"><row r="5"><cell r="E5" t="b"><v>1</v></cell></row><row r="6"><cell r="F6" t="e"><v>#N/A</v></cell></row></sheetData>"##,
        EXCEL_LINK_TAIL,
    ]
    .concat();
    assert_eq!(entry_str(&out, "xl/externalLinks/externalLink1.xml"), expected);
    // The path it was saved with stays, and so does everything else.
    assert_untouched_identical(&base, &out, &["xl/externalLinks/externalLink1.xml"], &[]);
    assert_eq!(entries(&base).len(), entries(&out).len());
    super::validate::validate_package(&out).unwrap();
    assert_eq!(
        read(&out).unwrap()[0].sheets,
        vec!["Cover".to_string(), "Model".into(), "P&L".into(), "New Sheet".into()]
    );

    // Applying the same patch again changes nothing.
    let again = apply_patch_json(&out, &p).unwrap();
    assert_eq!(entry_str(&again, "xl/externalLinks/externalLink1.xml"), expected);

    // A second link goes after the saved one.
    let p = patch(r#""externalLinks":[{"index":2,"target":"ANT.xlsx","sheets":[{"name":"Summary"}]}]"#);
    let two = apply_patch_json(&out, &p).unwrap();
    let wb = entry_str(&two, "xl/workbook.xml");
    assert!(
        wb.contains(r#"<externalReferences><externalReference r:id="rId5"/><externalReference r:id="rId6"/></externalReferences>"#),
        "{wb}"
    );
    assert!(entries(&two).iter().any(|(n, _)| n == "xl/externalLinks/externalLink2.xml"));
    assert_eq!(read(&two).unwrap()[1].target.as_deref(), Some("ANT.xlsx"));
    super::validate::validate_package(&two).unwrap();
}

#[test]
fn link_parts_in_other_shapes() {
    use super::external_links::read;
    // A prefixed part whose book has nothing in it yet.
    let bare = r#"<x:externalLink xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></x:externalLink>"#;
    let mut base = linked_fixture();
    base = {
        let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
        for (name, data) in entries(&base) {
            zip.start_file(&name, SimpleFileOptions::default()).unwrap();
            zip.write_all(if name == "xl/externalLinks/externalLink1.xml" { bare.as_bytes() } else { &data }).unwrap();
        }
        zip.finish().unwrap().into_inner()
    };
    assert_eq!(read(&base).unwrap()[0].sheets, Vec::<String>::new());
    let p = patch(
        r#""externalLinks":[{"index":1,"sheets":[{"name":"Model","areas":[[0,0,0,0]],"cells":[{"r":0,"c":0,"v":{"t":"n","n":7}}]}]}]"#,
    );
    let out = apply_patch_json(&base, &p).unwrap();
    assert_eq!(
        entry_str(&out, "xl/externalLinks/externalLink1.xml"),
        r#"<x:externalLink xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"><x:sheetNames><x:sheetName val="Model"/></x:sheetNames><x:sheetDataSet><x:sheetData sheetId="0"><x:row r="1"><x:cell r="A1"><x:v>7</x:v></x:cell></x:row></x:sheetData></x:sheetDataSet></x:externalBook></x:externalLink>"#
    );

    // Names but no values yet, and the names self-closed.
    let update = |part: &str| {
        let sheets: Vec<super::patch::ExternalSheetPatch> = serde_json::from_str(
            r#"[{"name":"Model","areas":[[0,0,0,0]],"cells":[{"r":0,"c":0,"v":{"t":"n","n":7}}]}]"#,
        )
        .unwrap();
        String::from_utf8(super::external_links::update_part(part.as_bytes(), &sheets).unwrap()).unwrap()
    };
    assert_eq!(
        update(r#"<externalLink><externalBook r:id="rId1"><sheetNames><sheetName val="Model"/></sheetNames></externalBook></externalLink>"#),
        r#"<externalLink><externalBook r:id="rId1"><sheetNames><sheetName val="Model"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>7</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>"#
    );
    assert_eq!(
        update(r#"<externalLink><externalBook r:id="rId1"><sheetNames/><definedNames/><sheetDataSet/></externalBook></externalLink>"#),
        r#"<externalLink><externalBook r:id="rId1"><sheetNames><sheetName val="Model"/></sheetNames><definedNames/><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>7</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>"#
    );
    assert_eq!(
        update(r#"<externalLink><externalBook r:id="rId1"><definedNames/></externalBook></externalLink>"#),
        r#"<externalLink><externalBook r:id="rId1"><sheetNames><sheetName val="Model"/></sheetNames><definedNames/><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>7</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>"#
    );
    // With a byte-order mark in front, offsets still land.
    let with_bom = update("\u{feff}<externalLink><externalBook r:id=\"rId1\"><sheetNames><sheetName val=\"Model\"/></sheetNames><sheetDataSet><sheetData sheetId=\"0\"><row r=\"1\"><cell r=\"A1\"><v>1</v></cell><cell r=\"B1\"><v>2</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>");
    assert_eq!(
        with_bom,
        "\u{feff}<externalLink><externalBook r:id=\"rId1\"><sheetNames><sheetName val=\"Model\"/></sheetNames><sheetDataSet><sheetData sheetId=\"0\"><row r=\"1\"><cell r=\"A1\"><v>7</v></cell><cell r=\"B1\"><v>2</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>"
    );
}

#[test]
fn a_link_that_is_not_a_workbook_keeps_its_number() {
    use super::external_links::read;
    // Index 1 is a DDE link; the workbook link is index 2.
    let dde = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><ddeLink ddeService="Excel" ddeTopic="Quotes"><ddeItems><ddeItem name="Last" advise="1"/></ddeItems></ddeLink></externalLink>"#;
    let workbook = WORKBOOK.replace(
        "</sheets>",
        r#"</sheets><externalReferences><externalReference r:id="rId6"/><externalReference r:id="rId5"/></externalReferences>"#,
    );
    let rels = WORKBOOK_RELS.replace(
        "</Relationships>",
        &format!(r#"{LINK_REL}<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="/xl/externalLinks/externalLink2.xml"/></Relationships>"#),
    );
    let types = CONTENT_TYPES.replace(
        "</Types>",
        &format!("{LINK_OVERRIDE}{}</Types>", LINK_OVERRIDE.replace("externalLink1", "externalLink2")),
    );
    let link = format!("{EXCEL_LINK_HEAD}{EXCEL_LINK_NAMES_TAIL}{EXCEL_LINK_DATA}{EXCEL_LINK_TAIL}");
    let base = fixture_with(
        &[
            ("xl/workbook.xml", &workbook),
            ("xl/_rels/workbook.xml.rels", &rels),
            ("[Content_Types].xml", &types),
        ],
        &[
            ("xl/externalLinks/externalLink1.xml", &link),
            ("xl/externalLinks/_rels/externalLink1.xml.rels", LINK_PATH_RELS),
            ("xl/externalLinks/externalLink2.xml", dde),
        ],
    );
    let links = read(&base).unwrap();
    assert_eq!(links.len(), 2);
    assert_eq!((links[0].index, links[0].target.clone(), links[0].sheets.len()), (1, None, 0));
    assert_eq!(links[1].index, 2);
    assert_eq!(links[1].sheets.len(), 3);

    // Values cannot go into the DDE link.
    let p = patch(r#""externalLinks":[{"index":1,"sheets":[{"name":"Model"}]}]"#);
    assert!(matches!(apply_patch_json(&base, &p).unwrap_err(), super::PatchError::BadValue(_)));
    // The workbook link is updated by its own number, and a new link is 3,
    // in a part whose name is free.
    let p = patch(
        r#""externalLinks":[
            {"index":2,"sheets":[{"name":"Cover","areas":[[0,0,0,0]],"cells":[{"r":0,"c":0,"v":{"t":"s","s":"Draft"}}]}]},
            {"index":3,"target":"ANT.xlsx","sheets":[{"name":"Summary"}]}]"#,
    );
    let out = apply_patch_json(&base, &p).unwrap();
    let updated = entry_str(&out, "xl/externalLinks/externalLink1.xml");
    assert!(
        updated.contains(r#"<sheetData sheetId="0"><row r="1"><cell r="A1" t="str"><v>Draft</v></cell></row></sheetData>"#),
        "{updated}"
    );
    assert_eq!(entry_str(&out, "xl/externalLinks/externalLink2.xml"), dde);
    let links = read(&out).unwrap();
    assert_eq!(links.len(), 3);
    assert_eq!(links[2].target.as_deref(), Some("ANT.xlsx"));
    assert!(entries(&out).iter().any(|(n, _)| n == "xl/externalLinks/externalLink3.xml"));
    super::validate::validate_package(&out).unwrap();
}

#[test]
fn external_reference_lands_where_the_schema_wants_it() {
    let add = |xml: &str| {
        String::from_utf8(super::external_links::add_external_reference(xml.as_bytes(), "rId9").unwrap()).unwrap()
    };
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    // Prefixed workbook, relationships bound to another prefix.
    assert_eq!(
        add(&format!(r#"<x:workbook xmlns:x="main" xmlns:rel="{REL}"><x:sheets/><x:calcPr/></x:workbook>"#)),
        format!(r#"<x:workbook xmlns:x="main" xmlns:rel="{REL}"><x:sheets/><x:externalReferences><x:externalReference rel:id="rId9"/></x:externalReferences><x:calcPr/></x:workbook>"#)
    );
    // No prefix for relationships on the root: the element brings its own.
    assert_eq!(
        add(r#"<workbook xmlns="main"><sheets/></workbook>"#),
        format!(r#"<workbook xmlns="main"><sheets/><externalReferences><externalReference xmlns:r="{REL}" r:id="rId9"/></externalReferences></workbook>"#)
    );
    // A list that is there: at its end, also when it is empty.
    assert_eq!(
        add(&format!(r#"<workbook xmlns:r="{REL}"><sheets/><externalReferences><externalReference r:id="rId3"/></externalReferences><definedNames><definedName name="a">1</definedName></definedNames></workbook>"#)),
        format!(r#"<workbook xmlns:r="{REL}"><sheets/><externalReferences><externalReference r:id="rId3"/><externalReference r:id="rId9"/></externalReferences><definedNames><definedName name="a">1</definedName></definedNames></workbook>"#)
    );
    assert_eq!(
        add(&format!(r#"<workbook xmlns:r="{REL}"><sheets/><externalReferences/><extLst/></workbook>"#)),
        format!(r#"<workbook xmlns:r="{REL}"><sheets/><externalReferences><externalReference r:id="rId9"/></externalReferences><extLst/></workbook>"#)
    );
    assert_eq!(
        add(&format!(r#"<workbook xmlns:r="{REL}"><sheets/><externalReferences></externalReferences><definedNames></definedNames></workbook>"#)),
        format!(r#"<workbook xmlns:r="{REL}"><sheets/><externalReferences><externalReference r:id="rId9"/></externalReferences><definedNames></definedNames></workbook>"#)
    );
}

// ---------------------------------------------------------------------------
// sheet names that XML escapes
// ---------------------------------------------------------------------------

/// A sheet named with `&` is written `&amp;` in every part that names it.
/// Rows inserted on it must still move the formulas of other sheets that
/// read it, and a defined name that points at it.
#[test]
fn rows_inserted_on_a_sheet_with_an_ampersand_move_what_reads_it() {
    let sheet1 = SHEET1.replace(
        "</sheetData>",
        r#"<row r="9"><c r="A9"><f>'Beta &amp; Co'!A2+1</f><v>1</v></c></row></sheetData>"#,
    );
    let workbook = WORKBOOK.replace(
        "</definedNames>",
        r#"<definedName name="RATE">'Beta &amp; Co'!$B$2</definedName></definedNames>"#,
    );
    let base = fixture_with(&[("xl/worksheets/sheet1.xml", &sheet1), ("xl/workbook.xml", &workbook)], &[]);
    let p = patch(r#""rowColOps":[{"op":"insertRows","sheet":"Beta & Co","before":0,"count":2}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    let alpha = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(alpha.contains("<f>'Beta &amp; Co'!A4+1</f>"), "the formula on the other sheet follows: {alpha}");
    let wb = entry_str(&out, "xl/workbook.xml");
    assert!(wb.contains("'Beta &amp; Co'!$B$4"), "the defined name follows: {wb}");
}

/// A sheet that other sheets read cannot be deleted, whatever its name.
#[test]
fn a_referenced_sheet_with_an_ampersand_is_not_deleted() {
    let sheet1 = SHEET1.replace(
        "</sheetData>",
        r#"<row r="9"><c r="A9"><f>'Beta &amp; Co'!A2+1</f><v>1</v></c></row></sheetData>"#,
    );
    let base = fixture_with(&[("xl/worksheets/sheet1.xml", &sheet1)], &[]);
    let p = patch(r#""sheetOps":[{"op":"delete","name":"Beta & Co"}]"#);
    let err = apply_patch_json(&base, &p).unwrap_err();
    assert!(err.to_string().to_lowercase().contains("referenc"), "{err}");
    // Not read by anything: it goes.
    let out = apply_patch_json(&fixture(), &p).unwrap();
    assert!(!entry_str(&out, "xl/workbook.xml").contains("Beta &amp; Co"));
}

/// The same for a rename: the formulas that name the sheet are respelled.
#[test]
fn renaming_a_sheet_with_an_ampersand_respells_what_reads_it() {
    let sheet1 = SHEET1.replace(
        "</sheetData>",
        r#"<row r="9"><c r="A9"><f>'Beta &amp; Co'!A2+1</f><v>1</v></c></row></sheetData>"#,
    );
    let base = fixture_with(&[("xl/worksheets/sheet1.xml", &sheet1)], &[]);
    let p = patch(r#""sheetOps":[{"op":"rename","oldName":"Beta & Co","newName":"P&L"}]"#);
    let out = apply_patch_json(&base, &p).unwrap();
    let alpha = entry_str(&out, "xl/worksheets/sheet1.xml");
    assert!(alpha.contains("<f>'P&amp;L'!A2+1</f>"), "{alpha}");
}


// ---------------------------------------------------------------------------
// tables and pivot tables elsewhere in the workbook
// ---------------------------------------------------------------------------

const CALC_RELS_TABLE: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>"#;
const CALC_RELS_PIVOT: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>"#;
const TABLE_ON_CALC: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Totals" displayName="Totals" ref="A1:A3"><autoFilter ref="A1:A3"/><tableColumns count="1"><tableColumn id="1" name="Total"/></tableColumns></table>"#;
const PIVOT_ON_CALC: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1"><location ref="C1:D4" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/></pivotTableDefinition>"#;

fn pivot_cache(source_sheet: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshedBy="x"><cacheSource type="worksheet"><worksheetSource ref="A1:D6" sheet="{source_sheet}"/></cacheSource></pivotCacheDefinition>"#
    )
    .into_bytes()
}

const INSERT_ON_DATA: &str = r#""rowColOps":[{"op":"insertRows","sheet":"Data","before":0,"count":1}]"#;
const INSERT_ON_CALC: &str = r#""rowColOps":[{"op":"insertRows","sheet":"Calc","before":0,"count":1}]"#;

/// A workbook that has a table somewhere is edited like any other: rows
/// move on the sheets the table is not on, and the table is not touched.
/// Rows on the table's own sheet are still refused.
#[test]
fn a_table_on_another_sheet_does_not_stop_row_edits() {
    let base = structural::fixture_with(&[
        ("xl/worksheets/_rels/sheet2.xml.rels", CALC_RELS_TABLE),
        ("xl/tables/table1.xml", TABLE_ON_CALC),
    ]);
    let out = apply_patch_json(&base, &patch(INSERT_ON_DATA)).unwrap();
    assert_eq!(entry(&out, "xl/tables/table1.xml"), TABLE_ON_CALC, "the table is as it was");
    // The formula on the table's sheet that reads the edited sheet still follows.
    assert!(entry_str(&out, "xl/worksheets/sheet2.xml").contains("<f>Data!B4+1</f>"));

    let err = apply_patch_json(&base, &patch(INSERT_ON_CALC)).unwrap_err().to_string();
    assert!(err.contains("unsupported") && err.contains("tables") && err.contains("Calc"), "{err}");
}

/// A table's own formulas can read another sheet: rows moving there would
/// leave them pointing at the wrong cells, so that edit is refused.
#[test]
fn a_table_that_reads_the_edited_sheet_stops_the_edit() {
    let table = String::from_utf8(TABLE_ON_CALC.to_vec())
        .unwrap()
        .replace(
            r#"<tableColumn id="1" name="Total"/>"#,
            r#"<tableColumn id="1" name="Total"><calculatedColumnFormula>Data!B2*2</calculatedColumnFormula></tableColumn>"#,
        );
    let base = structural::fixture_with(&[
        ("xl/worksheets/_rels/sheet2.xml.rels", CALC_RELS_TABLE),
        ("xl/tables/table1.xml", table.as_bytes()),
    ]);
    let err = apply_patch_json(&base, &patch(INSERT_ON_DATA)).unwrap_err().to_string();
    assert!(err.contains("unsupported") && err.contains("refers to sheet"), "{err}");
}

/// A pivot table matters to an edit in two places: the sheet it sits on,
/// and the sheet its cache reads.
#[test]
fn a_pivot_table_stops_edits_only_where_it_sits_and_where_it_reads() {
    let with_source = |sheet: &str| {
        structural::fixture_with(&[
            ("xl/worksheets/_rels/sheet2.xml.rels", CALC_RELS_PIVOT),
            ("xl/pivotTables/pivotTable1.xml", PIVOT_ON_CALC),
            ("xl/pivotCache/pivotCacheDefinition1.xml", &pivot_cache(sheet)),
            // Cached rows that happen to hold the sheet's name are data.
            ("xl/pivotCache/pivotCacheRecords1.xml", b"<pivotCacheRecords><r><s v=\"Data\"/></r></pivotCacheRecords>"),
        ])
    };
    // Sits on Calc and reads Calc: Data is free.
    let base = with_source("Calc");
    let out = apply_patch_json(&base, &patch(INSERT_ON_DATA)).unwrap();
    assert_eq!(entry(&out, "xl/pivotTables/pivotTable1.xml"), PIVOT_ON_CALC);
    let err = apply_patch_json(&base, &patch(INSERT_ON_CALC)).unwrap_err().to_string();
    assert!(err.contains("pivotTables") && err.contains("Calc"), "{err}");
    // Reads Data: rows there would move its source.
    let err = apply_patch_json(&with_source("Data"), &patch(INSERT_ON_DATA)).unwrap_err().to_string();
    assert!(err.contains("pivotCache") && err.contains("refers to sheet"), "{err}");
    // A rename is about the name alone: the sheet the pivot sits on can be
    // renamed, the sheet its cache names cannot.
    let rename = |from: &str| patch(&format!(r#""sheetOps":[{{"op":"rename","oldName":"{from}","newName":"Renamed"}}]"#));
    apply_patch_json(&with_source("Data"), &rename("Calc")).unwrap();
    let err = apply_patch_json(&with_source("Data"), &rename("Data")).unwrap_err().to_string();
    assert!(err.contains("pivotCache"), "{err}");
}

/// A tab colour or hiding a sheet moves nothing and renames nothing.
#[test]
fn tab_colour_and_visibility_work_in_a_workbook_with_a_table() {
    let base = structural::fixture_with(&[
        ("xl/worksheets/_rels/sheet2.xml.rels", CALC_RELS_TABLE),
        ("xl/tables/table1.xml", TABLE_ON_CALC),
    ]);
    let out = apply_patch_json(&base, &patch(r##""sheetOps":[{"op":"setTabColor","name":"Calc","color":"#FF0000"}]"##)).unwrap();
    assert!(entry_str(&out, "xl/worksheets/sheet2.xml").contains("tabColor"));
    let out = apply_patch_json(&base, &patch(r#""sheetOps":[{"op":"setVisibility","name":"Calc","hidden":true}]"#)).unwrap();
    assert!(entry_str(&out, "xl/workbook.xml").contains(r#"state="hidden""#));
}
