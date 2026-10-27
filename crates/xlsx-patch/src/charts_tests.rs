use crate::{apply_patch_json, charts::*, package::PartStore};

const PART: &str = "xl/drawings/charts/chart1.xml";
const CHART: &str = r##"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Daily spend</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:lineChart><c:grouping val="standard"/><c:ser><c:tx><c:v>Spend</c:v></c:tx><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="187E70"/></a:solidFill></a:ln></c:spPr><c:cat><c:numRef><c:f>Alpha!$A$1:$A$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:cat><c:val><c:numRef><c:f>Alpha!$B$1:$B$3</c:f><c:numCache><c:formatCode>$0.00</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>10</c:v></c:pt><c:extLst><c:ext uri="retained"/></c:extLst></c:numCache></c:numRef></c:val></c:ser></c:lineChart><c:valAx><c:numFmt formatCode="$0"/><c:scaling><c:min val="0"/></c:scaling></c:valAx></c:plotArea></c:chart></c:chartSpace>"##;

fn fixture(chart: &str) -> Vec<u8> {
    let base = crate::tests::fixture();
    let mut store = PartStore::open(&base).unwrap();
    let types = String::from_utf8(store.read("[Content_Types].xml").unwrap()).unwrap().replace("</Types>", &format!(r#"<Override PartName="/{PART}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>"#));
    store.write("[Content_Types].xml", types.into_bytes());
    store.write("xl/worksheets/_rels/sheet1.xml.rels", br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="draw" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="/xl/drawings/drawing1.xml"/></Relationships>"#.to_vec());
    store.write("xl/drawings/_rels/drawing1.xml.rels", br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="chart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>"#.to_vec());
    store.write("xl/drawings/drawing1.xml", br#"<x:wsDr xmlns:x="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:twoCellAnchor><x:from><x:row>4</x:row><x:col>5</x:col><x:colOff>9525</x:colOff></x:from><x:to><x:row>20</x:row><x:col>12</x:col></x:to><x:graphicFrame><c:chart r:id="chart"/></x:graphicFrame></x:twoCellAnchor></x:wsDr>"#.to_vec());
    store.write(PART, chart.as_bytes().to_vec());
    store.finish().unwrap()
}
fn chart_text(bytes: &[u8]) -> String {
    String::from_utf8(PartStore::open(bytes).unwrap().read(PART).unwrap()).unwrap()
}

#[test]
fn discovers_relationship_paths_and_sparse_series() {
    let c = read(&fixture(CHART)).unwrap().remove(0);
    assert_eq!(c.part, PART);
    assert_eq!(c.sheet, "Alpha");
    assert_eq!(c.title, "Daily spend");
    assert_eq!(c.unsupported, None);
    assert_eq!(c.anchor.from.row, 4);
    assert_eq!(c.anchor.from.dx, 1.);
    assert_eq!(c.anchor.to.unwrap().col, 12);
    assert_eq!(c.series[0].values, vec![Some(10.), Some(10.), None]);
    assert_eq!(c.series[0].color, "#187E70");
    assert_eq!(c.y_min, Some(0.));
}

#[test]
fn native_cache_refresh_retains_everything_outside_changed_cache() {
    let bytes = fixture(CHART);
    let out = apply_patch_json(&bytes, r#"{"version":1,"sheets":[{"name":"Alpha","cells":[{"r":0,"c":1,"cache":{"t":"n","n":42}}]}]}"#).unwrap();
    let c = read(&out).unwrap().remove(0);
    assert_eq!(c.series[0].values, vec![Some(42.), Some(10.), None]);
    let original = chart_text(&bytes);
    let changed = chart_text(&out);
    let strip = |s: &str| {
        let start = s.find("<c:val>").unwrap();
        let end = s.find("</c:val>").unwrap();
        format!("{}{}", &s[..start], &s[end..])
    };
    assert_eq!(strip(&original), strip(&changed));
    assert!(changed.contains("<c:formatCode>$0.00</c:formatCode>"));
    assert!(changed.contains("<c:extLst><c:ext uri=\"retained\"/></c:extLst>"));
    let again = apply_patch_json(&out, r#"{"version":1,"sheets":[{"name":"Alpha","cells":[{"r":0,"c":1,"cache":{"t":"n","n":42}}]}]}"#).unwrap();
    assert_eq!(chart_text(&again), changed);
}

#[test]
fn unrelated_edit_does_not_rewrite_chart() {
    let bytes = fixture(CHART);
    let out = apply_patch_json(
        &bytes,
        r#"{"version":1,"sheets":[{"name":"Alpha","cells":[{"r":8,"c":8,"v":{"t":"n","n":42}}]}]}"#,
    )
    .unwrap();
    assert_eq!(chart_text(&bytes), chart_text(&out));
}

#[test]
fn rename_and_row_shift_find_chart_in_nonconventional_directory() {
    let out=apply_patch_json(&fixture(CHART), r#"{"version":1,"sheetOps":[{"op":"rename","oldName":"Alpha","newName":"New Alpha"}],"rowColOps":[{"op":"insertRows","sheet":"New Alpha","before":1,"count":1}]}"#).unwrap();
    let c = read(&out).unwrap().remove(0);
    assert_eq!(c.sheet, "New Alpha");
    assert_eq!(
        c.series[0].values_ref.as_deref(),
        Some("'New Alpha'!$B$1:$B$4")
    );
    assert_eq!(c.series[0].values, vec![Some(10.), None, Some(10.), None]);
    assert_eq!(c.anchor.from.row, 5);
}

#[test]
fn unsupported_and_malformed_charts_are_explicit() {
    for xml in [
        CHART.replace("lineChart", "barChart"),
        CHART.replace("Alpha!$B$1:$B$3", "MyNamedRange"),
        "<broken>".into(),
    ] {
        let c = read(&fixture(&xml)).unwrap().remove(0);
        assert!(c.unsupported.is_some());
    }
}

#[test]
fn reference_bounds_and_quoted_names() {
    assert_eq!(
        range("'O''Brien data'!$B$2:$B$5").unwrap().sheet,
        "O'Brien data"
    );
    for f in [
        "[other.xlsx]Sheet!A1",
        "Sheet!A1:B3",
        "Sheet!A1:A1048576",
        "NamedRange",
        "Sheet!#REF!",
        "'!A1",
    ] {
        assert!(range(f).is_none(), "{f}");
    }
}

#[test]
fn bom_parts_preserve_byte_offsets_when_refreshing() {
    let bytes = fixture(&format!("\u{feff}{CHART}"));
    let mut store = PartStore::open(&bytes).unwrap();
    for part in [
        "_rels/.rels",
        "xl/drawings/_rels/drawing1.xml.rels",
        "xl/worksheets/sheet1.xml",
    ] {
        let mut xml = vec![0xef, 0xbb, 0xbf];
        xml.extend(store.read(part).unwrap());
        store.write(part, xml);
    }
    let bytes = store.finish().unwrap();
    assert_eq!(read(&bytes).unwrap().len(), 1);
    let out=apply_patch_json(&bytes,r#"{"version":1,"sheets":[{"name":"Alpha","cells":[{"r":0,"c":1,"cache":{"t":"n","n":42}}]}]}"#).unwrap();
    assert_eq!(read(&out).unwrap()[0].series[0].values[0], Some(42.));
    assert!(chart_text(&out).starts_with('\u{feff}'));
}

#[test]
fn cache_local_namespace_is_retained() {
    let xml=CHART.replace("<c:numCache>","<q:numCache xmlns:q=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" custom=\"retained\">").replace("</c:numCache>","</q:numCache>");
    let out=apply_patch_json(&fixture(&xml),r#"{"version":1,"sheets":[{"name":"Alpha","cells":[{"r":0,"c":1,"cache":{"t":"n","n":42}}]}]}"#).unwrap();
    assert_eq!(read(&out).unwrap()[0].series[0].values[0], Some(42.));
    assert!(chart_text(&out).contains("<q:pt idx=\"0\"><q:v>42.0</q:v></q:pt>"));
    assert!(chart_text(&out).contains("custom=\"retained\""));
}

#[test]
fn source_strings_errors_case_and_empty_cache() {
    let start = CHART.find("<c:cat>").unwrap();
    let end = CHART.find("</c:cat>").unwrap() + 8;
    let chart = format!(
        "{}<c:cat><c:strRef><c:f>alpha!$D$1</c:f><c:strCache /></c:strRef></c:cat>{}",
        &CHART[..start],
        &CHART[end..]
    );
    let bytes = fixture(&chart);
    let mut store = PartStore::open(&bytes).unwrap();
    store.write(
        "xl/sharedStrings.xml",
        br#"<sst><si><t>R&amp;D &lt;x&gt;</t></si></sst>"#.to_vec(),
    );
    let out=apply_patch_json(&store.finish().unwrap(),r##"{"version":1,"sheets":[{"name":"Alpha","cells":[{"r":1,"c":1,"cache":{"t":"e","e":"#DIV/0!"}}]}]}"##).unwrap();
    let c = read(&out).unwrap().remove(0);
    assert_eq!(c.series[0].categories, vec![serde_json::json!("R&D <x>")]);
    assert_eq!(c.series[0].values, vec![Some(10.), None, None]);
    assert!(chart_text(&out).contains("R&amp;D &lt;x&gt;"));
}
