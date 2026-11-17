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
        CHART.replace("lineChart", "pieChart"),
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

// ---------------------------------------------------------------------------
// bars, bars with a line, a second value axis
// ---------------------------------------------------------------------------

/// One series the way Excel writes it: a name, categories and values that
/// point at cells of Alpha, with the values the file cached.
fn ser(idx: usize, values: &str, extra: &str) -> String {
    format!(
        r#"<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/><c:tx><c:strRef><c:f>Alpha!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Series {idx}</c:v></c:pt></c:strCache></c:strRef></c:tx>{extra}<c:cat><c:strRef><c:f>Alpha!$A$1:$A$2</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>FY24</c:v></c:pt><c:pt idx="1"><c:v>FY25</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Alpha!{values}</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"#
    )
}

fn chart_space(plot: &str) -> String {
    format!(
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:layout/>{plot}</c:plotArea><c:legend><c:legendPos val="r"/></c:legend></c:chart></c:chartSpace>"#
    )
}

#[test]
fn a_column_chart_is_read_with_its_fills_and_grouping() {
    let fill = r#"<c:spPr><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>"#;
    let theme = r#"<c:spPr><a:solidFill><a:schemeClr val="accent2"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></c:spPr>"#;
    let plot = format!(
        r##"<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/>{}{}{}<c:gapWidth val="50"/><c:overlap val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/></c:catAx><c:valAx><c:axId val="2"/><c:numFmt formatCode="#,##0" sourceLinked="0"/></c:valAx>"##,
        ser(0, "$B$1:$B$2", fill),
        ser(1, "$C$1:$C$2", theme),
        ser(2, "$D$1:$D$2", "")
    );
    let c = read(&fixture(&chart_space(&plot))).unwrap().remove(0);
    assert_eq!(c.unsupported, None);
    assert_eq!(c.kind, "barChart");
    assert_eq!((c.bar_dir.as_str(), c.bar_grouping.as_str()), ("col", "stacked"));
    assert_eq!(c.series.len(), 3);
    assert!(c.series.iter().all(|s| s.kind == "bar" && !s.secondary && !s.markers));
    // A bar is its fill, not its outline; a theme accent is the Office
    // palette's; nothing stated is the palette by the series' index.
    assert_eq!(c.series[0].color, "#1F4E79");
    assert_eq!(c.series[1].color, "#ED7D31");
    assert_eq!(c.series[2].color, "#A5A5A5");
    assert_eq!(c.series[1].values, vec![Some(1.), Some(2.)]);
    assert_eq!(c.number_format, "#,##0");
    assert_eq!((c.bar_gap_width, c.bar_overlap), (Some(50.), Some(100.)));
    // Lying bars, stacked to 100%.
    let lying = plot.replace(r#"barDir val="col""#, r#"barDir val="bar""#).replace("stacked", "percentStacked");
    let c = read(&fixture(&chart_space(&lying))).unwrap().remove(0);
    assert_eq!((c.unsupported, c.bar_dir.as_str(), c.bar_grouping.as_str()), (None, "bar", "percentStacked"));
}

#[test]
fn columns_with_a_line_on_a_second_axis() {
    let plot = format!(
        r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>{}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:lineChart><c:grouping val="standard"/>{}<c:axId val="3"/><c:axId val="4"/></c:lineChart><c:catAx><c:axId val="1"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:min val="0"/></c:scaling><c:axPos val="l"/><c:crosses val="autoZero"/></c:valAx><c:catAx><c:axId val="3"/><c:delete val="1"/></c:catAx><c:valAx><c:axId val="4"/><c:scaling><c:max val="1"/></c:scaling><c:numFmt formatCode="0%" sourceLinked="0"/><c:axPos val="r"/><c:crosses val="max"/></c:valAx>"#,
        ser(0, "$B$1:$B$2", ""),
        ser(1, "$C$1:$C$2", r#"<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="circle"/></c:marker><c:smooth val="1"/>"#)
    );
    let c = read(&fixture(&chart_space(&plot))).unwrap().remove(0);
    assert_eq!(c.unsupported, None);
    assert_eq!(c.series.iter().map(|s| (s.kind.as_str(), s.secondary)).collect::<Vec<_>>(), vec![("bar", false), ("line", true)]);
    assert_eq!((c.series[1].color.as_str(), c.series[1].width, c.series[1].markers, c.series[1].smooth), ("#C00000", 3., true, true));
    // An axis that follows its source is labelled in the format of the
    // cells plotted against it.
    let linked = plot
        .replace(r#"<c:numFmt formatCode="0%" sourceLinked="0"/>"#, r#"<c:numFmt formatCode="General" sourceLinked="1"/>"#)
        .replace(r#"<c:numRef><c:f>Alpha!$C$1:$C$2</c:f><c:numCache>"#, r#"<c:numRef><c:f>Alpha!$C$1:$C$2</c:f><c:numCache><c:formatCode>0.0%</c:formatCode>"#);
    let l = read(&fixture(&chart_space(&linked))).unwrap().remove(0);
    assert_eq!((l.series[1].format.as_str(), l.number_format2.as_str(), l.number_format.as_str()), ("0.0%", "0.0%", ""));
    // Each axis keeps its own scale and format.
    assert_eq!((c.y_min, c.y_max, c.number_format.as_str()), (Some(0.), None, ""));
    assert_eq!((c.y2_min, c.y2_max, c.number_format2.as_str()), (None, Some(1.), "0%"));
}

#[test]
fn what_is_not_drawn_says_so() {
    let pie = format!(r#"<c:pieChart>{}</c:pieChart>"#, ser(0, "$B$1:$B$2", ""));
    assert!(read(&fixture(&chart_space(&pie))).unwrap()[0].unsupported.as_deref().unwrap().contains("not displayed"));
    let stacked_lines = format!(r#"<c:lineChart><c:grouping val="stacked"/>{}</c:lineChart><c:valAx><c:axId val="2"/></c:valAx>"#, ser(0, "$B$1:$B$2", ""));
    assert!(read(&fixture(&chart_space(&stacked_lines))).unwrap()[0].unsupported.as_deref().unwrap().contains("Stacked lines"));
    // A date axis is drawn: its labels are the dates as the sheet shows them.
    let dated = format!(r#"<c:lineChart><c:grouping val="standard"/>{}</c:lineChart><c:dateAx><c:axId val="1"/></c:dateAx><c:valAx><c:axId val="2"/></c:valAx>"#, ser(0, "$B$1:$B$2", ""));
    assert_eq!(read(&fixture(&chart_space(&dated))).unwrap()[0].unsupported, None);
}

#[test]
fn live_data_replaces_the_cache_and_ranges_follow_rows_and_names() {
    let plot = format!(r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>{}</c:barChart><c:valAx><c:axId val="2"/></c:valAx>"#, ser(0, "$B$1:$B$2", ""));
    let mut charts = read(&fixture(&chart_space(&plot))).unwrap();
    // The cells as they are now: text for names and categories, numbers for values.
    let cells = |r: &Range, text: bool| -> Option<Vec<serde_json::Value>> {
        if r.sheet != "Alpha" {
            return None;
        }
        Some(r.cells().map(|(row, col)| if text { serde_json::json!(format!("r{row}c{col}")) } else { serde_json::json!((row * 10 + col) as f64) }).collect())
    };
    fill_live(&mut charts, &cells);
    assert_eq!(charts[0].series[0].name, "r0c0");
    assert_eq!(charts[0].series[0].categories, vec![serde_json::json!("r0c0"), serde_json::json!("r1c0")]);
    assert_eq!(charts[0].series[0].values, vec![Some(1.), Some(11.)]);
    assert_eq!(charts[0].unsupported, None);

    // Two rows inserted at the top of the sheet: ranges and the anchor move down.
    let row_before = charts[0].anchor.from.row;
    shift(&mut charts, "Alpha", &crate::refs::RowColShift { rows: true, start: 0, count: 2, insert: true }).unwrap();
    assert_eq!(charts[0].series[0].values_ref.as_deref(), Some("Alpha!$B$3:$B$4"));
    assert_eq!(charts[0].anchor.from.row, row_before + 2);
    // A rename is followed; a chart whose sheet cannot be read says so.
    rename_sheet(&mut charts, "Alpha", "P&L").unwrap();
    assert_eq!(charts[0].sheet, "P&L");
    assert_eq!(charts[0].series[0].values_ref.as_deref(), Some("'P&L'!$B$3:$B$4"));
    fill_live(&mut charts, &cells);
    assert!(charts[0].unsupported.as_deref().unwrap().contains("cannot be resolved"));
    remove_sheet(&mut charts, "P&L");
    assert!(charts.is_empty());
}

/// A spec for columns of B1:B2 with a line of C1:C2 on a second axis.
fn spec() -> ChartSpec {
    ChartSpec {
        id: "new:1".into(),
        sheet: "Alpha".into(),
        row: 5,
        col: 6,
        title: "Revenue & margin".into(),
        kind: "column".into(),
        categories: Some("A1:A2".into()),
        series: vec![
            // Asked for first, written after the bars: the file keeps bars first.
            SeriesSpec { name: Some("Margin <%>".into()), values: "Alpha!C1:C2".into(), kind: Some("line".into()), secondary: true, ..Default::default() },
            SeriesSpec { name_ref: Some("D1".into()), values: "$B$1:$B$2".into(), color: Some("#1f4e79".into()), ..Default::default() },
        ],
        ..Default::default()
    }
}

/// A workbook that has every sheet asked for, as asked, except "Gone".
fn as_spelled(name: &str) -> Option<String> {
    (name != "Gone").then(|| name.to_string())
}

/// The chart with the data a session would have read for it.
fn filled(spec: &ChartSpec) -> Chart {
    let mut charts = vec![build(spec, &as_spelled).unwrap()];
    fill_live(&mut charts, &|r: &Range, text: bool| {
        Some(r.cells().map(|(row, col)| if text { serde_json::json!(format!("FY{}", 2024 + row + col)) } else { serde_json::json!((row * 10 + col) as f64 + 0.5) }).collect())
    });
    link_formats(&mut charts, &|r: &Range| Some(if r.c0 == 2 { "0.0%" } else { "#,##0" }.to_string()));
    charts.remove(0)
}

fn patch_with(chart: &Chart) -> String {
    serde_json::json!({ "version": 1, "charts": [chart] }).to_string()
}

#[test]
fn a_spec_becomes_a_chart_or_says_why_not() {
    let c = build(&spec(), &as_spelled).unwrap();
    assert_eq!((c.kind.as_str(), c.bar_dir.as_str(), c.bar_grouping.as_str(), c.part.as_str()), ("barChart", "col", "clustered", ""));
    assert_eq!(c.series.iter().map(|s| (s.kind.as_str(), s.secondary, s.color.as_str())).collect::<Vec<_>>(), vec![("bar", false, "#1F4E79"), ("line", true, "#ED7D31")]);
    assert_eq!(c.series[0].values_ref.as_deref(), Some("Alpha!$B$1:$B$2"));
    assert_eq!(c.series[0].name_ref.as_deref(), Some("Alpha!$D$1"));
    assert_eq!(c.series[0].categories_ref.as_deref(), Some("Alpha!$A$1:$A$2"));
    assert_eq!((c.anchor.from.row, c.anchor.from.col, c.anchor.width, c.anchor.height, c.legend), (5, 6, 480., 288., true));
    // A sheet name that needs quotes gets them.
    let quoted = build(&ChartSpec { sheet: "P&L 2026".into(), ..spec() }, &as_spelled).unwrap();
    assert_eq!(quoted.series[0].categories_ref.as_deref(), Some("'P&L 2026'!$A$1:$A$2"));
    // Names come out as the workbook spells them.
    let spelled = build(&ChartSpec { sheet: "alpha".into(), ..spec() }, &|name: &str| Some(name.to_uppercase())).unwrap();
    assert_eq!((spelled.sheet.as_str(), spelled.series[0].values_ref.as_deref(), spelled.series[1].values_ref.as_deref()), ("ALPHA", Some("ALPHA!$B$1:$B$2"), Some("ALPHA!$C$1:$C$2")));

    let why = |change: &dyn Fn(&mut ChartSpec)| {
        let mut s = spec();
        change(&mut s);
        build(&s, &as_spelled).unwrap_err()
    };
    assert!(why(&|s| s.sheet = "Gone".into()).contains("no sheet \"Gone\" to put the chart on"));
    assert!(why(&|s| s.series[0].values = "Gone!C1:C2".into()).contains("series 1: values: there is no sheet \"Gone\""));
    assert!(why(&|s| s.kind = "pie".into()).contains("column, bar, line"));
    assert!(why(&|s| s.series[1].values = "B1:C2".into()).contains("one row or one column"));
    assert!(why(&|s| s.series[1].values = "B1:B3".into()).contains("same length"));
    assert!(why(&|s| s.kind = "bar".into()).contains("lying bars"));
    assert!(why(&|s| s.series[1].secondary = true).contains("only a line"));
    assert!(why(&|s| s.series.truncate(1)).contains("main axis"));
    assert!(why(&|s| s.series[0].color = Some("blue".into())).contains("#RRGGBB"));
    assert!(why(&|s| s.series.clear()).contains("1 to 24"));
    assert!(why(&|s| { s.kind = "line".into(); s.grouping = "stacked".into(); s.series.remove(1); s.series[0].secondary = false; }).contains("stacked lines"));
}

#[test]
fn a_new_chart_is_written_where_the_sheet_has_no_drawing() {
    let base = crate::tests::fixture();
    let chart = filled(&spec());
    let out = apply_patch_json(&base, &patch_with(&chart)).unwrap();
    crate::validate::validate_package(&out).unwrap();

    // What a reader finds is what was asked for.
    let mut read_back = read(&out).unwrap();
    assert_eq!(read_back.len(), 1);
    let c = read_back.remove(0);
    assert_eq!((c.sheet.as_str(), c.title.as_str(), c.unsupported.clone()), ("Alpha", "Revenue & margin", None));
    assert_eq!((c.bar_dir.as_str(), c.bar_grouping.as_str(), c.bar_gap_width), ("col", "clustered", Some(150.)));
    assert_eq!((c.anchor.from.row, c.anchor.from.col, c.anchor.to.is_none(), c.anchor.width, c.anchor.height), (5, 6, true, 480., 288.));
    for (got, want) in c.series.iter().zip(&chart.series) {
        assert_eq!((&got.name, &got.name_ref, &got.categories_ref, &got.values_ref), (&want.name, &want.name_ref, &want.categories_ref, &want.values_ref));
        assert_eq!((&got.categories, &got.values, &got.color, &got.kind, got.secondary, &got.format), (&want.categories, &want.values, &want.color, &want.kind, want.secondary, &want.format));
    }
    assert_eq!(c.series[1].name, "Margin <%>");
    assert_eq!(c.series[1].width, 3.);
    // Each axis reads in the format of the cells plotted against it.
    assert_eq!((c.number_format.as_str(), c.number_format2.as_str()), ("#,##0", "0.0%"));
    assert_eq!((chart.number_format.as_str(), chart.number_format2.as_str()), ("#,##0", "0.0%"));

    // The parts: a chart, a drawing, the sheet's relationship, the types.
    let mut store = PartStore::open(&out).unwrap();
    let text = |store: &mut PartStore, name: &str| String::from_utf8(store.read(name).unwrap()).unwrap();
    let sheet = text(&mut store, "xl/worksheets/sheet1.xml");
    // In the schema's order: after the page margins, at the end.
    assert!(sheet.ends_with(r#"<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></worksheet>"#), "{sheet}");
    assert!(text(&mut store, "xl/worksheets/_rels/sheet1.xml.rels").contains(r#"Target="../drawings/drawing1.xml""#));
    assert!(text(&mut store, "xl/drawings/_rels/drawing1.xml.rels").contains(r#"Target="../charts/chart1.xml""#));
    let types = text(&mut store, "[Content_Types].xml");
    assert!(types.contains(r#"PartName="/xl/charts/chart1.xml""#) && types.contains(r#"PartName="/xl/drawings/drawing1.xml""#));
    let xml = text(&mut store, "xl/charts/chart1.xml");
    assert!(xml.contains("<a:t>Revenue &amp; margin</a:t>") && xml.contains("<c:v>Margin &lt;%&gt;</c:v>"));
    assert!(xml.contains(r#"<c:crosses val="max"/>"#));
    // The axes state the format of the cells they plot, as following them:
    // a reader that does not look the cells up still labels them right.
    assert!(xml.contains(r##"<c:numFmt formatCode="#,##0" sourceLinked="1"/>"##) && xml.contains(r#"<c:numFmt formatCode="0.0%" sourceLinked="1"/>"#));
    // Nothing else moved: every other part is the file's own bytes.
    let before = crate::tests::fixture();
    let mut original = PartStore::open(&before).unwrap();
    for name in ["xl/workbook.xml", "xl/worksheets/sheet2.xml", "xl/styles.xml", "xl/calcChain.xml", "xl/media/blob.bin"] {
        assert_eq!(store.read(name).unwrap(), original.read(name).unwrap(), "{name}");
    }

    // Saved again from the same session (the patch is replayed on the file
    // as it was opened): still one chart.
    let again = apply_patch_json(&base, &patch_with(&chart)).unwrap();
    assert_eq!(read(&again).unwrap().len(), 1);
}

#[test]
fn a_new_chart_joins_the_drawing_the_sheet_has() {
    // The fixture's drawing spells its namespace `x:` and holds one chart.
    let base = fixture(CHART);
    let lines = ChartSpec { kind: "line".into(), title: String::new(), series: vec![SeriesSpec { values: "B1:B2".into(), ..Default::default() }], ..spec() };
    let out = apply_patch_json(&base, &serde_json::json!({ "version": 1, "charts": [filled(&lines), filled(&spec())] }).to_string()).unwrap();
    crate::validate::validate_package(&out).unwrap();
    let charts = read(&out).unwrap();
    assert_eq!(charts.iter().map(|c| c.title.as_str()).collect::<Vec<_>>(), vec!["Daily spend", "", "Revenue & margin"]);
    assert_eq!(charts[1].kind, "lineChart");
    assert!(!charts[1].legend);
    assert!(charts.iter().all(|c| c.unsupported.is_none() && c.sheet == "Alpha"));
    // The chart that was there is untouched, and the new parts took free names.
    assert_eq!(chart_text(&out), chart_text(&base));
    let mut store = PartStore::open(&out).unwrap();
    assert!(store.exists("xl/charts/chart1.xml") && store.exists("xl/charts/chart2.xml"));
    let drawing = String::from_utf8(store.read("xl/drawings/drawing1.xml").unwrap()).unwrap();
    assert_eq!(drawing.matches("<xdr:oneCellAnchor ").count(), 2);
    assert!(drawing.contains(r#"<xdr:cNvPr id="2" name="Chart 2"/>"#) && drawing.contains(r#"<xdr:cNvPr id="3" name="Chart 3"/>"#));
    assert!(drawing.ends_with("</x:wsDr>"));
    // A sheet that is not there is refused, and nothing is written.
    let lost = Chart { sheet: "Elsewhere".into(), ..filled(&spec()) };
    assert!(apply_patch_json(&base, &patch_with(&lost)).is_err());
}

#[test]
fn a_new_chart_follows_the_rows_inserted_before_it_was_saved() {
    // Session: chart made, then two rows inserted above its data. The engine
    // moved the chart's ranges and place; the patch carries both.
    let mut charts = vec![filled(&spec())];
    shift(&mut charts, "Alpha", &crate::refs::RowColShift { rows: true, start: 0, count: 2, insert: true }).unwrap();
    let patch = serde_json::json!({ "version": 1, "rowColOps": [{ "op": "insertRows", "sheet": "Alpha", "before": 0, "count": 2 }], "charts": charts }).to_string();
    let out = apply_patch_json(&crate::tests::fixture(), &patch).unwrap();
    crate::validate::validate_package(&out).unwrap();
    let c = read(&out).unwrap().remove(0);
    assert_eq!(c.series[0].values_ref.as_deref(), Some("Alpha!$B$3:$B$4"));
    assert_eq!(c.anchor.from.row, 7);
}
