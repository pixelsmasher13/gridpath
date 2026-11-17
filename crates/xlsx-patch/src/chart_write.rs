//! Write a chart made in this session into the package: the chart part, its
//! place in the sheet's drawing (made when the sheet has none), and the
//! relationships and content types that tie them together. Nothing that is
//! already in the file is regenerated: an existing drawing gets one more
//! anchor, an existing worksheet one `<drawing>` element.
use crate::{
    charts::{escape, relationships, rels_path, resolve, sheets, tree, Chart, Series},
    package::PartStore,
    sheet_xml::WORKSHEET_CHILD_ORDER,
    structure::{insert_before_container_end, max_rid},
    workbook_xml::{local_name, reader},
    PatchError,
};
use quick_xml::events::{BytesStart, Event};
use serde_json::Value;

const NS_C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const NS_RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_DRAWING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const REL_CHART: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const CT_DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CT_CHART: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const DECL: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
/// Pixels to the file's unit (EMU).
const EMU: f64 = 9525.;

/// Adds `chart` to the sheet it names.
pub fn add(store: &mut PartStore, chart: &Chart) -> Result<(), PatchError> {
    if chart.unsupported.is_some() || chart.series.is_empty() {
        return Err(PatchError::BadValue(format!("chart {:?} cannot be written: it has nothing to draw", chart.title)));
    }
    let sheet_path = sheets(store)?
        .into_iter()
        .find(|(name, _)| name.eq_ignore_ascii_case(&chart.sheet))
        .map(|(_, path)| path)
        .ok_or_else(|| PatchError::SheetNotFound(chart.sheet.clone()))?;
    let workbook = relationships(store, "")?
        .into_iter()
        .find(|(_, kind, _)| kind == "officeDocument")
        .map(|(_, _, path)| path)
        .unwrap_or_else(|| "xl/workbook.xml".into());
    let dir = workbook.rsplit_once('/').map(|(d, _)| format!("{d}/")).unwrap_or_default();

    let part = first_free(store, |n| format!("{dir}charts/chart{n}.xml"));
    store.write(&part, chart_xml(chart).into_bytes());
    content_type(store, &part, CT_CHART)?;

    let drawing = match relationships(store, &sheet_path)?.into_iter().find(|(_, kind, _)| kind == "drawing") {
        Some((_, _, path)) => path,
        None => new_drawing(store, &sheet_path, &dir)?,
    };
    let rid = relate(store, &drawing, REL_CHART, &part)?;
    let xml = store.read(&drawing)?;
    store.write(&drawing, with_anchor(&xml, chart, &rid)?);
    Ok(())
}

fn first_free(store: &mut PartStore, name: impl Fn(u32) -> String) -> String {
    let mut n = 1;
    loop {
        let candidate = name(n);
        if !store.exists(&candidate) {
            return candidate;
        }
        n += 1;
    }
}

/// `to` as a relationship of the part `from` spells it: `../charts/chart3.xml`.
fn relative(from: &str, to: &str) -> String {
    let dir: Vec<&str> = from.split('/').collect();
    let dir = &dir[..dir.len() - 1];
    let target: Vec<&str> = to.split('/').collect();
    let common = dir.iter().zip(&target[..target.len() - 1]).take_while(|(a, b)| a == b).count();
    format!("{}{}", "../".repeat(dir.len() - common), target[common..].join("/"))
}

/// Adds a relationship from `source` to the part `target`; gives its id.
fn relate(store: &mut PartStore, source: &str, kind: &str, target: &str) -> Result<String, PatchError> {
    let path = rels_path(source);
    let xml = if store.exists(&path) {
        store.read(&path)?
    } else {
        format!(r#"{DECL}<Relationships xmlns="{NS_RELS}"></Relationships>"#).into_bytes()
    };
    let rid = format!("rId{}", max_rid(&xml)? + 1);
    let spelled = relative(source, target);
    debug_assert_eq!(resolve(source, &spelled), target);
    let mut rel = BytesStart::new("Relationship");
    rel.push_attribute(("Id", rid.as_str()));
    rel.push_attribute(("Type", kind));
    rel.push_attribute(("Target", spelled.as_str()));
    store.write(&path, insert_before_container_end(&xml, b"Relationships", rel)?);
    Ok(rid)
}

fn content_type(store: &mut PartStore, part: &str, kind: &str) -> Result<(), PatchError> {
    let xml = store.read("[Content_Types].xml")?;
    let mut entry = BytesStart::new("Override");
    entry.push_attribute(("PartName", format!("/{part}").as_str()));
    entry.push_attribute(("ContentType", kind));
    store.write("[Content_Types].xml", insert_before_container_end(&xml, b"Types", entry)?);
    Ok(())
}

/// A drawing for a sheet that has none: the part, the sheet's relationship
/// to it and the `<drawing>` element in the worksheet.
fn new_drawing(store: &mut PartStore, sheet_path: &str, dir: &str) -> Result<String, PatchError> {
    let part = first_free(store, |n| format!("{dir}drawings/drawing{n}.xml"));
    store.write(&part, format!(r#"{DECL}<xdr:wsDr xmlns:xdr="{NS_XDR}" xmlns:a="{NS_A}"></xdr:wsDr>"#).into_bytes());
    content_type(store, &part, CT_DRAWING)?;
    let rid = relate(store, sheet_path, REL_DRAWING, &part)?;
    let xml = store.read(sheet_path)?;
    store.write(sheet_path, with_drawing_element(&xml, &rid)?);
    Ok(part)
}

/// The worksheet with `<drawing r:id>` at its place in the schema's order.
/// One pass over the events: a sheet can hold millions of cells.
fn with_drawing_element(xml: &[u8], rid: &str) -> Result<Vec<u8>, PatchError> {
    let position = |name: &[u8]| WORKSHEET_CHILD_ORDER.iter().position(|n| *n == name);
    let drawing = position(b"drawing").unwrap_or(usize::MAX);
    let bom = if xml.starts_with(&[0xef, 0xbb, 0xbf]) { 3 } else { 0 };
    let mut r = reader(&xml[bom..]);
    let mut buf = vec![];
    let mut depth = 0usize;
    let mut prefix = String::new();
    let mut at = None;
    loop {
        let start = r.buffer_position() as usize + bom;
        let event = r.read_event_into(&mut buf).map_err(|e| PatchError::Xml(format!("worksheet: {e}")))?;
        match &event {
            Event::Start(e) | Event::Empty(e) => {
                let qualified = e.name();
                let name = local_name(qualified.as_ref());
                if depth == 0 {
                    prefix = String::from_utf8_lossy(qualified.as_ref())
                        .rsplit_once(':')
                        .map(|(p, _)| format!("{p}:"))
                        .unwrap_or_default();
                } else if depth == 1 {
                    if name == b"drawing" {
                        return Err(PatchError::Unsupported(
                            "the sheet has a drawing that could not be followed, so a chart cannot be added to it".into(),
                        ));
                    }
                    // What comes after a drawing, or is not in the schema's
                    // list (wrapped controls and objects, which follow it).
                    if position(name).map_or(true, |i| i > drawing) {
                        at = Some(start);
                        break;
                    }
                }
                if matches!(event, Event::Start(_)) {
                    depth += 1;
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    at = Some(start);
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    let at = at.ok_or_else(|| PatchError::Xml("worksheet has no end".into()))?;
    let element = format!(r#"<{prefix}drawing xmlns:r="{NS_R}" r:id="{rid}"/>"#);
    let mut out = Vec::with_capacity(xml.len() + element.len());
    out.extend_from_slice(&xml[..at]);
    out.extend_from_slice(element.as_bytes());
    out.extend_from_slice(&xml[at..]);
    Ok(out)
}

/// The drawing with one more anchor: the chart's frame at its cell.
fn with_anchor(xml: &[u8], chart: &Chart, rid: &str) -> Result<Vec<u8>, PatchError> {
    let doc = tree(xml)?;
    let root = doc
        .children
        .iter()
        .find(|n| n.name == "wsDr")
        .ok_or_else(|| PatchError::Xml("drawing has no root".into()))?;
    // Shape ids are unique within a drawing.
    let id = root
        .all("cNvPr")
        .iter()
        .filter_map(|n| n.attrs.get("id")?.parse::<u64>().ok())
        .max()
        .unwrap_or(1)
        + 1;
    let a = &chart.anchor;
    let emu = |px: f64| (px.max(0.) * EMU).round() as i64;
    let anchor = format!(
        concat!(
            r#"<xdr:oneCellAnchor xmlns:xdr="{xdr}" xmlns:a="{a}">"#,
            "<xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>{dx}</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>{dy}</xdr:rowOff></xdr:from>",
            r#"<xdr:ext cx="{cx}" cy="{cy}"/>"#,
            r#"<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{id}" name="Chart {id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"#,
            r#"<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>"#,
            r#"<a:graphic><a:graphicData uri="{c}"><c:chart xmlns:c="{c}" xmlns:r="{r}" r:id="{rid}"/></a:graphicData></a:graphic>"#,
            "</xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>"
        ),
        xdr = NS_XDR,
        a = NS_A,
        c = NS_C,
        r = NS_R,
        col = a.from.col,
        row = a.from.row,
        dx = emu(a.from.dx),
        dy = emu(a.from.dy),
        cx = emu(a.width),
        cy = emu(a.height),
        id = id,
        rid = rid,
    );
    let mut out = Vec::with_capacity(xml.len() + anchor.len());
    if root.close == root.end {
        // `<xdr:wsDr …/>`: an empty drawing.
        let open = std::str::from_utf8(&xml[root.start..root.end]).map_err(|e| PatchError::Xml(e.to_string()))?;
        let open = open.strip_suffix("/>").ok_or_else(|| PatchError::Xml("drawing root".into()))?;
        out.extend_from_slice(&xml[..root.start]);
        out.extend_from_slice(format!("{open}>{anchor}</{}>", root.qualified).as_bytes());
        out.extend_from_slice(&xml[root.end..]);
    } else {
        out.extend_from_slice(&xml[..root.close]);
        out.extend_from_slice(anchor.as_bytes());
        out.extend_from_slice(&xml[root.close..]);
    }
    Ok(out)
}

// Axis ids: the main pair, and the pair of the second value axis.
const CAT: u32 = 500_000_001;
const VAL: u32 = 500_000_002;
const CAT2: u32 = 500_000_003;
const VAL2: u32 = 500_000_004;

fn hex(color: &str) -> String {
    let c = color.trim_start_matches('#');
    if c.len() == 6 && c.chars().all(|ch| ch.is_ascii_hexdigit()) { c.to_ascii_uppercase() } else { "4472C4".into() }
}

fn text_of(v: &Value) -> Option<String> {
    match v {
        Value::Null => None,
        Value::String(s) => Some(s.clone()),
        other => Some(other.to_string()),
    }
}

fn series_xml(x: &mut String, index: usize, s: &Series) {
    let bar = s.kind == "bar";
    x.push_str(&format!(r#"<c:ser><c:idx val="{index}"/><c:order val="{index}"/>"#));
    match &s.name_ref {
        Some(f) => x.push_str(&format!(
            r#"<c:tx><c:strRef><c:f>{}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{}</c:v></c:pt></c:strCache></c:strRef></c:tx>"#,
            escape(f),
            escape(&s.name)
        )),
        None => x.push_str(&format!("<c:tx><c:v>{}</c:v></c:tx>", escape(&s.name))),
    }
    let color = hex(&s.color);
    if bar {
        x.push_str(&format!(r#"<c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>"#));
    } else {
        x.push_str(&format!(
            r#"<c:spPr><a:ln w="{}" cap="rnd"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>"#,
            (s.width.max(0.25) * EMU).round() as i64
        ));
    }
    if let Some(f) = &s.categories_ref {
        x.push_str(&format!(r#"<c:cat><c:strRef><c:f>{}</c:f><c:strCache><c:ptCount val="{}"/>"#, escape(f), s.categories.len()));
        for (i, v) in s.categories.iter().enumerate() {
            if let Some(t) = text_of(v) {
                x.push_str(&format!(r#"<c:pt idx="{i}"><c:v>{}</c:v></c:pt>"#, escape(&t)));
            }
        }
        x.push_str("</c:strCache></c:strRef></c:cat>");
    }
    if let Some(f) = &s.values_ref {
        let format = if s.format.is_empty() { "General" } else { &s.format };
        x.push_str(&format!(
            r#"<c:val><c:numRef><c:f>{}</c:f><c:numCache><c:formatCode>{}</c:formatCode><c:ptCount val="{}"/>"#,
            escape(f),
            escape(format),
            s.values.len()
        ));
        for (i, v) in s.values.iter().enumerate() {
            if let Some(n) = v.filter(|n| n.is_finite()) {
                x.push_str(&format!(r#"<c:pt idx="{i}"><c:v>{n}</c:v></c:pt>"#));
            }
        }
        x.push_str("</c:numCache></c:numRef></c:val>");
    }
    if !bar {
        x.push_str(r#"<c:smooth val="0"/>"#);
    }
    x.push_str("</c:ser>");
}

fn category_axis(x: &mut String, id: u32, crosses: u32, position: &str, shown: bool) {
    x.push_str(&format!(
        concat!(
            r#"<c:catAx><c:axId val="{id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="{delete}"/><c:axPos val="{position}"/>"#,
            r#"<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>"#,
            r#"<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr>"#,
            r#"<c:crossAx val="{crosses}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>"#
        ),
        id = id,
        crosses = crosses,
        position = position,
        delete = if shown { 0 } else { 1 },
    ));
}

/// `format` is the axis' number format and whether it is the format of the
/// cells plotted against it (which is how Excel writes a new chart's axis:
/// the cells' format, marked as following them).
fn value_axis(x: &mut String, id: u32, crosses: u32, position: &str, format: (&str, bool), second: bool) {
    x.push_str(&format!(
        r#"<c:valAx><c:axId val="{id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="{position}"/>"#
    ));
    if !second {
        x.push_str(r#"<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>"#);
    }
    let (code, linked) = format;
    x.push_str(&format!(
        r#"<c:numFmt formatCode="{}" sourceLinked="{}"/>"#,
        escape(if code.is_empty() { "General" } else { code }),
        if linked { 1 } else { 0 }
    ));
    x.push_str(&format!(
        concat!(
            r#"<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>"#,
            r#"<c:crossAx val="{crosses}"/><c:crosses val="{at}"/><c:crossBetween val="between"/></c:valAx>"#
        ),
        crosses = crosses,
        at = if second { "max" } else { "autoZero" },
    ));
}

/// The chart part. Element order is the schema's: Excel refuses a chart
/// whose children are out of order.
pub fn chart_xml(chart: &Chart) -> String {
    let mut x = String::with_capacity(4096);
    x.push_str(DECL);
    x.push_str(&format!(
        r#"<c:chartSpace xmlns:c="{NS_C}" xmlns:a="{NS_A}" xmlns:r="{NS_R}"><c:date1904 val="0"/><c:roundedCorners val="0"/><c:chart>"#
    ));
    if chart.title.is_empty() {
        x.push_str(r#"<c:autoTitleDeleted val="1"/>"#);
    } else {
        x.push_str(&format!(
            concat!(
                r#"<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr>"#,
                r#"<a:r><a:rPr lang="en-US" sz="1400" b="0"/><a:t>{}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>"#
            ),
            escape(&chart.title)
        ));
    }
    x.push_str("<c:plotArea><c:layout/>");
    let lying = chart.bar_dir == "bar";
    let numbered: Vec<(usize, &Series)> = chart.series.iter().enumerate().collect();
    let bars: Vec<_> = numbered.iter().filter(|(_, s)| s.kind == "bar").collect();
    if !bars.is_empty() {
        let grouping = if chart.bar_grouping.is_empty() { "clustered" } else { &chart.bar_grouping };
        x.push_str(&format!(
            r#"<c:barChart><c:barDir val="{}"/><c:grouping val="{grouping}"/><c:varyColors val="0"/>"#,
            if lying { "bar" } else { "col" }
        ));
        for (i, s) in &bars {
            series_xml(&mut x, *i, s);
        }
        x.push_str(&format!(r#"<c:gapWidth val="{}"/>"#, chart.bar_gap_width.unwrap_or(150.).clamp(0., 500.).round() as i64));
        if let Some(overlap) = chart.bar_overlap {
            x.push_str(&format!(r#"<c:overlap val="{}"/>"#, overlap.clamp(-100., 100.).round() as i64));
        }
        x.push_str(&format!(r#"<c:axId val="{CAT}"/><c:axId val="{VAL}"/></c:barChart>"#));
    }
    let mut second = false;
    for on_second in [false, true] {
        let lines: Vec<_> = numbered.iter().filter(|(_, s)| s.kind != "bar" && s.secondary == on_second).collect();
        if lines.is_empty() {
            continue;
        }
        second |= on_second;
        x.push_str(r#"<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>"#);
        for (i, s) in &lines {
            series_xml(&mut x, *i, s);
        }
        let (cat, val) = if on_second { (CAT2, VAL2) } else { (CAT, VAL) };
        x.push_str(&format!(r#"<c:marker val="1"/><c:axId val="{cat}"/><c:axId val="{val}"/></c:lineChart>"#));
    }
    category_axis(&mut x, CAT, VAL, if lying { "l" } else { "b" }, true);
    value_axis(
        &mut x,
        VAL,
        CAT,
        if lying { "b" } else { "l" },
        // Stacked to 100%, the axis reads in percent whatever the cells hold.
        if chart.bar_grouping == "percentStacked" { ("0%", false) } else { (chart.number_format.as_str(), true) },
        false,
    );
    if second {
        category_axis(&mut x, CAT2, VAL2, "b", false);
        value_axis(&mut x, VAL2, CAT2, "r", (chart.number_format2.as_str(), true), true);
    }
    x.push_str("</c:plotArea>");
    if chart.legend {
        let position = match chart.legend_position.as_str() {
            p @ ("b" | "t" | "l" | "r" | "tr") => p,
            _ => "b",
        };
        x.push_str(&format!(r#"<c:legend><c:legendPos val="{position}"/><c:overlay val="0"/></c:legend>"#));
    }
    x.push_str(r#"<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>"#);
    x
}
