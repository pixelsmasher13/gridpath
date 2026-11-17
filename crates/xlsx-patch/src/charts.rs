//! Read native charts through OPC relationships; refresh only their data caches.
//! The renderer consumes a deliberately bounded subset. Unsupported objects stay
//! in the package and are represented explicitly, never converted to pictures.
use crate::{
    package::PartStore,
    refs,
    workbook_xml::{attr_value, local_name, reader},
    PatchError,
};
use quick_xml::events::Event;
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::collections::{HashMap, HashSet};

pub const MAX_POINTS: usize = 10_000;

#[derive(Clone, Debug, Default)]
pub(crate) struct Node {
    pub(crate) name: String,
    pub(crate) qualified: String,
    pub(crate) attrs: HashMap<String, String>,
    pub(crate) text: String,
    pub(crate) children: Vec<Node>,
    pub(crate) start: usize,
    pub(crate) end: usize,
    pub(crate) open_end: usize,
    pub(crate) close: usize,
}
impl Node {
    pub(crate) fn child(&self, name: &str) -> Option<&Node> {
        self.children.iter().find(|n| n.name == name)
    }
    fn descendants<'a>(&'a self, name: &str, out: &mut Vec<&'a Node>) {
        if self.name == name {
            out.push(self);
        }
        for c in &self.children {
            c.descendants(name, out);
        }
    }
    pub(crate) fn all(&self, name: &str) -> Vec<&Node> {
        let mut v = vec![];
        self.descendants(name, &mut v);
        v
    }
    pub(crate) fn first(&self, name: &str) -> Option<&Node> {
        if self.name == name {
            return Some(self);
        }
        self.children.iter().find_map(|c| c.first(name))
    }
    pub(crate) fn val(&self) -> Option<&str> {
        self.attrs.get("val").map(String::as_str)
    }
    fn number(&self) -> f64 {
        self.text.parse().unwrap_or(0.)
    }
}
pub(crate) fn tree(xml: &[u8]) -> Result<Node, PatchError> {
    // quick-xml excludes a UTF-8 BOM from buffer_position. Keep replacement
    // offsets in the original part's byte coordinates.
    let bom = if xml.starts_with(&[0xef, 0xbb, 0xbf]) {
        3
    } else {
        0
    };
    let mut r = reader(&xml[bom..]);
    let mut buf = vec![];
    let mut stack = vec![Node::default()];
    loop {
        let start = r.buffer_position() as usize + bom;
        let event = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("chart: {e}")))?;
        let end = r.buffer_position() as usize + bom;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(e) | Event::Empty(e) => {
                let mut node = Node {
                    name: String::from_utf8_lossy(local_name(e.name().as_ref())).into_owned(),
                    qualified: String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                    start,
                    end,
                    open_end: end,
                    close: end,
                    ..Default::default()
                };
                for a in e.attributes().with_checks(false).flatten() {
                    node.attrs.insert(
                        String::from_utf8_lossy(local_name(a.key.as_ref())).into_owned(),
                        a.unescape_value()
                            .map_err(|e| PatchError::Xml(e.to_string()))?
                            .into_owned(),
                    );
                }
                if empty {
                    stack.last_mut().unwrap().children.push(node);
                } else {
                    stack.push(node);
                }
            }
            Event::End(_) => {
                if stack.len() <= 1 {
                    return Err(PatchError::Xml("unexpected chart closing element".into()));
                }
                let mut n = stack.pop().unwrap();
                n.close = start;
                n.end = end;
                stack.last_mut().unwrap().children.push(n);
            }
            Event::Text(e) => stack
                .last_mut()
                .unwrap()
                .text
                .push_str(&e.unescape().map_err(|e| PatchError::Xml(e.to_string()))?),
            Event::CData(e) => stack
                .last_mut()
                .unwrap()
                .text
                .push_str(&String::from_utf8_lossy(&e)),
            Event::Eof => break,
            _ => {}
        }
        if stack.len() > 128 {
            return Err(PatchError::Xml("chart XML nesting limit".into()));
        }
        buf.clear();
    }
    if stack.len() != 1 {
        return Err(PatchError::Xml("incomplete chart XML".into()));
    }
    Ok(stack.remove(0))
}

pub fn resolve(source: &str, target: &str) -> String {
    let base = if target.starts_with('/') {
        String::new()
    } else {
        source
            .rsplit_once('/')
            .map(|(d, _)| format!("{d}/"))
            .unwrap_or_default()
    };
    let joined = format!("{base}{}", target.trim_start_matches('/'));
    let mut parts = vec![];
    for p in joined.split('/') {
        match p {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(p),
        }
    }
    parts.join("/")
}
pub(crate) fn rels_path(source: &str) -> String {
    let (dir, file) = source
        .rsplit_once('/')
        .map(|(d, f)| (format!("{d}/"), f))
        .unwrap_or((String::new(), source));
    format!("{dir}_rels/{file}.rels")
}
pub(crate) fn relationships(
    store: &mut PartStore,
    source: &str,
) -> Result<Vec<(String, String, String)>, PatchError> {
    let path = if source.is_empty() {
        "_rels/.rels".into()
    } else {
        rels_path(source)
    };
    if !store.exists(&path) {
        return Ok(vec![]);
    }
    let doc = tree(&store.read(&path)?)?;
    Ok(doc
        .all("Relationship")
        .into_iter()
        .filter(|n| n.attrs.get("TargetMode").map(String::as_str) != Some("External"))
        .filter_map(|n| {
            Some((
                n.attrs.get("Id")?.clone(),
                n.attrs.get("Type")?.rsplit('/').next()?.into(),
                resolve(source, n.attrs.get("Target")?),
            ))
        })
        .collect())
}
pub(crate) fn sheets(store: &mut PartStore) -> Result<Vec<(String, String)>, PatchError> {
    let wb = relationships(store, "")?
        .into_iter()
        .find(|(_, t, _)| t == "officeDocument")
        .map(|(_, _, p)| p)
        .ok_or_else(|| PatchError::MissingPart("workbook relationship".into()))?;
    let rels = relationships(store, &wb)?;
    let doc = tree(&store.read(&wb)?)?;
    Ok(doc
        .all("sheet")
        .into_iter()
        .filter_map(|s| {
            let id = s.attrs.get("id")?;
            let (_, _, path) = rels.iter().find(|(r, t, _)| r == id && t == "worksheet")?;
            Some((s.attrs.get("name")?.clone(), path.clone()))
        })
        .collect())
}

#[derive(Clone, Debug, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Marker {
    pub row: u32,
    pub col: u32,
    pub dx: f64,
    pub dy: f64,
}
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Anchor {
    pub from: Marker,
    pub to: Option<Marker>,
    pub width: f64,
    pub height: f64,
    pub absolute: bool,
}
#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Series {
    pub name: String,
    pub name_ref: Option<String>,
    pub categories_ref: Option<String>,
    pub values_ref: Option<String>,
    pub categories: Vec<Value>,
    pub values: Vec<Option<f64>>,
    pub color: String,
    pub width: f64,
    pub markers: bool,
    /// How the series is drawn: "line" or "bar". One chart can hold both
    /// (columns for revenue, a line for the margin).
    #[serde(default = "line_kind")]
    pub kind: String,
    /// Plotted against the second value axis (the one on the right).
    #[serde(default)]
    pub secondary: bool,
    /// A line drawn smoothed.
    #[serde(default)]
    pub smooth: bool,
    /// The number format of the cells the values come from, as the chart
    /// recorded it ("0.0%", "$#,##0"); empty when it did not.
    #[serde(default)]
    pub format: String,
}
fn line_kind() -> String {
    "line".into()
}
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Chart {
    pub id: String,
    pub part: String,
    pub sheet: String,
    pub title: String,
    pub kind: String,
    pub unsupported: Option<String>,
    pub anchor: Anchor,
    pub series: Vec<Series>,
    pub legend: bool,
    pub legend_position: String,
    pub y_min: Option<f64>,
    pub y_max: Option<f64>,
    pub number_format: String,
    pub blanks: String,
    /// Bars: "col" (upright) or "bar" (lying); empty when the chart has none.
    #[serde(default)]
    pub bar_dir: String,
    /// Bars: "clustered", "stacked" or "percentStacked".
    #[serde(default)]
    pub bar_grouping: String,
    /// Bars: the gap between categories, in percent of a bar's width
    /// (Excel's "gap width"; 150 when the file says nothing).
    #[serde(default)]
    pub bar_gap_width: Option<f64>,
    /// Bars side by side: how far each overlaps the one before it, in
    /// percent of a bar's width (negative leaves a gap).
    #[serde(default)]
    pub bar_overlap: Option<f64>,
    /// The second value axis, when series are plotted against one.
    #[serde(default)]
    pub y2_min: Option<f64>,
    #[serde(default)]
    pub y2_max: Option<f64>,
    #[serde(default)]
    pub number_format2: String,
}
#[derive(Clone, Debug)]
pub struct Range {
    pub sheet: String,
    pub r0: u32,
    pub c0: u32,
    pub r1: u32,
    pub c1: u32,
}
impl Range {
    pub fn len(&self) -> usize {
        ((self.r1 - self.r0 + 1) as usize) * ((self.c1 - self.c0 + 1) as usize)
    }
    pub fn cells(&self) -> impl Iterator<Item = (u32, u32)> + '_ {
        (self.r0..=self.r1).flat_map(move |r| (self.c0..=self.c1).map(move |c| (r, c)))
    }
}
pub fn range(formula: &str) -> Option<Range> {
    let (s, a1) = formula.trim_start_matches('=').rsplit_once('!')?;
    if s.len() < 2 && s.starts_with('\'') {
        return None;
    }
    let sheet = if s.starts_with('\'') && s.ends_with('\'') {
        s[1..s.len() - 1].replace("''", "'")
    } else {
        if s.contains(['\'', ' ', '[', ']', ':', ',']) {
            return None;
        }
        s.into()
    };
    if sheet.contains(['[', ']']) {
        return None;
    }
    let (r0, c0, r1, c1) = refs::parse_a1_range(&a1.replace('$', ""))?;
    let r = Range {
        sheet,
        r0,
        c0,
        r1,
        c1,
    };
    if (r0 != r1 && c0 != c1) || r.len() > MAX_POINTS {
        return None;
    }
    Some(r)
}
fn marker(n: Option<&Node>) -> Marker {
    let get = |s: &str| n.and_then(|x| x.child(s)).map(Node::number).unwrap_or(0.);
    Marker {
        row: get("row").max(0.) as u32,
        col: get("col").max(0.) as u32,
        dx: get("colOff") / 9525.,
        dy: get("rowOff") / 9525.,
    }
}
fn cached(n: &Node, numeric: bool) -> Vec<Value> {
    let Some(cache) = n
        .children
        .iter()
        .find(|n| n.name.ends_with("Cache") || n.name.ends_with("Lit"))
    else {
        return vec![];
    };
    let len = cache
        .child("ptCount")
        .and_then(Node::val)
        .and_then(|s| s.parse::<usize>().ok())
        .unwrap_or(0)
        .min(MAX_POINTS);
    let mut values = vec![Value::Null; len];
    for pt in cache.children.iter().filter(|p| p.name == "pt") {
        let Some(i) = pt
            .attrs
            .get("idx")
            .and_then(|s| s.parse::<usize>().ok())
            .filter(|i| *i < MAX_POINTS)
        else {
            continue;
        };
        if values.len() <= i {
            values.resize(i + 1, Value::Null);
        }
        if let Some(v) = pt.child("v") {
            values[i] = if numeric {
                v.text
                    .parse::<f64>()
                    .ok()
                    .filter(|n| n.is_finite())
                    .map(Value::from)
                    .unwrap_or(Value::Null)
            } else {
                Value::String(v.text.clone())
            };
        }
    }
    values
}
fn chart(
    store: &mut PartStore,
    sheet: &str,
    part: &str,
    anchor: Anchor,
    id: usize,
) -> Result<Chart, PatchError> {
    let root = tree(&store.read(part)?)?;
    let plot = root
        .first("plotArea")
        .ok_or_else(|| PatchError::Xml("chart has no plot area".into()))?;
    let plots: Vec<_> = plot
        .children
        .iter()
        .filter(|n| n.name.ends_with("Chart"))
        .collect();
    let kind = plots
        .first()
        .map(|n| n.name.clone())
        .unwrap_or_else(|| "unknown".into());
    // What is drawn: lines, bars (upright or lying; side by side, stacked
    // or stacked to 100%), and the two in one chart, with a second value
    // axis. Everything else stays in the package and is named, not drawn.
    let mut unsupported = None;
    if plots.is_empty() || plots.iter().any(|p| !matches!(p.name.as_str(), "lineChart" | "barChart")) {
        unsupported = Some("This chart type is preserved but not displayed yet.".into());
    }
    if plots.iter().any(|p| {
        p.name == "lineChart" && p.child("grouping").and_then(Node::val).is_some_and(|v| v != "standard")
    }) {
        unsupported = Some("Stacked lines are not displayed yet.".into());
    }
    let bars: Vec<&&Node> = plots.iter().filter(|p| p.name == "barChart").collect();
    let bar_dir = bars
        .first()
        .and_then(|p| p.child("barDir"))
        .and_then(Node::val)
        .unwrap_or(if bars.is_empty() { "" } else { "col" })
        .to_string();
    let bar_grouping = bars
        .first()
        .and_then(|p| p.child("grouping"))
        .and_then(Node::val)
        .unwrap_or(if bars.is_empty() { "" } else { "clustered" })
        .to_string();
    if !bar_grouping.is_empty() && !matches!(bar_grouping.as_str(), "clustered" | "stacked" | "percentStacked") {
        unsupported = Some("This bar layout is not displayed yet.".into());
    }
    // Value axes: the first that does not cross at the far end is the main
    // one; a second one is drawn on the right.
    let value_axes: Vec<&Node> = plot.children.iter().filter(|n| n.name == "valAx").collect();
    let crosses_max = |a: &Node| a.child("crosses").and_then(Node::val) == Some("max");
    let axis = value_axes.iter().find(|a| !crosses_max(a)).or(value_axes.first()).copied();
    let axis2 = value_axes.iter().find(|a| axis.is_some_and(|m| !std::ptr::eq(**a, m))).copied();
    if !plot.all("logBase").is_empty() || value_axes.len() > 2 {
        unsupported = Some("Logarithmic axes and more than two value axes are not displayed yet.".into());
    }
    if axis2.is_some() && bar_dir == "bar" {
        unsupported = Some("Lying bars with a second axis are not displayed yet.".into());
    }
    let axis2_id = axis2.and_then(|a| a.child("axId")).and_then(Node::val).map(str::to_string);
    let colors = PALETTE;
    // A colour as the chart states it: an RGB value, or one of the theme's
    // six accents (the Office palette; tints and shades are not applied).
    let stated_color = |n: Option<&Node>| -> Option<String> {
        let n = n?;
        if let Some(rgb) = n
            .first("srgbClr")
            .and_then(Node::val)
            .filter(|s| s.len() == 6 && s.chars().all(|c| c.is_ascii_hexdigit()))
        {
            return Some(format!("#{rgb}"));
        }
        let accent = n.first("schemeClr").and_then(Node::val)?.strip_prefix("accent")?;
        let i: usize = accent.parse().ok()?;
        colors.get(i.checked_sub(1)?).map(|c| c.to_string())
    };
    let mut series = vec![];
    for p in &plots {
        let is_bar = p.name == "barChart";
        let on_second = axis2_id.as_deref().is_some_and(|id| {
            p.children.iter().any(|n| n.name == "axId" && n.val() == Some(id))
        });
        for s in p.children.iter().filter(|n| n.name == "ser") {
            let i = series.len();
            let cat = s.child("cat");
            let val = s.child("val");
            let cat_ref = cat.and_then(|n| n.first("f")).map(|n| n.text.clone());
            let values_ref = val.and_then(|n| n.first("f")).map(|n| n.text.clone());
            if cat_ref
                .iter()
                .chain(values_ref.iter())
                .any(|f| range(f).is_none())
                || val.is_none()
            {
                unsupported =
                    Some("This chart uses a data reference that is not supported yet.".into());
            }
            let name_ref = s
                .child("tx")
                .and_then(|n| n.first("f"))
                .map(|n| n.text.clone());
            let name = s
                .child("tx")
                .and_then(|n| n.first("v"))
                .map(|n| n.text.clone())
                .unwrap_or_else(|| format!("Series {}", i + 1));
            let sp = s.child("spPr");
            let ln = sp.and_then(|n| n.child("ln"));
            // A bar is its fill; a line is its stroke. Without a stated
            // colour Excel takes the palette by the series' index.
            let index = s
                .child("idx")
                .and_then(Node::val)
                .and_then(|v| v.parse::<usize>().ok())
                .unwrap_or(i);
            let color = if is_bar {
                stated_color(sp.and_then(|n| n.child("solidFill")))
            } else {
                stated_color(ln)
            }
            .unwrap_or_else(|| colors[index % colors.len()].into());
            // Reference nodes contain the cache; literals are their own cache.
            let read = |n: Option<&Node>, numeric| {
                n.and_then(|n| n.children.first())
                    .map(|n| {
                        if n.name.ends_with("Ref") {
                            cached(n, numeric)
                        } else {
                            cached(
                                &Node {
                                    children: vec![n.clone()],
                                    ..Default::default()
                                },
                                numeric,
                            )
                        }
                    })
                    .unwrap_or_default()
            };
            series.push(Series {
                name,
                name_ref,
                categories_ref: cat_ref,
                values_ref,
                categories: read(cat, false),
                values: read(val, true).iter().map(Value::as_f64).collect(),
                color,
                width: ln
                    .and_then(|n| n.attrs.get("w"))
                    .and_then(|s| s.parse::<f64>().ok())
                    .map(|n| n / 9525.)
                    .unwrap_or(2.),
                markers: !is_bar
                    && s.child("marker")
                        .and_then(|n| n.first("symbol"))
                        .and_then(Node::val)
                        .is_some_and(|v| v != "none"),
                kind: if is_bar { "bar" } else { "line" }.into(),
                secondary: on_second,
                smooth: !is_bar && s.child("smooth").and_then(Node::val) == Some("1"),
                format: val
                    .and_then(|n| n.first("formatCode"))
                    .map(|n| n.text.clone())
                    .unwrap_or_default(),
            });
        }
    }
    if series.is_empty() {
        unsupported = Some("This chart has no supported data series.".into());
    }
    let title = root
        .first("chart")
        .and_then(|c| c.child("title"))
        .map(|t| {
            t.all("t")
                .iter()
                .map(|n| n.text.as_str())
                .collect::<String>()
        })
        .unwrap_or_default();
    let axis_num = |axis: Option<&Node>, name: &str| -> Option<f64> {
        axis.and_then(|n| n.first(name))
            .and_then(Node::val)
            .and_then(|n| n.parse().ok())
    };
    // An axis is labelled in its own format, or, when it says it follows
    // its source (or says nothing), in the format of the cells plotted
    // against it: a margin line's axis reads 10%, 20%, not 0, 0.
    let axis_format = |axis: Option<&Node>, second: bool| -> String {
        let stated = axis.and_then(|n| n.child("numFmt"));
        let code = stated
            .and_then(|n| n.attrs.get("formatCode"))
            .cloned()
            .unwrap_or_default();
        let follows_source = stated.and_then(|n| n.attrs.get("sourceLinked")).map(String::as_str) == Some("1");
        if !follows_source && !code.is_empty() && !code.eq_ignore_ascii_case("General") {
            return code;
        }
        series
            .iter()
            .find(|s| s.secondary == second && !s.format.is_empty() && !s.format.eq_ignore_ascii_case("General"))
            .map(|s| s.format.clone())
            .unwrap_or(code)
    };
    let (number_format, number_format2) = (axis_format(axis, false), axis_format(axis2, true));
    Ok(Chart {
        id: format!("{part}#{id}"),
        part: part.into(),
        sheet: sheet.into(),
        title,
        kind,
        unsupported,
        anchor,
        series,
        legend: root.first("legend").is_some(),
        legend_position: root
            .first("legendPos")
            .and_then(Node::val)
            .unwrap_or("b")
            .into(),
        y_min: axis_num(axis, "min"),
        y_max: axis_num(axis, "max"),
        number_format,
        blanks: root
            .first("dispBlanksAs")
            .and_then(Node::val)
            .unwrap_or("gap")
            .into(),
        bar_dir,
        bar_grouping,
        bar_gap_width: bars.first().and_then(|p| p.child("gapWidth")).and_then(Node::val).and_then(|v| v.parse().ok()),
        bar_overlap: bars.first().and_then(|p| p.child("overlap")).and_then(Node::val).and_then(|v| v.parse().ok()),
        y2_min: axis_num(axis2, "min"),
        y2_max: axis_num(axis2, "max"),
        number_format2,
    })
}

/// Fills the series of `charts` from the live cells. `read(range, text)`
/// gives the cells of a range, row by row: as the text the sheet shows
/// when `text`, as values otherwise; `None` when the range's sheet is not
/// there. A chart whose data cannot be read says so instead of drawing
/// stale numbers.
pub fn fill_live(charts: &mut [Chart], read: &dyn Fn(&Range, bool) -> Option<Vec<Value>>) {
    let cells = |formula: &str, text: bool| range(formula).and_then(|r| read(&r, text));
    for chart in charts.iter_mut() {
        for series in &mut chart.series {
            if let Some(f) = &series.name_ref {
                if let Some(v) = cells(f, true) {
                    series.name = v.first().and_then(Value::as_str).unwrap_or("").into();
                }
            }
            if let Some(f) = &series.categories_ref {
                match cells(f, true) {
                    Some(v) => series.categories = v,
                    None => chart.unsupported = Some("The category reference cannot be resolved.".into()),
                }
            }
            if let Some(f) = &series.values_ref {
                match cells(f, false) {
                    Some(v) => series.values = v.iter().map(Value::as_f64).collect(),
                    None => chart.unsupported = Some("The value reference cannot be resolved.".into()),
                }
            }
        }
    }
}

/// The Office palette a series takes its colour from when it states none.
pub const PALETTE: [&str; 6] = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];

/// A chart to create, as a caller asks for it: where it sits, what kind it
/// is and the ranges it plots. `build` turns it into a `Chart`.
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartSpec {
    /// The new chart's id (it is what takes the chart away again).
    pub id: String,
    /// The sheet the chart sits on. Ranges without a sheet are on it too.
    pub sheet: String,
    /// Its top-left cell, 0-based.
    pub row: u32,
    pub col: u32,
    /// Size in pixels; 480 x 288 (Excel's own) when left out.
    #[serde(default)]
    pub width: Option<f64>,
    #[serde(default)]
    pub height: Option<f64>,
    #[serde(default)]
    pub title: String,
    /// "column", "bar" (lying bars) or "line".
    pub kind: String,
    /// Bars: "clustered" (when left out), "stacked" or "percentStacked".
    #[serde(default)]
    pub grouping: String,
    /// The category labels: one row or one column of cells.
    #[serde(default)]
    pub categories: Option<String>,
    pub series: Vec<SeriesSpec>,
}
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SeriesSpec {
    /// The series' name as text, or (`name_ref`) the cell that holds it.
    #[serde(default)]
    pub name: Option<String>,
    #[serde(default)]
    pub name_ref: Option<String>,
    /// The values: one row or one column of cells.
    pub values: String,
    /// "line" or "column", when it differs from the chart's kind (a margin
    /// line over revenue columns).
    #[serde(default)]
    pub kind: Option<String>,
    /// Plotted against a second value axis on the right. Lines only.
    #[serde(default)]
    pub secondary: bool,
    /// "#RRGGBB"; the Office palette by position when left out.
    #[serde(default)]
    pub color: Option<String>,
}

/// Gives a sheet's name as the workbook spells it; `None` when the workbook
/// has no such sheet.
pub type SheetNamed<'a> = &'a dyn Fn(&str) -> Option<String>;

/// A reference as a chart keeps it: `'P&L'!$B$5:$M$5`. `text` may leave the
/// sheet out, which then is `sheet`.
fn chart_reference(text: &str, sheet: &str, what: &str, sheet_named: SheetNamed) -> Result<(String, Range), String> {
    let text = text.trim().trim_start_matches('=');
    let qualified = if text.contains('!') {
        text.to_string()
    } else {
        format!("{}!{text}", refs::format_sheet_qualifier(sheet))
    };
    let mut r = range(&qualified)
        .ok_or_else(|| format!("{what} \"{text}\" is not one row or one column of cells, such as B5:M5 or 'P&L'!B5:M5"))?;
    r.sheet = sheet_named(&r.sheet)
        .ok_or_else(|| format!("{what}: there is no sheet \"{}\" in this workbook (a chart plots cells of its own workbook)", r.sheet))?;
    let cell = |row: u32, col: u32| format!("${}${}", refs::col_index_to_letters(col), row + 1);
    let cells = if r.len() == 1 {
        cell(r.r0, r.c0)
    } else {
        format!("{}:{}", cell(r.r0, r.c0), cell(r.r1, r.c1))
    };
    Ok((format!("{}!{cells}", refs::format_sheet_qualifier(&r.sheet)), r))
}

/// The chart a spec asks for, or why it cannot be made. Its series are in
/// the order the file keeps them (bars, lines, lines on the second axis),
/// and hold no data yet: `fill_live` reads it. Sheet names come out as
/// `sheet_named` spells them.
pub fn build(spec: &ChartSpec, sheet_named: SheetNamed) -> Result<Chart, String> {
    let lying = match spec.kind.as_str() {
        "column" | "line" => false,
        "bar" => true,
        other => return Err(format!("type \"{other}\" is not one of column, bar, line")),
    };
    let grouping = match spec.grouping.as_str() {
        "" | "clustered" => "clustered",
        "stacked" => "stacked",
        "percentStacked" => "percentStacked",
        other => return Err(format!("grouping \"{other}\" is not one of clustered, stacked, percentStacked")),
    };
    if spec.series.is_empty() || spec.series.len() > 24 {
        return Err(format!("a chart takes 1 to 24 series, not {}", spec.series.len()));
    }
    if spec.sheet.trim().is_empty() {
        return Err("the chart needs the sheet it sits on".into());
    }
    let sheet = sheet_named(spec.sheet.trim()).ok_or_else(|| format!("there is no sheet \"{}\" to put the chart on", spec.sheet.trim()))?;
    let categories = spec
        .categories
        .as_deref()
        .filter(|c| !c.trim().is_empty())
        .map(|c| chart_reference(c, &sheet, "categories", sheet_named))
        .transpose()?;
    let mut series = Vec::with_capacity(spec.series.len());
    for (i, s) in spec.series.iter().enumerate() {
        let n = i + 1;
        let is_bar = match s.kind.as_deref().unwrap_or(&spec.kind) {
            "line" => false,
            "column" | "bar" => true,
            other => return Err(format!("series {n}: type \"{other}\" is not line or column")),
        };
        if lying && (!is_bar || s.secondary) {
            return Err("lying bars cannot be combined with a line or a second axis: use type column".into());
        }
        if s.secondary && is_bar {
            return Err(format!("series {n}: only a line can go on the second axis"));
        }
        let (values_ref, values) = chart_reference(&s.values, &sheet, &format!("series {n}: values"), sheet_named)?;
        if let Some((_, c)) = &categories {
            if c.len() != values.len() {
                return Err(format!(
                    "series {n}: values has {} cells and categories has {}: they must be the same length",
                    values.len(),
                    c.len()
                ));
            }
        }
        let name_ref = s
            .name_ref
            .as_deref()
            .filter(|c| !c.trim().is_empty())
            .map(|c| chart_reference(c, &sheet, &format!("series {n}: name cell"), sheet_named))
            .transpose()?;
        if name_ref.as_ref().is_some_and(|(_, r)| r.len() != 1) {
            return Err(format!("series {n}: the name cell must be one cell"));
        }
        let color = match s.color.as_deref().map(str::trim).filter(|c| !c.is_empty()) {
            Some(c) => {
                let hex = c.trim_start_matches('#');
                if hex.len() != 6 || !hex.chars().all(|ch| ch.is_ascii_hexdigit()) {
                    return Err(format!("series {n}: colour \"{c}\" is not #RRGGBB"));
                }
                Some(format!("#{}", hex.to_ascii_uppercase()))
            }
            None => None,
        };
        series.push((
            color,
            Series {
                name: s.name.clone().filter(|t| !t.trim().is_empty()).unwrap_or_else(|| format!("Series {n}")),
                name_ref: name_ref.map(|(f, _)| f),
                categories_ref: categories.as_ref().map(|(f, _)| f.clone()),
                values_ref: Some(values_ref),
                categories: vec![],
                values: vec![],
                color: String::new(),
                width: 3.,
                markers: false,
                kind: if is_bar { "bar" } else { "line" }.into(),
                secondary: s.secondary,
                smooth: false,
                format: String::new(),
            },
        ));
    }
    if series.iter().all(|(_, s)| s.secondary) {
        return Err("at least one series must be on the main axis".into());
    }
    // The file keeps bars, then lines, then the second axis' lines.
    series.sort_by_key(|(_, s)| (s.kind != "bar", s.secondary));
    let series: Vec<Series> = series
        .into_iter()
        .enumerate()
        .map(|(i, (color, mut s))| {
            s.color = color.unwrap_or_else(|| PALETTE[i % PALETTE.len()].into());
            s
        })
        .collect();
    let bars = series.iter().any(|s| s.kind == "bar");
    if !bars && grouping != "clustered" {
        return Err("stacked lines are not drawn: use type column for a stacked chart".into());
    }
    Ok(Chart {
        id: spec.id.clone(),
        part: String::new(),
        sheet,
        title: spec.title.trim().to_string(),
        kind: if bars { "barChart" } else { "lineChart" }.into(),
        unsupported: None,
        anchor: Anchor {
            from: Marker { row: spec.row, col: spec.col, dx: 0., dy: 0. },
            to: None,
            width: spec.width.filter(|w| w.is_finite()).unwrap_or(480.).clamp(160., 2400.),
            height: spec.height.filter(|h| h.is_finite()).unwrap_or(288.).clamp(120., 2400.),
            absolute: false,
        },
        legend: series.len() > 1,
        legend_position: "b".into(),
        blanks: "gap".into(),
        bar_dir: if !bars { "" } else if lying { "bar" } else { "col" }.into(),
        bar_grouping: if bars { grouping } else { "" }.into(),
        bar_gap_width: bars.then_some(150.),
        bar_overlap: (bars && grouping != "clustered").then_some(100.),
        series,
        ..Default::default()
    })
}

/// A chart made in this session takes its axis formats from the cells it
/// plots, as Excel does for a new chart: `format_of(range)` is the number
/// format of a range's first cell. Charts from the file keep what the file
/// says.
pub fn link_formats(charts: &mut [Chart], format_of: &dyn Fn(&Range) -> Option<String>) {
    for chart in charts.iter_mut().filter(|c| c.part.is_empty()) {
        for series in &mut chart.series {
            series.format = series
                .values_ref
                .as_deref()
                .and_then(range)
                .and_then(|r| format_of(&r))
                .unwrap_or_default();
        }
        let linked = |second: bool| {
            chart
                .series
                .iter()
                .find(|s| s.secondary == second && !s.format.is_empty() && !s.format.eq_ignore_ascii_case("General"))
                .map(|s| s.format.clone())
                .unwrap_or_default()
        };
        let (main, second) = (linked(false), linked(true));
        chart.number_format = if chart.bar_grouping == "percentStacked" { "0%".into() } else { main };
        chart.number_format2 = second;
    }
}

/// Rows or columns moved on `sheet`: the ranges the charts read follow,
/// and so do the charts that sit on that sheet.
pub fn shift(charts: &mut [Chart], sheet: &str, shift: &refs::RowColShift) -> Result<(), PatchError> {
    for chart in charts.iter_mut() {
        for series in &mut chart.series {
            for f in [&mut series.name_ref, &mut series.categories_ref, &mut series.values_ref]
                .into_iter()
                .flatten()
            {
                *f = refs::adjust_formula(f, shift, sheet, false)?;
            }
        }
        if chart.sheet.eq_ignore_ascii_case(sheet) && !chart.anchor.absolute {
            for marker in std::iter::once(&mut chart.anchor.from).chain(chart.anchor.to.iter_mut()) {
                let value = if shift.rows { &mut marker.row } else { &mut marker.col };
                *value = shift.clamp_index(*value);
            }
        }
    }
    Ok(())
}

/// A sheet was renamed: charts on it and ranges that name it follow.
pub fn rename_sheet(charts: &mut [Chart], old: &str, new: &str) -> Result<(), PatchError> {
    for chart in charts.iter_mut() {
        if chart.sheet == old {
            chart.sheet = new.into();
        }
        for series in &mut chart.series {
            for f in [&mut series.name_ref, &mut series.categories_ref, &mut series.values_ref]
                .into_iter()
                .flatten()
            {
                *f = refs::rename_sheet_in_formula(f, old, new)?;
            }
        }
    }
    Ok(())
}

/// A sheet was deleted: the charts on it go with it.
pub fn remove_sheet(charts: &mut Vec<Chart>, name: &str) {
    charts.retain(|c| c.sheet != name);
}

pub fn read(bytes: &[u8]) -> Result<Vec<Chart>, PatchError> {
    catalog(&mut PartStore::open(bytes)?)
}
pub fn catalog(store: &mut PartStore) -> Result<Vec<Chart>, PatchError> {
    let mut out = vec![];
    for (sheet, path) in sheets(store)? {
        for (_, kind, drawing) in relationships(store, &path)? {
            if kind != "drawing" {
                continue;
            }
            let rels = relationships(store, &drawing)?;
            let root = tree(&store.read(&drawing)?)?;
            for a in root
                .children
                .iter()
                .flat_map(|n| n.children.iter())
                .filter(|n| n.name.ends_with("Anchor"))
            {
                let Some(link) = a.first("chart") else {
                    continue;
                };
                let Some((_, _, part)) = rels
                    .iter()
                    .find(|(id, t, _)| t == "chart" && Some(id) == link.attrs.get("id"))
                else {
                    continue;
                };
                let ext = a.child("ext");
                let pos = a.child("pos");
                let mut from = marker(a.child("from"));
                if let Some(p) = pos {
                    from.dx = p
                        .attrs
                        .get("x")
                        .and_then(|s| s.parse::<f64>().ok())
                        .unwrap_or(0.)
                        / 9525.;
                    from.dy = p
                        .attrs
                        .get("y")
                        .and_then(|s| s.parse::<f64>().ok())
                        .unwrap_or(0.)
                        / 9525.;
                }
                let anchor = Anchor {
                    from,
                    to: a.child("to").map(|n| marker(Some(n))),
                    absolute: pos.is_some(),
                    width: ext
                        .and_then(|n| n.attrs.get("cx"))
                        .and_then(|s| s.parse::<f64>().ok())
                        .unwrap_or(4_762_500.)
                        / 9525.,
                    height: ext
                        .and_then(|n| n.attrs.get("cy"))
                        .and_then(|s| s.parse::<f64>().ok())
                        .unwrap_or(2_857_500.)
                        / 9525.,
                };
                let id = out.len();
                out.push(
                    chart(store, &sheet, part, anchor.clone(), id).unwrap_or_else(|_| Chart {
                        id: format!("{part}#{id}"),
                        part: part.clone(),
                        sheet: sheet.clone(),
                        title: "Excel chart".into(),
                        kind: "unknown".into(),
                        unsupported: Some("This chart's definition could not be read.".into()),
                        anchor,
                        series: vec![],
                        legend_position: "b".into(),
                        blanks: "gap".into(),
                        ..Default::default()
                    }),
                );
            }
        }
    }
    Ok(out)
}

/// Locate chart parts using content types, not their conventional directory.
pub fn chart_parts(store: &mut PartStore) -> Result<Vec<String>, PatchError> {
    let doc = tree(&store.read("[Content_Types].xml")?)?;
    Ok(doc
        .all("Override")
        .into_iter()
        .filter(|n| {
            n.attrs
                .get("ContentType")
                .is_some_and(|t| t.ends_with("drawingml.chart+xml"))
        })
        .filter_map(|n| {
            n.attrs
                .get("PartName")
                .map(|s| s.trim_start_matches('/').to_owned())
        })
        .collect())
}

pub(crate) fn escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn source_range(node: &Node, paths: &HashMap<String, String>) -> Option<Range> {
    let mut r = range(&node.child("f")?.text)?;
    r.sheet = paths
        .keys()
        .find(|name| name.eq_ignore_ascii_case(&r.sheet))?
        .clone();
    Some(r)
}

/// Refresh supported direct references from the already-patched worksheet caches.
/// Unknown references remain unchanged. Never regenerate the chart or its styles.
pub fn refresh_caches(
    store: &mut PartStore,
    affected: impl Fn(&Range) -> bool,
) -> Result<(), PatchError> {
    let parts = chart_parts(store)?;
    if parts.is_empty() {
        return Ok(());
    }
    let paths: HashMap<_, _> = sheets(store)?.into_iter().collect();
    let mut docs = vec![];
    let mut needed: HashMap<String, HashSet<(u32, u32)>> = HashMap::new();
    for part in parts {
        let xml = store.read(&part)?;
        // An unsupported/malformed chart must not prevent an unrelated cell save.
        let Ok(doc) = tree(&xml) else { continue };
        let references: Vec<_> = doc
            .all("numRef")
            .into_iter()
            .chain(doc.all("strRef"))
            .collect();
        if !references.iter().any(|n| {
            n.child("f")
                .and_then(|n| range(&n.text))
                .is_some_and(|r| affected(&r))
        }) {
            continue;
        }
        // Once a chart's data changed, fill all its resolvable caches (some
        // producers omit category/name caches even when their cells exist).
        for node in references {
            if let Some(r) = source_range(node, &paths) {
                if paths.contains_key(&r.sheet) {
                    needed.entry(r.sheet.clone()).or_default().extend(r.cells());
                }
            }
        }
        docs.push((part, xml, doc));
    }
    if needed.is_empty() {
        return Ok(());
    }
    let mut sst = vec![];
    if store.exists("xl/sharedStrings.xml") {
        let doc = tree(&store.read("xl/sharedStrings.xml")?)?;
        sst = doc
            .all("si")
            .iter()
            .map(|n| {
                n.all("t")
                    .iter()
                    .map(|t| t.text.as_str())
                    .collect::<String>()
            })
            .collect();
    }
    let mut values = HashMap::new();
    for (sheet, wanted) in needed {
        let xml = store.read(&paths[&sheet])?;
        let xml = xml.strip_prefix(&[0xef, 0xbb, 0xbf]).unwrap_or(&xml);
        let mut r = reader(xml);
        let mut buf = vec![];
        let mut start = None;
        let mut key = None;
        loop {
            let before = r.buffer_position() as usize;
            match r
                .read_event_into(&mut buf)
                .map_err(|e| PatchError::Xml(e.to_string()))?
            {
                Event::Start(e) if local_name(e.name().as_ref()) == b"c" => {
                    key = attr_value(&e, b"r")
                        .and_then(|s| refs::parse_cell_ref(&s))
                        .filter(|rc| wanted.contains(rc));
                    start = key.map(|_| before);
                }
                Event::End(e) if local_name(e.name().as_ref()) == b"c" => {
                    if let (Some(start), Some((row, col))) = (start.take(), key.take()) {
                        let node = tree(&xml[start..r.buffer_position() as usize])?;
                        let cell = &node.children[0];
                        let raw = cell.child("v").map(|v| v.text.as_str()).unwrap_or("");
                        let v = match cell.attrs.get("t").map(String::as_str) {
                            Some("s") => raw
                                .parse::<usize>()
                                .ok()
                                .and_then(|i| sst.get(i))
                                .cloned()
                                .map(Value::String)
                                .unwrap_or(Value::Null),
                            Some("inlineStr") => Value::String(
                                cell.all("t").iter().map(|n| n.text.as_str()).collect(),
                            ),
                            Some("str") => Value::String(raw.into()),
                            Some("e") => Value::Null,
                            _ => raw
                                .parse::<f64>()
                                .ok()
                                .filter(|n| n.is_finite())
                                .map(Value::from)
                                .unwrap_or(Value::Null),
                        };
                        values.insert((sheet.clone(), row, col), v);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
    }
    for (part, mut xml, doc) in docs {
        let mut replacements = vec![];
        for node in doc.all("numRef").into_iter().chain(doc.all("strRef")) {
            let Some(r) = source_range(node, &paths) else {
                continue;
            };
            let numeric = node.name == "numRef";
            let new: Vec<_> = r
                .cells()
                .map(|(row, col)| {
                    let v = values
                        .get(&(r.sheet.clone(), row, col))
                        .cloned()
                        .unwrap_or(Value::Null);
                    if numeric {
                        if v.is_number() {
                            v
                        } else {
                            Value::Null
                        }
                    } else {
                        match v {
                            Value::Null => Value::Null,
                            Value::String(_) => v,
                            Value::Number(n) => Value::String(n.as_f64().unwrap_or(0.).to_string()),
                            _ => Value::String(v.to_string()),
                        }
                    }
                })
                .collect();
            if cached(node, numeric) == new {
                continue;
            }
            let tag = if numeric { "numCache" } else { "strCache" };
            let cache = node.child(tag);
            let prefix = cache
                .unwrap_or(node)
                .qualified
                .rsplit_once(':')
                .map(|(p, _)| format!("{p}:"))
                .unwrap_or_default();
            let q = |n: &str| format!("{prefix}{n}");
            let mut text = if let Some(cache) = cache {
                let open = std::str::from_utf8(&xml[cache.start..cache.open_end])
                    .map_err(|e| PatchError::Xml(e.to_string()))?;
                if let Some(open) = open.strip_suffix("/>") {
                    format!("{open}>")
                } else {
                    open.into()
                }
            } else {
                format!("<{}>", q(tag))
            };
            if numeric {
                if let Some(fmt) = cache.and_then(|n| n.child("formatCode")) {
                    text.push_str(
                        std::str::from_utf8(&xml[fmt.start..fmt.end])
                            .map_err(|e| PatchError::Xml(e.to_string()))?,
                    );
                } else {
                    text.push_str(&format!(
                        "<{}>General</{}>",
                        q("formatCode"),
                        q("formatCode")
                    ));
                }
            }
            text.push_str(&format!("<{} val=\"{}\"/>", q("ptCount"), new.len()));
            for (i, v) in new.iter().enumerate().filter(|(_, v)| !v.is_null()) {
                let s = v
                    .as_str()
                    .map(str::to_owned)
                    .unwrap_or_else(|| v.to_string());
                text.push_str(&format!(
                    "<{} idx=\"{i}\"><{}>{}</{}></{}>",
                    q("pt"),
                    q("v"),
                    escape(&s),
                    q("v"),
                    q("pt")
                ));
            }
            // Preserve extension metadata inside a cache if present.
            if let Some(ext) = cache.and_then(|n| n.child("extLst")) {
                text.push_str(
                    std::str::from_utf8(&xml[ext.start..ext.end])
                        .map_err(|e| PatchError::Xml(e.to_string()))?,
                );
            }
            text.push_str(&format!("</{}>", q(tag)));
            replacements.push((
                cache.map(|n| n.start).unwrap_or(node.close),
                cache.map(|n| n.end).unwrap_or(node.close),
                text.into_bytes(),
            ));
        }
        if !replacements.is_empty() {
            replacements.sort_by_key(|(start, _, _)| std::cmp::Reverse(*start));
            for (start, end, bytes) in replacements {
                xml.splice(start..end, bytes);
            }
            store.write(&part, xml);
        }
    }
    Ok(())
}
