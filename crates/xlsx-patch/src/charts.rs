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
struct Node {
    name: String,
    qualified: String,
    attrs: HashMap<String, String>,
    text: String,
    children: Vec<Node>,
    start: usize,
    end: usize,
    open_end: usize,
    close: usize,
}
impl Node {
    fn child(&self, name: &str) -> Option<&Node> {
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
    fn all(&self, name: &str) -> Vec<&Node> {
        let mut v = vec![];
        self.descendants(name, &mut v);
        v
    }
    fn first(&self, name: &str) -> Option<&Node> {
        if self.name == name {
            return Some(self);
        }
        self.children.iter().find_map(|c| c.first(name))
    }
    fn val(&self) -> Option<&str> {
        self.attrs.get("val").map(String::as_str)
    }
    fn number(&self) -> f64 {
        self.text.parse().unwrap_or(0.)
    }
}
fn tree(xml: &[u8]) -> Result<Node, PatchError> {
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
fn rels_path(source: &str) -> String {
    let (dir, file) = source
        .rsplit_once('/')
        .map(|(d, f)| (format!("{d}/"), f))
        .unwrap_or((String::new(), source));
    format!("{dir}_rels/{file}.rels")
}
fn relationships(
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
fn sheets(store: &mut PartStore) -> Result<Vec<(String, String)>, PatchError> {
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
}
#[derive(Clone, Debug, Serialize, Deserialize)]
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
    let mut unsupported = None;
    if plots.len() != 1 || kind != "lineChart" {
        unsupported = Some("This chart type is preserved but not displayed yet.".into());
    }
    if let Some(line) = plots.first() {
        if line
            .child("grouping")
            .and_then(Node::val)
            .is_some_and(|v| v != "standard")
            || line.all("smooth").iter().any(|n| n.val() == Some("1"))
        {
            unsupported = Some("Stacked or smoothed lines are not displayed yet.".into());
        }
    }
    if !plot.all("dateAx").is_empty()
        || !plot.all("logBase").is_empty()
        || plot.all("valAx").len() > 1
    {
        unsupported = Some("Date, logarithmic or secondary axes are not displayed yet.".into());
    }
    let colors = [
        "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47",
    ];
    let mut series = vec![];
    if let Some(p) = plots.first() {
        for (i, s) in p.children.iter().filter(|n| n.name == "ser").enumerate() {
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
            let color = ln
                .and_then(|n| n.first("srgbClr"))
                .and_then(Node::val)
                .filter(|s| s.len() == 6 && s.chars().all(|c| c.is_ascii_hexdigit()))
                .map(|s| format!("#{s}"))
                .unwrap_or_else(|| colors[i % colors.len()].into());
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
                markers: s
                    .child("marker")
                    .and_then(|n| n.first("symbol"))
                    .and_then(Node::val)
                    .is_some_and(|v| v != "none"),
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
    let axis = plot.child("valAx");
    let axis_num = |name: &str| {
        axis.and_then(|n| n.first(name))
            .and_then(Node::val)
            .and_then(|n| n.parse().ok())
    };
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
        y_min: axis_num("min"),
        y_max: axis_num("max"),
        number_format: axis
            .and_then(|n| n.child("numFmt"))
            .and_then(|n| n.attrs.get("formatCode"))
            .cloned()
            .unwrap_or_default(),
        blanks: root
            .first("dispBlanksAs")
            .and_then(Node::val)
            .unwrap_or("gap")
            .into(),
    })
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
                        legend: false,
                        legend_position: "b".into(),
                        y_min: None,
                        y_max: None,
                        number_format: String::new(),
                        blanks: "gap".into(),
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

fn escape(s: &str) -> String {
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
