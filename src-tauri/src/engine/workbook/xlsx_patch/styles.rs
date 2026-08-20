//! Append-only interning into xl/styles.xml.
//!
//! Style patches are PARTIAL ("make this bold") so every derived style is
//! composed from the cell's current xf: clone the base font/border/xf,
//! apply the overrides, then find-or-append each component. Existing
//! entries are NEVER modified or renumbered — every s= index already used
//! by ten thousand untouched cells stays valid.

use std::collections::HashMap;

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::Writer;

use super::patch::{BorderSide, StylePatch};
use super::workbook_xml::{local_name, reader, xml_err};
use super::PatchError;

// ---------------------------------------------------------------------------
// Generic element tree (styles.xml is small; a full parse is fine)
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq)]
pub struct Elem {
    pub tag: String,
    pub attrs: Vec<(String, String)>,
    pub children: Vec<Elem>,
    /// Text content (formatCode strings never appear as text — numFmt uses
    /// attributes — but keep this for safety).
    pub text: String,
}

impl Elem {
    fn new(tag: &str) -> Self {
        Elem {
            tag: tag.into(),
            attrs: Vec::new(),
            children: Vec::new(),
            text: String::new(),
        }
    }

    fn attr(&self, key: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(k, _)| k == key)
            .map(|(_, v)| v.as_str())
    }

    fn set_attr(&mut self, key: &str, value: String) {
        if let Some(slot) = self.attrs.iter_mut().find(|(k, _)| k == key) {
            slot.1 = value;
        } else {
            self.attrs.push((key.into(), value));
        }
    }

    fn remove_attr(&mut self, key: &str) {
        self.attrs.retain(|(k, _)| k != key);
    }

    fn child(&self, tag: &str) -> Option<&Elem> {
        self.children.iter().find(|c| c.tag == tag)
    }

    fn remove_children(&mut self, tag: &str) {
        self.children.retain(|c| c.tag != tag);
    }

    fn upsert_child(&mut self, elem: Elem) {
        if let Some(slot) = self.children.iter_mut().find(|c| c.tag == elem.tag) {
            *slot = elem;
        } else {
            self.children.push(elem);
        }
    }

    /// Canonical form used for dedup comparisons (attribute ORDER matters —
    /// a rare false-negative just appends a redundant entry, never corrupts).
    fn canonical(&self) -> String {
        let mut out = String::new();
        self.write_canonical(&mut out);
        out
    }

    fn write_canonical(&self, out: &mut String) {
        out.push('<');
        out.push_str(&self.tag);
        let mut attrs = self.attrs.clone();
        attrs.sort();
        for (k, v) in &attrs {
            out.push_str(&format!(" {k}=\"{v}\""));
        }
        out.push('>');
        out.push_str(&self.text);
        for c in &self.children {
            c.write_canonical(out);
        }
        out.push_str("</");
        out.push_str(&self.tag);
        out.push('>');
    }

    fn write_xml(&self, w: &mut Writer<Vec<u8>>) -> Result<(), PatchError> {
        let mut start = BytesStart::new(self.tag.clone());
        for (k, v) in &self.attrs {
            start.push_attribute((k.as_str(), v.as_str()));
        }
        if self.children.is_empty() && self.text.is_empty() {
            w.write_event(Event::Empty(start)).map_err(xml_err)?;
        } else {
            w.write_event(Event::Start(start)).map_err(xml_err)?;
            if !self.text.is_empty() {
                w.write_event(Event::Text(quick_xml::events::BytesText::new(&self.text)))
                    .map_err(xml_err)?;
            }
            for c in &self.children {
                c.write_xml(w)?;
            }
            w.write_event(Event::End(BytesEnd::new(self.tag.clone())))
                .map_err(xml_err)?;
        }
        Ok(())
    }
}

fn parse_document(xml: &[u8]) -> Result<Elem, PatchError> {
    let mut r = reader(xml);
    let mut buf = Vec::new();
    let mut stack: Vec<Elem> = vec![Elem::new("#root")];
    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("styles parse: {e}")))?;
        match ev {
            Event::Eof => break,
            Event::Start(e) => {
                stack.push(elem_from_start(&e)?);
            }
            Event::Empty(e) => {
                let el = elem_from_start(&e)?;
                stack.last_mut().unwrap().children.push(el);
            }
            Event::End(_) => {
                let el = stack.pop().ok_or_else(|| {
                    PatchError::Xml("styles parse: unbalanced end tag".into())
                })?;
                stack
                    .last_mut()
                    .ok_or_else(|| PatchError::Xml("styles parse: underflow".into()))?
                    .children
                    .push(el);
            }
            Event::Text(t) => {
                let txt = t
                    .unescape()
                    .map_err(|e| PatchError::Xml(format!("styles text: {e}")))?;
                if !txt.trim().is_empty() {
                    stack.last_mut().unwrap().text.push_str(&txt);
                }
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(stack.pop().unwrap())
}

fn elem_from_start(e: &BytesStart) -> Result<Elem, PatchError> {
    let mut el = Elem::new(&String::from_utf8_lossy(local_name(e.name().as_ref())));
    for a in e.attributes().with_checks(false).flatten() {
        let k = String::from_utf8_lossy(a.key.as_ref()).into_owned();
        let v = a
            .unescape_value()
            .map_err(|e| PatchError::Xml(format!("styles attr: {e}")))?
            .into_owned();
        el.attrs.push((k, v));
    }
    Ok(el)
}

// ---------------------------------------------------------------------------
// Editor
// ---------------------------------------------------------------------------

pub struct StylesEditor {
    original: Vec<u8>,
    num_fmts: Vec<Elem>,
    fonts: Vec<Elem>,
    fills: Vec<Elem>,
    borders: Vec<Elem>,
    cell_xfs: Vec<Elem>,
    /// Count of pre-existing entries per list (everything past that index in
    /// the vecs above is appended by this session).
    base_counts: HashMap<&'static str, usize>,
    dirty: bool,
}

impl StylesEditor {
    pub fn parse(xml: Vec<u8>) -> Result<Self, PatchError> {
        let doc = parse_document(&xml)?;
        let root = doc
            .children
            .iter()
            .find(|c| c.tag == "styleSheet")
            .ok_or_else(|| PatchError::Xml("styles.xml has no styleSheet".into()))?;
        let list = |name: &str| -> Vec<Elem> {
            root.child(name)
                .map(|l| l.children.clone())
                .unwrap_or_default()
        };
        let num_fmts = list("numFmts");
        let fonts = list("fonts");
        let fills = list("fills");
        let borders = list("borders");
        let cell_xfs = list("cellXfs");
        let mut base_counts = HashMap::new();
        base_counts.insert("numFmts", num_fmts.len());
        base_counts.insert("fonts", fonts.len());
        base_counts.insert("fills", fills.len());
        base_counts.insert("borders", borders.len());
        base_counts.insert("cellXfs", cell_xfs.len());
        Ok(StylesEditor {
            original: xml,
            num_fmts,
            fonts,
            fills,
            borders,
            cell_xfs,
            base_counts,
            dirty: false,
        })
    }

    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Compose a new xf from `base_xf` + the patch and return its index.
    pub fn intern(&mut self, base_xf: u32, p: &StylePatch) -> Result<u32, PatchError> {
        let base = self
            .cell_xfs
            .get(base_xf as usize)
            .cloned()
            .unwrap_or_else(|| {
                let mut xf = Elem::new("xf");
                for k in ["numFmtId", "fontId", "fillId", "borderId", "xfId"] {
                    xf.set_attr(k, "0".into());
                }
                xf
            });
        let mut xf = base;

        // --- font ---
        let font_touched = p.bold.is_some()
            || p.italic.is_some()
            || p.underline.is_some()
            || p.strike.is_some()
            || p.font_color.is_some()
            || p.font_size.is_some()
            || p.font_family.is_some();
        if font_touched {
            let base_font_id: usize = xf
                .attr("fontId")
                .and_then(|v| v.parse().ok())
                .unwrap_or(0);
            let mut font = self
                .fonts
                .get(base_font_id)
                .cloned()
                .unwrap_or_else(|| Elem::new("font"));
            for (flag, tag) in [
                (p.bold, "b"),
                (p.italic, "i"),
                (p.underline, "u"),
                (p.strike, "strike"),
            ] {
                if let Some(on) = flag {
                    font.remove_children(tag);
                    if on {
                        font.children.insert(0, Elem::new(tag));
                    }
                }
            }
            if let Some(color) = &p.font_color {
                let argb = css_to_argb(color)?;
                font.remove_children("color");
                let mut c = Elem::new("color");
                c.set_attr("rgb", argb);
                font.children.push(c);
            }
            if let Some(sz) = p.font_size {
                font.remove_children("sz");
                let mut s = Elem::new("sz");
                s.set_attr("val", trim_num(sz));
                font.children.push(s);
            }
            if let Some(name) = &p.font_family {
                font.remove_children("name");
                // A theme scheme would override the explicit name.
                font.remove_children("scheme");
                let mut n = Elem::new("name");
                n.set_attr("val", name.clone());
                font.children.push(n);
            }
            let id = find_or_append(&mut self.fonts, font);
            xf.set_attr("fontId", id.to_string());
            xf.set_attr("applyFont", "1".into());
            self.dirty = true;
        }

        // --- fill ---
        if let Some(bg) = &p.background_color {
            let fill_id = match bg {
                Some(color) => {
                    let argb = css_to_argb(color)?;
                    let mut pf = Elem::new("patternFill");
                    pf.set_attr("patternType", "solid".into());
                    let mut fg = Elem::new("fgColor");
                    fg.set_attr("rgb", argb);
                    pf.children.push(fg);
                    let mut bgc = Elem::new("bgColor");
                    bgc.set_attr("indexed", "64".into());
                    pf.children.push(bgc);
                    let mut fill = Elem::new("fill");
                    fill.children.push(pf);
                    find_or_append(&mut self.fills, fill)
                }
                // Fill 0 is "none" by convention in every Excel-written file.
                None => 0,
            };
            xf.set_attr("fillId", fill_id.to_string());
            xf.set_attr("applyFill", "1".into());
            self.dirty = true;
        }

        // --- border ---
        if let Some(bp) = &p.borders {
            let base_border_id: usize = xf
                .attr("borderId")
                .and_then(|v| v.parse().ok())
                .unwrap_or(0);
            let mut border = self
                .borders
                .get(base_border_id)
                .cloned()
                .unwrap_or_else(|| Elem::new("border"));
            for (side, tag) in [
                (&bp.left, "left"),
                (&bp.right, "right"),
                (&bp.top, "top"),
                (&bp.bottom, "bottom"),
            ] {
                if let Some(new_side) = side {
                    border.remove_children(tag);
                    border.children.push(border_side_elem(tag, new_side.as_ref())?);
                }
            }
            // CT_Border wants sides in a fixed order.
            let order = ["left", "right", "top", "bottom", "diagonal", "vertical", "horizontal"];
            border.children.sort_by_key(|c| {
                order.iter().position(|t| *t == c.tag).unwrap_or(order.len())
            });
            let id = find_or_append(&mut self.borders, border);
            xf.set_attr("borderId", id.to_string());
            xf.set_attr("applyBorder", "1".into());
            self.dirty = true;
        }

        // --- number format ---
        if let Some(code) = &p.number_format {
            let id = self.num_fmt_id(code);
            xf.set_attr("numFmtId", id.to_string());
            xf.set_attr("applyNumberFormat", "1".into());
            self.dirty = true;
        }

        // --- alignment ---
        let align_touched = p.horizontal_align.is_some()
            || p.vertical_align.is_some()
            || p.wrap_text.is_some()
            || p.indent.is_some();
        if align_touched {
            let mut al = xf
                .child("alignment")
                .cloned()
                .unwrap_or_else(|| Elem::new("alignment"));
            if let Some(h) = &p.horizontal_align {
                al.set_attr("horizontal", h.clone());
            }
            if let Some(v) = &p.vertical_align {
                // Frontend says "middle"; OOXML says "center".
                al.set_attr("vertical", if v == "middle" { "center".into() } else { v.clone() });
            }
            if let Some(wt) = p.wrap_text {
                if wt {
                    al.set_attr("wrapText", "1".into());
                } else {
                    al.remove_attr("wrapText");
                }
            }
            if let Some(ind) = p.indent {
                if ind > 0 {
                    al.set_attr("indent", ind.to_string());
                } else {
                    al.remove_attr("indent");
                }
            }
            xf.upsert_child(al);
            xf.set_attr("applyAlignment", "1".into());
            self.dirty = true;
        }

        let idx = find_or_append(&mut self.cell_xfs, xf);
        Ok(idx as u32)
    }

    fn num_fmt_id(&mut self, code: &str) -> u32 {
        if let Some(id) = builtin_num_fmt(code) {
            return id;
        }
        for f in &self.num_fmts {
            if f.attr("formatCode") == Some(code) {
                if let Some(id) = f.attr("numFmtId").and_then(|v| v.parse().ok()) {
                    return id;
                }
            }
        }
        let next = self
            .num_fmts
            .iter()
            .filter_map(|f| f.attr("numFmtId").and_then(|v| v.parse::<u32>().ok()))
            .max()
            .map(|m| m + 1)
            .unwrap_or(164)
            .max(164);
        let mut nf = Elem::new("numFmt");
        nf.set_attr("numFmtId", next.to_string());
        nf.set_attr("formatCode", code.to_string());
        self.num_fmts.push(nf);
        self.dirty = true;
        next
    }

    /// Re-emit styles.xml: original events copied through; each list gets its
    /// count updated and appended items written before its end tag. A missing
    /// numFmts list is created before <fonts> when needed.
    pub fn serialize(&self) -> Result<Vec<u8>, PatchError> {
        let appended = |name: &'static str| -> &[Elem] {
            let items: &Vec<Elem> = match name {
                "numFmts" => &self.num_fmts,
                "fonts" => &self.fonts,
                "fills" => &self.fills,
                "borders" => &self.borders,
                "cellXfs" => &self.cell_xfs,
                _ => unreachable!(),
            };
            &items[self.base_counts[name]..]
        };
        let list_names: [&'static str; 5] = ["numFmts", "fonts", "fills", "borders", "cellXfs"];

        let need_numfmts_insert =
            self.base_counts["numFmts"] == 0 && !appended("numFmts").is_empty();

        let mut r = reader(&self.original);
        let mut w = Writer::new(Vec::new());
        let mut buf = Vec::new();
        let mut numfmts_inserted = !need_numfmts_insert;

        loop {
            let ev = r
                .read_event_into(&mut buf)
                .map_err(|e| PatchError::Xml(format!("styles serialize: {e}")))?;
            match ev {
                Event::Eof => break,
                Event::Start(e) => {
                    let name = local_name(e.name().as_ref()).to_vec();
                    if !numfmts_inserted && name == b"fonts" {
                        write_list(&mut w, "numFmts", appended("numFmts"))?;
                        numfmts_inserted = true;
                    }
                    if let Some(list) = list_names.iter().find(|n| n.as_bytes() == name) {
                        let total = match *list {
                            "numFmts" => self.num_fmts.len(),
                            "fonts" => self.fonts.len(),
                            "fills" => self.fills.len(),
                            "borders" => self.borders.len(),
                            "cellXfs" => self.cell_xfs.len(),
                            _ => unreachable!(),
                        };
                        w.write_event(Event::Start(super::workbook_xml::with_attr(
                            &e,
                            "count",
                            &total.to_string(),
                        )))
                        .map_err(xml_err)?;
                    } else {
                        w.write_event(Event::Start(e)).map_err(xml_err)?;
                    }
                }
                Event::End(e) => {
                    let name = local_name(e.name().as_ref()).to_vec();
                    if let Some(list) = list_names.iter().find(|n| n.as_bytes() == name) {
                        for item in appended(list) {
                            item.write_xml(&mut w)?;
                        }
                    }
                    w.write_event(Event::End(e)).map_err(xml_err)?;
                }
                other => {
                    w.write_event(other).map_err(xml_err)?;
                }
            }
            buf.clear();
        }
        Ok(w.into_inner())
    }
}

fn write_list(w: &mut Writer<Vec<u8>>, name: &str, items: &[Elem]) -> Result<(), PatchError> {
    let mut start = BytesStart::new(name.to_string());
    start.push_attribute(("count", items.len().to_string().as_str()));
    w.write_event(Event::Start(start)).map_err(xml_err)?;
    for item in items {
        item.write_xml(w)?;
    }
    w.write_event(Event::End(BytesEnd::new(name.to_string())))
        .map_err(xml_err)?;
    Ok(())
}

fn find_or_append(list: &mut Vec<Elem>, elem: Elem) -> usize {
    let key = elem.canonical();
    for (i, existing) in list.iter().enumerate() {
        if existing.canonical() == key {
            return i;
        }
    }
    list.push(elem);
    list.len() - 1
}

fn border_side_elem(tag: &str, side: Option<&BorderSide>) -> Result<Elem, PatchError> {
    let mut el = Elem::new(tag);
    if let Some(s) = side {
        el.set_attr("style", s.style.clone());
        let mut color = Elem::new("color");
        match &s.color {
            Some(c) => color.set_attr("rgb", css_to_argb(c)?),
            None => color.set_attr("auto", "1".into()),
        }
        el.children.push(color);
    }
    Ok(el)
}

/// "#1F4E79" → "FF1F4E79".
fn css_to_argb(css: &str) -> Result<String, PatchError> {
    let hex = css.trim().trim_start_matches('#');
    if hex.len() == 6 && hex.bytes().all(|b| b.is_ascii_hexdigit()) {
        Ok(format!("FF{}", hex.to_ascii_uppercase()))
    } else if hex.len() == 8 && hex.bytes().all(|b| b.is_ascii_hexdigit()) {
        Ok(hex.to_ascii_uppercase())
    } else {
        Err(PatchError::BadValue(format!("color {css:?}")))
    }
}

fn trim_num(v: f64) -> String {
    if v == v.trunc() {
        format!("{}", v as i64)
    } else {
        format!("{v}")
    }
}

fn builtin_num_fmt(code: &str) -> Option<u32> {
    Some(match code {
        "General" => 0,
        "0" => 1,
        "0.00" => 2,
        "#,##0" => 3,
        "#,##0.00" => 4,
        "0%" => 9,
        "0.00%" => 10,
        "0.00E+00" => 11,
        "# ?/?" => 12,
        "# ??/??" => 13,
        "mm-dd-yy" => 14,
        "d-mmm-yy" => 15,
        "d-mmm" => 16,
        "mmm-yy" => 17,
        "h:mm AM/PM" => 18,
        "h:mm:ss AM/PM" => 19,
        "h:mm" => 20,
        "h:mm:ss" => 21,
        "m/d/yy h:mm" => 22,
        "#,##0 ;(#,##0)" => 37,
        "#,##0 ;[Red](#,##0)" => 38,
        "#,##0.00;(#,##0.00)" => 39,
        "#,##0.00;[Red](#,##0.00)" => 40,
        "mm:ss" => 45,
        "[h]:mm:ss" => 46,
        "mmss.0" => 47,
        "##0.0E+0" => 48,
        "@" => 49,
        _ => return None,
    })
}
