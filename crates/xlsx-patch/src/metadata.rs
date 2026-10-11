//! Dynamic-array metadata: `xl/metadata.xml` and the `cm` cell attribute.
//!
//! Excel 365 evaluates a formula in one of two modes depending on how the
//! FILE marks it. A plain `<f>` is a *legacy* formula: arrays in scalar
//! positions are implicitly intersected (Excel shows it as `=@IF(...)`),
//! and the classic multi-criteria idiom `MATCH(1, (rng="x")*(rng2<>"y"), 0)`
//! returns `#N/A` unless array-entered — the agent's row-finder helper
//! columns did exactly that, and every cell keyed off them went blank in
//! Excel while GridPath's engine (which evaluates arrays natively) showed
//! numbers. A cell with `cm="N"` pointing at the XLDAPR dynamic-array
//! metadata record, and `<f t="array" ref="…">`, is a *dynamic-array*
//! formula: arrays evaluate as arrays and may spill. That is the semantics
//! our engine implements, so every formula the patch writes is marked that
//! way; older Excel reads `t="array"` as a CSE formula with the same math.
//!
//! The metadata part is shared workbook-wide: one XLDAPR `metadataType`,
//! one `futureMetadata` block, one `cellMetadata` record, referenced by
//! every dynamic cell through its 1-based index.

use quick_xml::events::{BytesStart, Event};

use super::package::PartStore;
use super::structure::insert_before_container_end;
use super::workbook_xml::{attr_value, local_name, reader};
use super::PatchError;

pub const SHEET_METADATA_CT: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
pub const SHEET_METADATA_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";

/// A fresh part as Excel writes it for a workbook whose first dynamic-array
/// formula was just entered.
const FRESH_METADATA_XML: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    "\r\n",
    r#"<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
    r#"xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">"#,
    r#"<metadataTypes count="1">"#,
    r#"<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" "#,
    r#"merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>"#,
    r#"</metadataTypes>"#,
    r#"<futureMetadata name="XLDAPR" count="1"><bk><extLst>"#,
    r#"<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext>"#,
    r#"</extLst></bk></futureMetadata>"#,
    r#"<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>"#,
    r#"</metadata>"#,
);

/// The `cm` index that marks a cell as a dynamic-array formula, creating
/// the metadata part when the workbook has none. `None` when the workbook
/// already carries a metadata part we cannot safely extend (no XLDAPR
/// record and other metadata present): formulas are then written legacy,
/// exactly as before.
pub fn ensure_dynamic_array_cm(
    store: &mut PartStore,
    wb_dir: &str,
    wb_rels_path: &str,
) -> Result<Option<u32>, PatchError> {
    let rels_xml = store.read(wb_rels_path)?;
    let existing = find_metadata_part(&rels_xml, wb_dir)?;
    if let Some(part) = existing {
        if !store.exists(&part) {
            return Ok(None);
        }
        let xml = store.read(&part)?;
        return Ok(find_xldapr_cm(&xml)?);
    }
    // No metadata part: create one with exactly the dynamic-array record.
    let part = format!("{wb_dir}metadata.xml");
    if store.exists(&part) {
        // A part without a workbook relationship — don't guess, stay legacy.
        return Ok(None);
    }
    store.write(&part, FRESH_METADATA_XML.as_bytes().to_vec());

    let rid = format!("rId{}", max_rid(&rels_xml) + 1);
    let mut rel = BytesStart::new("Relationship");
    rel.push_attribute(("Id", rid.as_str()));
    rel.push_attribute(("Type", SHEET_METADATA_REL_TYPE));
    rel.push_attribute(("Target", "metadata.xml"));
    let new_rels = insert_before_container_end(&rels_xml, b"Relationships", rel)?;
    store.write(wb_rels_path, new_rels);

    let ct = store.read("[Content_Types].xml")?;
    let mut ov = BytesStart::new("Override");
    ov.push_attribute(("PartName", format!("/{part}").as_str()));
    ov.push_attribute(("ContentType", SHEET_METADATA_CT));
    let new_ct = insert_before_container_end(&ct, b"Types", ov)?;
    store.write("[Content_Types].xml", new_ct);
    Ok(Some(1))
}

/// Path of the metadata part the workbook relationships point at, if any.
fn find_metadata_part(rels_xml: &[u8], wb_dir: &str) -> Result<Option<String>, PatchError> {
    let mut r = reader(rels_xml);
    let mut buf = Vec::new();
    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("rels: {e}")))?;
        match ev {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"Relationship" => {
                let ty = attr_value(&e, b"Type").unwrap_or_default();
                if ty.ends_with("/sheetMetadata") {
                    let target = attr_value(&e, b"Target").unwrap_or_default();
                    let path = if let Some(abs) = target.strip_prefix('/') {
                        abs.to_string()
                    } else {
                        format!("{wb_dir}{target}")
                    };
                    return Ok(Some(path));
                }
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(None)
}

fn max_rid(rels_xml: &[u8]) -> u32 {
    let mut max = 0u32;
    let text = String::from_utf8_lossy(rels_xml);
    for piece in text.split("Id=\"rId").skip(1) {
        let digits: String = piece.chars().take_while(|c| c.is_ascii_digit()).collect();
        if let Ok(n) = digits.parse::<u32>() {
            max = max.max(n);
        }
    }
    max
}

/// In an existing metadata part: the 1-based `cellMetadata` record index
/// whose `rc` points at the XLDAPR metadata type, if both exist.
fn find_xldapr_cm(xml: &[u8]) -> Result<Option<u32>, PatchError> {
    let mut r = reader(xml);
    let mut buf = Vec::new();
    let mut type_index = 0u32; // 1-based position among metadataType entries
    let mut xldapr_type: Option<u32> = None;
    let mut in_cell_metadata = false;
    let mut bk_index = 0u32;
    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("metadata: {e}")))?;
        match ev {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref()).to_vec();
                match name.as_slice() {
                    b"metadataType" => {
                        type_index += 1;
                        if attr_value(&e, b"name").as_deref() == Some("XLDAPR") {
                            xldapr_type = Some(type_index);
                        }
                    }
                    b"cellMetadata" => in_cell_metadata = true,
                    b"bk" if in_cell_metadata => {
                        bk_index += 1;
                    }
                    b"rc" if in_cell_metadata => {
                        let rc_t = attr_value(&e, b"t").and_then(|v| v.parse::<u32>().ok());
                        if let (Some(t), Some(x)) = (rc_t, xldapr_type) {
                            if t == x {
                                return Ok(Some(bk_index));
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                if local_name(e.name().as_ref()) == b"cellMetadata" {
                    in_cell_metadata = false;
                }
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(None)
}
