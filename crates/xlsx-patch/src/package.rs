//! Zip-level access to the xlsx package.
//!
//! The whole point of surgical save: every entry we have no reason to touch
//! is copied RAW — original compressed bytes, original metadata — so content
//! we don't model (charts, external links, web extensions, custom XML…)
//! survives by construction. Only explicitly replaced entries are re-written,
//! and removed entries are skipped.

use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Read, Write};

use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use super::PatchError;

pub struct Package<'a> {
    archive: ZipArchive<Cursor<&'a [u8]>>,
}

impl<'a> Package<'a> {
    pub fn open(bytes: &'a [u8]) -> Result<Self, PatchError> {
        let archive = ZipArchive::new(Cursor::new(bytes))
            .map_err(|e| PatchError::Zip(format!("open: {e}")))?;
        Ok(Self { archive })
    }

    pub fn has_part(&mut self, name: &str) -> bool {
        self.archive.by_name(name).is_ok()
    }

    pub fn read_part(&mut self, name: &str) -> Result<Vec<u8>, PatchError> {
        let mut f = self
            .archive
            .by_name(name)
            .map_err(|e| PatchError::MissingPart(format!("{name}: {e}")))?;
        let mut out = Vec::with_capacity(f.size() as usize);
        f.read_to_end(&mut out)
            .map_err(|e| PatchError::Zip(format!("read {name}: {e}")))?;
        Ok(out)
    }

    /// Rebuild the package: original entry order, raw copies for everything
    /// not in `replaced`/`removed`, brand-new `added` entries appended at
    /// the end. Replaced entries keep their original modification time so a
    /// saved file doesn't churn metadata it didn't need to.
    pub fn rebuild(
        &mut self,
        replaced: &HashMap<String, Vec<u8>>,
        removed: &HashSet<String>,
        added: &[(String, Vec<u8>)],
    ) -> Result<Vec<u8>, PatchError> {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));

        for i in 0..self.archive.len() {
            let name = {
                let f = self
                    .archive
                    .by_index_raw(i)
                    .map_err(|e| PatchError::Zip(format!("index {i}: {e}")))?;
                f.name().to_string()
            };

            if removed.contains(&name) {
                continue;
            }

            if let Some(data) = replaced.get(&name) {
                let original = self
                    .archive
                    .by_index_raw(i)
                    .map_err(|e| PatchError::Zip(format!("index {i}: {e}")))?;
                let mut options = SimpleFileOptions::default()
                    .compression_method(CompressionMethod::Deflated);
                if let Some(mtime) = original.last_modified() {
                    options = options.last_modified_time(mtime);
                }
                drop(original);
                writer
                    .start_file(name.clone(), options)
                    .map_err(|e| PatchError::Zip(format!("start {name}: {e}")))?;
                writer
                    .write_all(data)
                    .map_err(|e| PatchError::Zip(format!("write {name}: {e}")))?;
            } else {
                let raw = self
                    .archive
                    .by_index_raw(i)
                    .map_err(|e| PatchError::Zip(format!("index {i}: {e}")))?;
                writer
                    .raw_copy_file(raw)
                    .map_err(|e| PatchError::Zip(format!("raw copy {name}: {e}")))?;
            }
        }

        for (name, data) in added {
            let options =
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
            writer
                .start_file(name.clone(), options)
                .map_err(|e| PatchError::Zip(format!("start {name}: {e}")))?;
            writer
                .write_all(data)
                .map_err(|e| PatchError::Zip(format!("write {name}: {e}")))?;
        }

        let cursor = writer
            .finish()
            .map_err(|e| PatchError::Zip(format!("finish: {e}")))?;
        Ok(cursor.into_inner())
    }
}

/// A package plus an overlay of pending edits. Structural operations (sheet
/// create/rename/delete, row/col shifts) and cell patches all accumulate
/// here so later stages read each part's CURRENT content, and `finish`
/// raw-copies everything never written to.
pub struct PartStore<'a> {
    pkg: Package<'a>,
    replaced: HashMap<String, Vec<u8>>,
    removed: HashSet<String>,
    added: Vec<(String, Vec<u8>)>,
}

impl<'a> PartStore<'a> {
    pub fn open(bytes: &'a [u8]) -> Result<Self, PatchError> {
        Ok(Self {
            pkg: Package::open(bytes)?,
            replaced: HashMap::new(),
            removed: HashSet::new(),
            added: Vec::new(),
        })
    }

    pub fn read(&mut self, name: &str) -> Result<Vec<u8>, PatchError> {
        if self.removed.contains(name) {
            return Err(PatchError::MissingPart(format!("{name}: removed")));
        }
        if let Some(data) = self.replaced.get(name) {
            return Ok(data.clone());
        }
        if let Some((_, data)) = self.added.iter().find(|(n, _)| n == name) {
            return Ok(data.clone());
        }
        self.pkg.read_part(name)
    }

    pub fn exists(&mut self, name: &str) -> bool {
        !self.removed.contains(name)
            && (self.replaced.contains_key(name)
                || self.added.iter().any(|(n, _)| n == name)
                || self.pkg.has_part(name))
    }

    pub fn write(&mut self, name: &str, data: Vec<u8>) {
        self.removed.remove(name);
        if let Some(slot) = self.added.iter_mut().find(|(n, _)| n == name) {
            slot.1 = data;
            return;
        }
        if self.pkg.has_part(name) {
            self.replaced.insert(name.to_string(), data);
        } else {
            self.added.push((name.to_string(), data));
        }
    }

    pub fn remove(&mut self, name: &str) {
        self.replaced.remove(name);
        self.added.retain(|(n, _)| n != name);
        if self.pkg.has_part(name) {
            self.removed.insert(name.to_string());
        }
    }

    /// Parts this patch wrote (replaced or added) — the validation surface.
    pub fn dirty_names(&self) -> Vec<String> {
        let mut out: Vec<String> = self.replaced.keys().cloned().collect();
        out.extend(self.added.iter().map(|(n, _)| n.clone()));
        out
    }

    /// Parts this patch created (didn't exist in the base package).
    pub fn added_names(&self) -> Vec<String> {
        self.added.iter().map(|(n, _)| n.clone()).collect()
    }

    /// Current part names: original order, minus removed, plus added.
    pub fn names(&self) -> Vec<String> {
        let mut out: Vec<String> = self
            .pkg
            .archive
            .file_names()
            .filter(|n| !self.removed.contains(*n))
            .map(|n| n.to_string())
            .collect();
        out.extend(self.added.iter().map(|(n, _)| n.clone()));
        out
    }

    pub fn finish(mut self) -> Result<Vec<u8>, PatchError> {
        self.pkg
            .rebuild(&self.replaced, &self.removed, &self.added)
    }
}
