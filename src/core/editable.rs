//! Editable OPC package for read-modify-write roundtrips.
//!
//! Loads all parts and relationships into memory so unmodified parts
//! can be written back verbatim (preserving images, charts, custom XML, etc.).

use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Seek, Write};
use std::path::Path;

use zip::CompressionMethod;
use zip::write::{SimpleFileOptions, ZipWriter};

use super::content_types::{ContentTypes, ContentTypesBuilder};
use super::error::Result;
use super::opc::PartName;
use super::relationships::{Relationships, RelationshipsBuilder};

/// A mutable in-memory representation of an OPC package.
///
/// All parts are loaded into memory so individual parts can be replaced
/// while everything else is preserved on save.
pub struct EditablePackage {
    /// Raw bytes for each part.
    parts: HashMap<PartName, Vec<u8>>,
    /// Content type mapping.
    content_types: ContentTypes,
    /// Package-level relationships (_rels/.rels).
    package_rels: Relationships,
    /// Part-level relationships keyed by part name.
    part_rels: HashMap<PartName, Relationships>,
}

impl EditablePackage {
    /// Load an OPC package into an editable in-memory representation.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = File::open(path)?;
        Self::from_reader(file)
    }

    /// Load from any `Read + Seek` source.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let mut opc = super::opc::OpcReader::new(reader)?;
        let content_types = opc.content_types().clone();
        let package_rels = opc.package_rels().clone();

        let part_names = opc.part_names();
        let mut parts = HashMap::new();
        let mut part_rels = HashMap::new();

        for name in &part_names {
            let data = opc.read_part(name)?;
            parts.insert(name.clone(), data);

            let rels = opc.read_rels_for(name)?;
            if !rels.all().is_empty() {
                part_rels.insert(name.clone(), rels);
            }
        }

        Ok(Self {
            parts,
            content_types,
            package_rels,
            part_rels,
        })
    }

    /// Get a part's raw bytes.
    pub fn get_part(&self, name: &PartName) -> Option<&[u8]> {
        self.parts.get(name).map(|v| v.as_slice())
    }

    /// Replace or insert a part's raw bytes.
    pub fn set_part(&mut self, name: PartName, data: Vec<u8>) {
        self.parts.insert(name, data);
    }

    /// Remove a part, returning its bytes if present.
    pub fn remove_part(&mut self, name: &PartName) -> Option<Vec<u8>> {
        self.content_types.remove_override(name);
        self.part_rels.remove(name);
        self.parts.remove(name)
    }

    /// Add a part and register its content type as an override.
    pub fn add_part_with_content_type(
        &mut self,
        name: PartName,
        data: Vec<u8>,
        content_type: &str,
    ) {
        self.parts.insert(name.clone(), data);
        self.content_types.add_override(name, content_type);
    }

    /// Add a relationship for `source` (a part name; use `/` for the package
    /// root). Returns the newly assigned rId.
    pub fn add_relationship(
        &mut self,
        source: &PartName,
        rel_type: &str,
        target: &str,
    ) -> String {
        let next = self.next_rid_all();
        let id = format!("rId{next}");
        let rel = super::relationships::Relationship {
            id: id.clone(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: super::relationships::TargetMode::Internal,
        };
        self.part_rels
            .entry(source.clone())
            .or_insert_with(super::relationships::Relationships::empty)
            .add(rel);
        id
    }

    /// Remove the relationship with `id` from `source`'s relationships.
    pub fn remove_relationship(&mut self, source: &PartName, id: &str) {
        if let Some(rels) = self.part_rels.get_mut(source) {
            rels.remove_by_id(id);
        }
        if self.part_rels.get(source).map(|r| r.all().is_empty()) == Some(true) {
            self.part_rels.remove(source);
        }
    }

    /// Next rId number across package-level and all part-level rels.
    fn next_rid_all(&self) -> u32 {
        let mut max = 0u32;
        for r in self.package_rels.all() {
            max = max.max(parse_rid(&r.id));
        }
        for rels in self.part_rels.values() {
            for r in rels.all() {
                max = max.max(parse_rid(&r.id));
            }
        }
        max + 1
    }

    /// Get the content types table.
    pub fn content_types(&self) -> &ContentTypes {
        &self.content_types
    }

    /// Get the package-level relationships.
    pub fn package_rels(&self) -> &Relationships {
        &self.package_rels
    }

    /// Get part-level relationships for a part.
    pub fn part_rels(&self, name: &PartName) -> Option<&Relationships> {
        self.part_rels.get(name)
    }

    /// Save the package to a file.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        let file = File::create(path)?;
        self.write_to(file)
    }

    /// Write the package to any `Write + Seek` destination.
    pub fn write_to<W: Write + Seek>(&self, writer: W) -> Result<()> {
        let mut zip = ZipWriter::new(writer);
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

        // Write all parts
        for (name, data) in &self.parts {
            let zip_path = &name.as_str()[1..]; // strip leading /
            zip.start_file(zip_path, options)?;
            zip.write_all(data)?;
        }

        // Write part-level .rels files
        for (source, rels) in &self.part_rels {
            if rels.all().is_empty() {
                continue;
            }
            let rels_path = source.rels_path();
            let zip_path = &rels_path[1..];
            let mut builder = RelationshipsBuilder::new();
            for rel in rels.all() {
                builder.add_with_id(&rel.id, &rel.rel_type, &rel.target, rel.target_mode);
            }
            let data = builder.serialize();
            zip.start_file(zip_path, options)?;
            zip.write_all(&data)?;
        }

        // Write _rels/.rels
        {
            let mut builder = RelationshipsBuilder::new();
            for rel in self.package_rels.all() {
                builder.add_with_id(&rel.id, &rel.rel_type, &rel.target, rel.target_mode);
            }
            let data = builder.serialize();
            zip.start_file("_rels/.rels", options)?;
            zip.write_all(&data)?;
        }

        // Write [Content_Types].xml
        {
            let mut ct_builder = ContentTypesBuilder::new();
            for (ext, ct) in self.content_types.defaults() {
                ct_builder.add_default(ext, ct);
            }
            for (pn, ct) in self.content_types.overrides() {
                ct_builder.add_override(pn.clone(), ct);
            }
            let data = ct_builder.serialize();
            zip.start_file("[Content_Types].xml", options)?;
            zip.write_all(&data)?;
        }

        zip.finish()?;
        Ok(())
    }
}

/// Parse the numeric suffix of an rId (e.g. "rId12" -> 12). Returns 0 if none.
fn parse_rid(id: &str) -> u32 {
    id.strip_prefix("rId")
        .and_then(|n| n.parse::<u32>().ok())
        .unwrap_or(0)
}
