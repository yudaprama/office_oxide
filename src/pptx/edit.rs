//! PPTX editing via raw XML text replacement.
//!
//! Uses the `EditablePackage` from core to preserve all parts,
//! replacing text in slide XML `<a:t>` elements.

use crate::core::editable::EditablePackage;
use crate::core::opc::PartName;

use super::Result;

/// An editable PPTX document that supports text replacement and saving.
pub struct EditablePptx {
    package: EditablePackage,
}

impl EditablePptx {
    /// Open a PPTX file for editing.
    pub fn open(path: impl AsRef<std::path::Path>) -> Result<Self> {
        let package = EditablePackage::open(&path)?;
        Ok(Self { package })
    }

    /// Open from any `Read + Seek` source.
    pub fn from_reader<R: std::io::Read + std::io::Seek>(reader: R) -> Result<Self> {
        let package = EditablePackage::from_reader(reader)?;
        Ok(Self { package })
    }

    /// Replace all occurrences of `find` with `replace` across all slides.
    /// Returns the total number of replacements made.
    pub fn replace_text(&mut self, find: &str, replace: &str) -> usize {
        let mut total = 0;

        // Find all slide parts
        for i in 1..=100 {
            let part_name = match PartName::new(&format!("/ppt/slides/slide{i}.xml")) {
                Ok(pn) => pn,
                Err(_) => break,
            };
            let Some(data) = self.package.get_part(&part_name) else {
                break;
            };
            let xml_str = String::from_utf8_lossy(data);
            let (new_xml, count) = replace_in_at_elements(&xml_str, find, replace);
            if count > 0 {
                self.package.set_part(part_name, new_xml.into_bytes());
                total += count;
            }
        }

        total
    }

    /// Save the edited document to a file.
    pub fn save(&self, path: impl AsRef<std::path::Path>) -> Result<()> {
        self.package.save(path)?;
        Ok(())
    }

    /// Write the edited document to any `Write + Seek` destination.
    pub fn write_to<W: std::io::Write + std::io::Seek>(&self, writer: W) -> Result<()> {
        self.package.write_to(writer)?;
        Ok(())
    }

    /// Append slides to the deck by generating them with the same writer the
    /// create path uses and splicing the produced slide parts (plus their
    /// relationships) into the existing package. Returns the number appended.
    pub fn append_slides(&mut self, slides: &[crate::edit::PptxSlideSpec]) -> crate::Result<usize> {
        use std::io::Cursor;

        if slides.is_empty() {
            return Ok(0);
        }

        // Generate a throwaway deck containing exactly the new slides.
        let mut w = crate::pptx::write::PptxWriter::new();
        for spec in slides {
            let slide = w.add_slide();
            slide.set_title(&spec.title);
            let body: Vec<&str> = spec.body.iter().map(|s| s.as_str()).collect();
            if !body.is_empty() {
                slide.add_bullet_list(&body);
            }
        }
        let mut buf = Cursor::new(Vec::new());
        w.write_to(&mut buf)?;
        let temp = crate::core::editable::EditablePackage::from_reader(Cursor::new(buf.into_inner()))?;

        let presentation_part = PartName::new("/ppt/presentation.xml")?;
        let max_existing = max_slide_index(&self.package);
        let slide_ct = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

        let mut appended = 0usize;
        for (i, spec) in slides.iter().enumerate() {
            let src = i + 1;
            let src_part = PartName::new(&format!("/ppt/slides/slide{src}.xml"))?;
            let Some(data) = temp.get_part(&src_part) else {
                continue;
            };
            let target_index = max_existing + 1 + i as u32;
            let target_part = PartName::new(&format!("/ppt/slides/slide{target_index}.xml"))?;
            self.package
                .add_part_with_content_type(target_part.clone(), data.to_vec(), slide_ct);

            // Copy the slide's relationships part verbatim.
            let src_rels = PartName::new(&format!("/ppt/slides/_rels/slide{src}.xml.rels"))?;
            if let Some(rels_data) = temp.get_part(&src_rels) {
                let target_rels =
                    PartName::new(&format!("/ppt/slides/_rels/slide{target_index}.xml.rels"))?;
                self.package.set_part(target_rels, rels_data.to_vec());
            }

            // Register the slide with the presentation relationships.
            let rid = self.package.add_relationship(
                &presentation_part,
                crate::core::relationships::rel_types::SLIDE,
                &format!("slides/slide{target_index}.xml"),
            );

            // Insert a <p:sldId> entry into the presentation's sldIdLst.
            self.insert_sld_id(&presentation_part, &rid)?;
            appended += 1;
            let _ = spec;
        }
        Ok(appended)
    }

    /// Insert a `<p:sldId r:id="..."/>` entry into `/ppt/presentation.xml`'s
    /// `<p:sldIdLst>`, assigning a unique (>= 256) id.
    fn insert_sld_id(&mut self, presentation_part: &PartName, rid: &str) -> Result<()> {
        let Some(data) = self.package.get_part(presentation_part) else {
            return Err(super::PptxError::Core(crate::core::Error::MissingPart(
                presentation_part.as_str().to_string(),
            )));
        };
        let xml = String::from_utf8_lossy(data).into_owned();
        let new_id = next_sld_id(&xml);
        let entry = format!(r#"<p:sldId id="{new_id}" r:id="{rid}"/>"#);

        let new_xml = if let Some(lst_start) = xml.find("<p:sldIdLst") {
            // Insert before the closing </p:sldIdLst>.
            let close = xml[lst_start..]
                .find("</p:sldIdLst>")
                .map(|o| lst_start + o)
                .unwrap_or_else(|| xml.len());
            let mut out = String::with_capacity(xml.len() + entry.len());
            out.push_str(&xml[..close]);
            out.push_str(&entry);
            out.push_str(&xml[close..]);
            out
        } else {
            // No sldIdLst: create one right after the <p:presentation ...> tag.
            let open = xml.find("<p:presentation").and_then(|s| xml[s..].find('>').map(|o| s + o + 1));
            let open = open.unwrap_or(0);
            let mut out = String::with_capacity(xml.len() + entry.len() + 40);
            out.push_str(&xml[..open]);
            out.push_str(&format!("<p:sldIdLst>{entry}</p:sldIdLst>"));
            out.push_str(&xml[open..]);
            out
        };
        self.package
            .set_part(presentation_part.clone(), new_xml.into_bytes());
        Ok(())
    }

    /// Remove the first slide whose text contains `find`.
    /// Returns 1 if a slide was removed, 0 otherwise.
    pub fn remove_slide(&mut self, find: &str) -> crate::Result<usize> {
        let presentation_part = PartName::new("/ppt/presentation.xml")?;
        for n in 1..=200 {
            let part = PartName::new(&format!("/ppt/slides/slide{n}.xml"))?;
            let Some(data) = self.package.get_part(&part) else {
                break;
            };
            let xml = String::from_utf8_lossy(data);
            if slide_text_contains(&xml, find) {
                // Resolve the relationship that points at this slide.
                let target = format!("slides/slide{n}.xml");
                let mut rid_to_remove: Option<String> = None;
                if let Some(rels) = self.package.part_rels(&presentation_part) {
                    for r in rels.all() {
                        if r.rel_type == crate::core::relationships::rel_types::SLIDE
                            && r.target == target
                        {
                            rid_to_remove = Some(r.id.clone());
                        }
                    }
                }
                // Remove the sldId entry from presentation.xml.
                if let Some(rid) = &rid_to_remove {
                    if let Some(pdata) = self.package.get_part(&presentation_part) {
                        let pxml = String::from_utf8_lossy(pdata).into_owned();
                        let new_pxml = remove_sld_id_entry(&pxml, rid);
                        self.package
                            .set_part(presentation_part.clone(), new_pxml.into_bytes());
                    }
                    self.package.remove_relationship(&presentation_part, rid);
                }
                // Remove the slide part and its rels.
                self.package.remove_part(&part);
                let rels_part = PartName::new(&format!("/ppt/slides/_rels/slide{n}.xml.rels"))?;
                self.package.remove_part(&rels_part);
                return Ok(1);
            }
        }
        Ok(0)
    }
}

/// Highest existing `/ppt/slides/slideN.xml` index, or 0 if none.
fn max_slide_index(pkg: &crate::core::editable::EditablePackage) -> u32 {
    let mut max = 0u32;
    for n in 1..=200 {
        let part = match PartName::new(&format!("/ppt/slides/slide{n}.xml")) {
            Ok(p) => p,
            Err(_) => break,
        };
        if pkg.get_part(&part).is_some() {
            max = n;
        } else {
            break;
        }
    }
    max
}

/// Next unique `<p:sldId>` id (>= 256).
fn next_sld_id(presentation_xml: &str) -> u32 {
    let mut max = 255u32;
    let mut scan = presentation_xml;
    while let Some(pos) = scan.find("<p:sldId") {
        let rest = &scan[pos..];
        let Some(id_pos) = rest.find("id=\"") else { break };
        let num_start = id_pos + 4;
        let Some(num_end) = rest[num_start..].find('"') else { break };
        if let Ok(num) = rest[num_start..num_start + num_end].parse::<u32>() {
            max = max.max(num);
        }
        let Some(next) = scan[pos + 1..].find("<p:sldId") else {
            break;
        };
        scan = &scan[pos + 1 + next..];
    }
    max + 1
}

/// Remove the `<p:sldId ... r:id="rid"/>` entry referencing `rid`.
fn remove_sld_id_entry(presentation_xml: &str, rid: &str) -> String {
    let needle = format!(r#"r:id="{rid}""#);
    let mut out = String::with_capacity(presentation_xml.len());
    let mut pos = 0;
    let mut removed = false;
    while pos < presentation_xml.len() {
        if !removed && presentation_xml[pos..].contains("<p:sldId") {
            let start = pos + presentation_xml[pos..].find("<p:sldId").unwrap();
            let Some(end) = presentation_xml[start..].find("/>") else {
                out.push_str(&presentation_xml[pos..]);
                break;
            };
            let end = start + end + 2;
            let span = &presentation_xml[start..end];
            if span.contains(&needle) {
                removed = true;
                pos = end;
                continue;
            }
            out.push_str(&presentation_xml[pos..end]);
            pos = end;
        } else {
            out.push_str(&presentation_xml[pos..]);
            break;
        }
    }
    out
}

/// Extract the visible text of a slide (all `<a:t>` elements concatenated).
fn slide_text_contains(slide_xml: &str, find: &str) -> bool {
    let mut text = String::new();
    let mut pos = 0;
    while let Some(ts) = slide_xml[pos..].find("<a:t") {
        let ts = pos + ts;
        let Some(te) = slide_xml[ts..].find('>') else { break };
        let te = ts + te + 1;
        let Some(ce) = slide_xml[te..].find("</a:t>") else { break };
        let ce = te + ce;
        text.push_str(&slide_xml[te..ce]);
        pos = ce;
    }
    text.contains(find)
}

/// Replace text within `<a:t>...</a:t>` elements in DrawingML XML.
fn replace_in_at_elements(xml: &str, find: &str, replace: &str) -> (String, usize) {
    let mut result = String::with_capacity(xml.len());
    let mut count = 0;
    let mut pos = 0;

    while pos < xml.len() {
        if let Some(tag_start) = xml[pos..].find("<a:t") {
            let tag_start = pos + tag_start;

            let Some(tag_end_offset) = xml[tag_start..].find('>') else {
                result.push_str(&xml[pos..]);
                break;
            };
            let tag_end = tag_start + tag_end_offset + 1;

            // Self-closing tag
            if xml[tag_start..tag_end].ends_with("/>") {
                result.push_str(&xml[pos..tag_end]);
                pos = tag_end;
                continue;
            }

            let Some(close_offset) = xml[tag_end..].find("</a:t>") else {
                result.push_str(&xml[pos..]);
                break;
            };
            let close_start = tag_end + close_offset;

            let text_content = &xml[tag_end..close_start];
            let occ = text_content.matches(find).count();
            count += occ;

            let replaced = text_content.replace(find, replace);
            result.push_str(&xml[pos..tag_end]);
            result.push_str(&replaced);

            pos = close_start;
        } else {
            result.push_str(&xml[pos..]);
            break;
        }
    }

    (result, count)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn replace_in_at_simple() {
        let xml = r#"<a:p><a:r><a:t>Hello World</a:t></a:r></a:p>"#;
        let (result, count) = replace_in_at_elements(xml, "World", "PPTX");
        assert_eq!(count, 1);
        assert!(result.contains("<a:t>Hello PPTX</a:t>"));
    }

    #[test]
    fn replace_in_at_multiple_runs() {
        let xml = r#"<a:r><a:t>foo</a:t></a:r><a:r><a:t>foo</a:t></a:r>"#;
        let (result, count) = replace_in_at_elements(xml, "foo", "bar");
        assert_eq!(count, 2);
        assert_eq!(result.matches("bar").count(), 2);
    }

    #[test]
    fn no_match_returns_zero() {
        let xml = r#"<a:r><a:t>Hello</a:t></a:r>"#;
        let (result, count) = replace_in_at_elements(xml, "xyz", "abc");
        assert_eq!(count, 0);
        assert_eq!(result, xml);
    }
}
