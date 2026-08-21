//! DOCX editing via raw XML text replacement.
//!
//! Uses the `EditablePackage` from core to preserve all parts,
//! replacing text in the document.xml body by string substitution
//! in the raw XML `<w:t>` elements.

use crate::core::editable::EditablePackage;
use crate::core::opc::PartName;

use super::Result;

/// An editable DOCX document that supports text replacement and saving.
pub struct EditableDocx {
    package: EditablePackage,
    main_part: PartName,
}

impl EditableDocx {
    /// Open a DOCX file for editing.
    pub fn open(path: impl AsRef<std::path::Path>) -> Result<Self> {
        let package = EditablePackage::open(&path)?;
        let main_part = PartName::new("/word/document.xml")?;
        Ok(Self { package, main_part })
    }

    /// Open from any `Read + Seek` source.
    pub fn from_reader<R: std::io::Read + std::io::Seek>(reader: R) -> Result<Self> {
        let package = EditablePackage::from_reader(reader)?;
        let main_part = PartName::new("/word/document.xml")?;
        Ok(Self { package, main_part })
    }

    /// Replace all occurrences of `find` with `replace` in the document body.
    /// Returns the number of replacements made.
    pub fn replace_text(&mut self, find: &str, replace: &str) -> usize {
        let Some(data) = self.package.get_part(&self.main_part) else {
            return 0;
        };
        let xml_str = String::from_utf8_lossy(data);

        // Replace text within <w:t> elements.
        // Strategy: find text between <w:t...> and </w:t> tags and do replacements there.
        let (new_xml, count) = replace_in_wt_elements(&xml_str, find, replace);

        if count > 0 {
            self.package
                .set_part(self.main_part.clone(), new_xml.into_bytes());
        }
        count
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

    /// Append paragraph/heading/list blocks to the end of the document body.
    ///
    /// Returns the number of blocks appended. Faithful generation delegates to
    /// the same writer the create path uses, so the produced `<w:p>`/`<w:tbl>`
    /// elements are valid OOXML spliced in verbatim before `</w:body>`.
    pub fn append_blocks(&mut self, blocks: &[crate::edit::DocxBlock]) -> crate::Result<usize> {
        let fragment = generate_docx_fragment(blocks)?;
        if fragment.is_empty() {
            return Ok(0);
        }
        let Some(data) = self.package.get_part(&self.main_part) else {
            return Ok(0);
        };
        let xml = String::from_utf8_lossy(data).into_owned();
        if let Some(pos) = xml.rfind("</w:body>") {
            let mut out = String::with_capacity(xml.len() + fragment.len());
            out.push_str(&xml[..pos]);
            out.push_str(&fragment);
            out.push_str(&xml[pos..]);
            self.package
                .set_part(self.main_part.clone(), out.into_bytes());
            Ok(blocks.len())
        } else {
            Ok(0)
        }
    }

    /// Append a table to the end of the document body.
    /// Returns 1 if a table was appended, 0 otherwise.
    pub fn append_table(&mut self, rows: &[Vec<String>]) -> crate::Result<usize> {
        if rows.is_empty() {
            return Ok(0);
        }
        let fragment = generate_docx_table(rows)?;
        let Some(data) = self.package.get_part(&self.main_part) else {
            return Ok(0);
        };
        let xml = String::from_utf8_lossy(data).into_owned();
        if let Some(pos) = xml.rfind("</w:body>") {
            let mut out = String::with_capacity(xml.len() + fragment.len());
            out.push_str(&xml[..pos]);
            out.push_str(&fragment);
            out.push_str(&xml[pos..]);
            self.package
                .set_part(self.main_part.clone(), out.into_bytes());
            Ok(1)
        } else {
            Ok(0)
        }
    }

    /// Delete every paragraph whose visible text contains `find`.
    /// Returns the number of paragraphs removed.
    pub fn delete_paragraphs(&mut self, find: &str) -> crate::Result<usize> {
        let Some(data) = self.package.get_part(&self.main_part) else {
            return Ok(0);
        };
        let xml = String::from_utf8_lossy(data).into_owned();
        let new_xml = remove_paragraphs_containing(&xml, find);
        let count = count_paragraph_removals(&xml, &new_xml);
        if count > 0 {
            self.package
                .set_part(self.main_part.clone(), new_xml.into_bytes());
        }
        Ok(count)
    }

    /// Apply formatting to the first paragraph whose visible text contains `find`.
    /// Returns 1 if a paragraph was formatted, 0 otherwise.
    pub fn format_paragraph(&mut self, find: &str, fmt: &crate::edit::DocxFormat) -> crate::Result<usize> {
        let Some(data) = self.package.get_part(&self.main_part) else {
            return Ok(0);
        };
        let xml = String::from_utf8_lossy(data).into_owned();
        if let Some(new_xml) = format_paragraph_containing(&xml, find, fmt) {
            self.package
                .set_part(self.main_part.clone(), new_xml.into_bytes());
            Ok(1)
        } else {
            Ok(0)
        }
    }
}

/// Build a DOCX XML fragment (the children of `<w:body>`) for the given blocks
/// by writing a throwaway document and extracting its body contents.
fn generate_docx_fragment(blocks: &[crate::edit::DocxBlock]) -> crate::Result<String> {
    use std::io::Cursor;
    let mut w = crate::docx::write::DocxWriter::new();
    for block in blocks {
        let runs: Vec<crate::docx::write::Run> = block
            .runs
            .iter()
            .map(|r| {
                let mut run = crate::docx::write::Run::new(r.text.clone());
                if r.bold {
                    run = run.bold();
                }
                if r.italic {
                    run = run.italic();
                }
                if let Some(size) = r.size {
                    run = run.font_size(size);
                }
                if let Some(color) = &r.color {
                    run = run.color(color.clone());
                }
                if let Some(font) = &r.font {
                    run = run.font(font.clone());
                }
                run
            })
            .collect();
        let text: String = block.runs.iter().map(|r| r.text.clone()).collect();
        match block.kind {
            crate::edit::DocxBlockKind::Paragraph => {
                w.add_rich_paragraph(&runs);
            },
            crate::edit::DocxBlockKind::Bullet => {
                let items: Vec<&str> = if runs.is_empty() {
                    vec![text.as_str()]
                } else {
                    block.runs.iter().map(|r| r.text.as_str()).collect()
                };
                w.add_list(&items, false);
            },
            crate::edit::DocxBlockKind::Title => {
                w.add_heading(&text, 0);
            },
            crate::edit::DocxBlockKind::Heading1 => {
                w.add_heading(&text, 1);
            },
            crate::edit::DocxBlockKind::Heading2 => {
                w.add_heading(&text, 2);
            },
            crate::edit::DocxBlockKind::Heading3 => {
                w.add_heading(&text, 3);
            },
        }
    }
    let mut buf = Cursor::new(Vec::new());
    w.write_to(&mut buf)?;
    let pkg = crate::core::editable::EditablePackage::from_reader(Cursor::new(buf.into_inner()))?;
    let part = pkg
        .get_part(&crate::core::opc::PartName::new("/word/document.xml")?)
        .ok_or_else(|| {
            crate::OfficeError::Core(crate::core::Error::MissingPart(
                "/word/document.xml".to_string(),
            ))
        })?;
    let xml = String::from_utf8_lossy(part);
    Ok(extract_body_children(&xml).to_string())
}

/// Build a DOCX table fragment from string rows.
fn generate_docx_table(rows: &[Vec<String>]) -> crate::Result<String> {
    use std::io::Cursor;
    let borrowed: Vec<Vec<&str>> = rows
        .iter()
        .map(|r| r.iter().map(|s| s.as_str()).collect())
        .collect();
    let mut w = crate::docx::write::DocxWriter::new();
    w.add_table(&borrowed);
    let mut buf = Cursor::new(Vec::new());
    w.write_to(&mut buf)?;
    let pkg = crate::core::editable::EditablePackage::from_reader(Cursor::new(buf.into_inner()))?;
    let part = pkg
        .get_part(&crate::core::opc::PartName::new("/word/document.xml")?)
        .ok_or_else(|| {
            crate::OfficeError::Core(crate::core::Error::MissingPart(
                "/word/document.xml".to_string(),
            ))
        })?;
    let xml = String::from_utf8_lossy(part);
    Ok(extract_body_children(&xml).to_string())
}

/// Extract the children of `<w:body>` (everything between the tags) as a string.
fn extract_body_children(xml: &str) -> &str {
    let start = xml.find("<w:body").and_then(|s| xml[s..].find('>').map(|o| s + o + 1));
    let end = xml.rfind("</w:body>");
    match (start, end) {
        (Some(s), Some(e)) if e > s => &xml[s..e],
        _ => "",
    }
}

/// Extract the visible text of a single `<w:p>...</w:p>` span (tags stripped).
fn paragraph_text(span: &str) -> String {
    let mut out = String::new();
    let mut in_tag = false;
    for ch in span.chars() {
        match ch {
            '<' => in_tag = true,
            '>' => in_tag = false,
            _ if !in_tag => out.push(ch),
            _ => {},
        }
    }
    unescape_xml(&out)
}

fn unescape_xml(s: &str) -> String {
    s.replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
}

/// Remove every `<w:p>` element whose text contains `find`.
fn remove_paragraphs_containing(xml: &str, find: &str) -> String {
    let mut result = String::with_capacity(xml.len());
    let mut pos = 0;
    while pos < xml.len() {
        let Some(p_start) = xml[pos..].find("<w:p") else {
            result.push_str(&xml[pos..]);
            break;
        };
        let p_start = pos + p_start;
        let Some(tag_end) = xml[p_start..].find('>') else {
            result.push_str(&xml[pos..]);
            break;
        };
        let tag_end = p_start + tag_end + 1;
        let Some(close) = xml[tag_end..].find("</w:p>") else {
            result.push_str(&xml[pos..]);
            break;
        };
        let close_start = tag_end + close;
        let close_end = close_start + "</w:p>".len();
        let span = &xml[p_start..close_end];
        if paragraph_text(span).contains(find) {
            // Drop this paragraph entirely.
            pos = close_end;
            continue;
        }
        result.push_str(&xml[pos..close_end]);
        pos = close_end;
    }
    result
}

/// Count how many paragraphs were removed between two XML strings by comparing
/// the number of `<w:p` openings.
fn count_paragraph_removals(old_xml: &str, new_xml: &str) -> usize {
    let count_open = |s: &str| s.matches("<w:p").count();
    count_open(old_xml).saturating_sub(count_open(new_xml))
}

/// Format the first `<w:p>` whose text contains `find` by injecting/merging a
/// `<w:pPr>` with the requested properties.
fn format_paragraph_containing(xml: &str, find: &str, fmt: &crate::edit::DocxFormat) -> Option<String> {
    let mut result = String::with_capacity(xml.len());
    let mut pos = 0;
    let mut formatted = false;
    while pos < xml.len() {
        let Some(p_start) = xml[pos..].find("<w:p") else {
            result.push_str(&xml[pos..]);
            break;
        };
        let p_start = pos + p_start;
        let Some(tag_end) = xml[p_start..].find('>') else {
            result.push_str(&xml[pos..]);
            break;
        };
        let tag_end = p_start + tag_end + 1;
        let Some(close) = xml[tag_end..].find("</w:p>") else {
            result.push_str(&xml[pos..]);
            break;
        };
        let close_start = tag_end + close;
        let close_end = close_start + "</w:p>".len();
        let span = &xml[p_start..close_end];
        if !formatted && paragraph_text(span).contains(find) {
            // Build the pPr element.
            let mut ppr = String::new();
            if let Some(align) = fmt.alignment {
                let v = match align {
                    crate::edit::DocxAlign::Left => "left",
                    crate::edit::DocxAlign::Center => "center",
                    crate::edit::DocxAlign::Right => "right",
                    crate::edit::DocxAlign::Justify => "both",
                };
                ppr.push_str(&format!(r#"<w:jc w:val="{v}"/>"#));
            }
            let mut spacing = String::new();
            if let Some(b) = fmt.spacing_before {
                spacing.push_str(&format!(r#" w:before="{b}""#));
            }
            if let Some(a) = fmt.spacing_after {
                spacing.push_str(&format!(r#" w:after="{a}""#));
            }
            if !spacing.is_empty() {
                ppr.push_str(&format!("<w:spacing{spacing}/>"));
            }
            let mut ind = String::new();
            if let Some(l) = fmt.indent_left {
                ind.push_str(&format!(r#" w:left="{l}""#));
            }
            if let Some(r) = fmt.indent_right {
                ind.push_str(&format!(r#" w:right="{r}""#));
            }
            if !ind.is_empty() {
                ppr.push_str(&format!("<w:ind{ind}/>"));
            }
            if ppr.is_empty() {
                result.push_str(&xml[pos..close_end]);
                pos = close_end;
                continue;
            }
            // Determine where the existing <w:pPr> is, if any.
            let inner = &xml[tag_end..close_start];
            if let Some(ppr_start) = inner.find("<w:pPr") {
                let ppr_start_abs = tag_end + ppr_start;
                let Some(ppr_tag_end) = xml[ppr_start_abs..].find('>') else {
                    result.push_str(&xml[pos..close_end]);
                    pos = close_end;
                    continue;
                };
                let ppr_tag_end = ppr_start_abs + ppr_tag_end + 1;
                let ppr_close = xml[ppr_tag_end..].find("</w:pPr>");
                let insert_at = if let Some(c) = ppr_close {
                    ppr_tag_end + c // position just before </w:pPr>
                } else {
                    ppr_tag_end // right after <w:pPr> (self-closing or open)
                };
                result.push_str(&xml[pos..insert_at]);
                result.push_str(&ppr);
                result.push_str(&xml[insert_at..close_end]);
            } else {
                // No pPr: insert one right after the opening <w:p ...> tag.
                result.push_str(&xml[pos..tag_end]);
                result.push_str(&format!("<w:pPr>{ppr}</w:pPr>"));
                result.push_str(&xml[tag_end..close_end]);
            }
            formatted = true;
            pos = close_end;
            continue;
        }
        result.push_str(&xml[pos..close_end]);
        pos = close_end;
    }
    if formatted {
        Some(result)
    } else {
        None
    }
}

/// Replace text within `<w:t>...</w:t>` elements in a WML XML string.
/// Returns the new string and the count of replacements.
fn replace_in_wt_elements(xml: &str, find: &str, replace: &str) -> (String, usize) {
    let mut result = String::with_capacity(xml.len());
    let mut count = 0;
    let mut pos = 0;

    while pos < xml.len() {
        // Find next <w:t> or <w:t ...>
        if let Some(tag_start) = xml[pos..].find("<w:t") {
            let tag_start = pos + tag_start;

            // Find the end of the opening tag
            let Some(tag_end_offset) = xml[tag_start..].find('>') else {
                result.push_str(&xml[pos..]);
                break;
            };
            let tag_end = tag_start + tag_end_offset + 1;

            // Check if it's a self-closing tag
            if xml[tag_start..tag_end].ends_with("/>") {
                result.push_str(&xml[pos..tag_end]);
                pos = tag_end;
                continue;
            }

            // Find closing </w:t>
            let Some(close_offset) = xml[tag_end..].find("</w:t>") else {
                result.push_str(&xml[pos..]);
                break;
            };
            let close_start = tag_end + close_offset;

            // Extract text content between tags
            let text_content = &xml[tag_end..close_start];
            count += text_content.matches(find).count();
            let replaced = text_content.replace(find, replace);

            // Write: everything before this tag + tag + replaced text + close tag
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
    fn replace_in_wt_simple() {
        let xml = r#"<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>"#;
        let (result, count) = replace_in_wt_elements(xml, "World", "Rust");
        assert_eq!(count, 1);
        assert!(result.contains("<w:t>Hello Rust</w:t>"));
    }

    #[test]
    fn replace_in_wt_multiple() {
        let xml = r#"<w:r><w:t>foo bar foo</w:t></w:r>"#;
        let (result, count) = replace_in_wt_elements(xml, "foo", "baz");
        assert_eq!(count, 2);
        assert!(result.contains("<w:t>baz bar baz</w:t>"));
    }

    #[test]
    fn replace_preserves_attributes() {
        let xml = r#"<w:r><w:t xml:space="preserve"> Hello </w:t></w:r>"#;
        let (result, count) = replace_in_wt_elements(xml, "Hello", "World");
        assert_eq!(count, 1);
        assert!(result.contains(r#"xml:space="preserve"> World </w:t>"#));
    }

    #[test]
    fn no_match_returns_zero() {
        let xml = r#"<w:r><w:t>Hello</w:t></w:r>"#;
        let (result, count) = replace_in_wt_elements(xml, "xyz", "abc");
        assert_eq!(count, 0);
        assert_eq!(result, xml);
    }
}
