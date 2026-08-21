//! Unified document editing API.
//!
//! Provides a read-modify-write workflow that preserves all unmodified
//! parts (images, charts, custom XML, etc.) through the `EditablePackage`.

use std::io::{Read, Seek, Write};
use std::path::Path;

use crate::Result;
use crate::format::DocumentFormat;

// ── edit operation specifications ────────────────────────────────────────────
//
// These describe declarative edits the caller wants applied. They are produced
// from a higher-level tool contract (e.g. kawai's office_edit_document) and
// kept here so the edit logic is self-contained.

/// A text run inside a DOCX paragraph/heading/list block.
#[derive(Debug, Clone)]
pub struct DocxRun {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    /// Font size in points.
    pub size: Option<f64>,
    /// Hex color, e.g. "FF0000".
    pub color: Option<String>,
    /// Font family name.
    pub font: Option<String>,
}

impl DocxRun {
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            bold: false,
            italic: false,
            size: None,
            color: None,
            font: None,
        }
    }
}

/// The kind of DOCX block to append.
#[derive(Debug, Clone, Copy)]
pub enum DocxBlockKind {
    Paragraph,
    Title,
    Heading1,
    Heading2,
    Heading3,
    Bullet,
}

/// A DOCX block: a heading/paragraph/list with one or more runs.
#[derive(Debug, Clone)]
pub struct DocxBlock {
    pub kind: DocxBlockKind,
    pub runs: Vec<DocxRun>,
}

/// Horizontal alignment for a DOCX paragraph.
#[derive(Debug, Clone, Copy)]
pub enum DocxAlign {
    Left,
    Center,
    Right,
    Justify,
}

/// Formatting to apply to a DOCX paragraph matched by text.
#[derive(Debug, Clone, Default)]
pub struct DocxFormat {
    pub alignment: Option<DocxAlign>,
    /// Spacing before the paragraph, in twips.
    pub spacing_before: Option<u32>,
    /// Spacing after the paragraph, in twips.
    pub spacing_after: Option<u32>,
    /// Left indent, in twips.
    pub indent_left: Option<u32>,
    /// Right indent, in twips.
    pub indent_right: Option<u32>,
}

/// A single spreadsheet cell value (type-inferred on write).
#[derive(Debug, Clone)]
pub enum XlsxCellValue {
    String(String),
    Number(f64),
    Boolean(bool),
}

/// A slide to append to a PPTX deck.
#[derive(Debug, Clone)]
pub struct PptxSlideSpec {
    pub title: String,
    pub body: Vec<String>,
}

/// An editable document that supports text replacement and saving.
///
/// The document is loaded in its entirety (all OPC parts) so that
/// unmodified parts are preserved verbatim on save.
pub struct EditableDocument {
    inner: EditableInner,
}

enum EditableInner {
    Docx(crate::docx::edit::EditableDocx),
    Xlsx(crate::xlsx::edit::EditableXlsx),
    Pptx(crate::pptx::edit::EditablePptx),
}

impl EditableDocument {
    /// Open a document for editing. Format is detected from the file extension.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref();
        let format = DocumentFormat::from_path(path).ok_or_else(|| {
            crate::OfficeError::UnsupportedFormat(
                path.extension()
                    .and_then(|e| e.to_str())
                    .unwrap_or("(none)")
                    .to_string(),
            )
        })?;
        match format {
            DocumentFormat::Docx => {
                let doc = crate::docx::edit::EditableDocx::open(path)?;
                Ok(Self {
                    inner: EditableInner::Docx(doc),
                })
            },
            DocumentFormat::Xlsx => {
                let doc = crate::xlsx::edit::EditableXlsx::open(path)?;
                Ok(Self {
                    inner: EditableInner::Xlsx(doc),
                })
            },
            DocumentFormat::Pptx => {
                let doc = crate::pptx::edit::EditablePptx::open(path)?;
                Ok(Self {
                    inner: EditableInner::Pptx(doc),
                })
            },
            _ => Err(crate::OfficeError::UnsupportedFormat(format!("{format:?}"))),
        }
    }

    /// Open a document for editing from a reader with explicit format.
    pub fn from_reader<R: Read + Seek>(reader: R, format: DocumentFormat) -> Result<Self> {
        match format {
            DocumentFormat::Docx => {
                let doc = crate::docx::edit::EditableDocx::from_reader(reader)?;
                Ok(Self {
                    inner: EditableInner::Docx(doc),
                })
            },
            DocumentFormat::Xlsx => {
                let doc = crate::xlsx::edit::EditableXlsx::from_reader(reader)?;
                Ok(Self {
                    inner: EditableInner::Xlsx(doc),
                })
            },
            DocumentFormat::Pptx => {
                let doc = crate::pptx::edit::EditablePptx::from_reader(reader)?;
                Ok(Self {
                    inner: EditableInner::Pptx(doc),
                })
            },
            _ => Err(crate::OfficeError::UnsupportedFormat(format!("{format:?}"))),
        }
    }

    /// Replace all occurrences of `find` with `replace` in text content.
    /// Returns the number of replacements made.
    ///
    /// For DOCX: replaces text in `<w:t>` elements.
    /// For PPTX: replaces text in `<a:t>` elements across all slides.
    /// For XLSX: not applicable (use `set_cell` instead) — returns 0.
    pub fn replace_text(&mut self, find: &str, replace: &str) -> usize {
        match &mut self.inner {
            EditableInner::Docx(doc) => doc.replace_text(find, replace),
            EditableInner::Pptx(doc) => doc.replace_text(find, replace),
            EditableInner::Xlsx(_) => 0,
        }
    }

    /// Set a cell value in an XLSX document.
    /// Returns an error if this is not an XLSX document.
    pub fn set_cell(
        &mut self,
        sheet_index: usize,
        cell_ref: &str,
        value: crate::xlsx::edit::CellValue,
    ) -> Result<()> {
        match &mut self.inner {
            EditableInner::Xlsx(doc) => {
                doc.set_cell(sheet_index, cell_ref, value)?;
                Ok(())
            },
            _ => Err(crate::OfficeError::UnsupportedFormat(
                "set_cell is only supported for XLSX".to_string(),
            )),
        }
    }

    /// Append DOCX blocks (paragraphs/headings/lists) before `</w:body>`.
    pub fn append_docx_blocks(&mut self, blocks: &[DocxBlock]) -> Result<usize> {
        match &mut self.inner {
            EditableInner::Docx(doc) => doc.append_blocks(blocks),
            _ => Err(crate::OfficeError::UnsupportedFormat(
                "append_docx_blocks is only supported for DOCX".to_string(),
            )),
        }
    }

    /// Append a DOCX table before `</w:body>`.
    pub fn append_docx_table(&mut self, rows: &[Vec<String>]) -> Result<usize> {
        match &mut self.inner {
            EditableInner::Docx(doc) => doc.append_table(rows),
            _ => Err(crate::OfficeError::UnsupportedFormat(
                "append_docx_table is only supported for DOCX".to_string(),
            )),
        }
    }

    /// Delete DOCX paragraphs whose text contains `find`.
    pub fn delete_docx_paragraphs(&mut self, find: &str) -> Result<usize> {
        match &mut self.inner {
            EditableInner::Docx(doc) => doc.delete_paragraphs(find),
            _ => Err(crate::OfficeError::UnsupportedFormat(
                "delete_docx_paragraphs is only supported for DOCX".to_string(),
            )),
        }
    }

    /// Format the first DOCX paragraph whose text contains `find`.
    pub fn format_docx_paragraph(&mut self, find: &str, fmt: &DocxFormat) -> Result<usize> {
        match &mut self.inner {
            EditableInner::Docx(doc) => doc.format_paragraph(find, fmt),
            _ => Err(crate::OfficeError::UnsupportedFormat(
                "format_docx_paragraph is only supported for DOCX".to_string(),
            )),
        }
    }

    /// Replace text across all XLSX cells (shared strings + inline strings).
    /// Returns the number of replacements made.
    pub fn xlsx_replace_text(&mut self, find: &str, replace: &str) -> Result<usize> {
        match &mut self.inner {
            EditableInner::Xlsx(doc) => Ok(doc.replace_text(find, replace)?),
            _ => Err(crate::OfficeError::UnsupportedFormat(
                "xlsx_replace_text is only supported for XLSX".to_string(),
            )),
        }
    }

    /// Append rows to an XLSX worksheet. `sheet_index` is 0-based.
    /// Returns the number of rows appended.
    pub fn append_xlsx_rows(
        &mut self,
        sheet_index: usize,
        rows: &[Vec<XlsxCellValue>],
    ) -> Result<usize> {
        match &mut self.inner {
            EditableInner::Xlsx(doc) => Ok(doc.append_rows(sheet_index, rows)?),
            _ => Err(crate::OfficeError::UnsupportedFormat(
                "append_xlsx_rows is only supported for XLSX".to_string(),
            )),
        }
    }

    /// Append slides to a PPTX deck. Returns the number of slides appended.
    pub fn append_pptx_slides(&mut self, slides: &[PptxSlideSpec]) -> Result<usize> {
        match &mut self.inner {
            EditableInner::Pptx(doc) => Ok(doc.append_slides(slides)?),
            _ => Err(crate::OfficeError::UnsupportedFormat(
                "append_pptx_slides is only supported for PPTX".to_string(),
            )),
        }
    }

    /// Remove the first PPTX slide whose text contains `find`.
    /// Returns 1 if a slide was removed, 0 otherwise.
    pub fn remove_pptx_slide(&mut self, find: &str) -> Result<usize> {
        match &mut self.inner {
            EditableInner::Pptx(doc) => Ok(doc.remove_slide(find)?),
            _ => Err(crate::OfficeError::UnsupportedFormat(
                "remove_pptx_slide is only supported for PPTX".to_string(),
            )),
        }
    }

    /// Save the edited document to a file.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        match &self.inner {
            EditableInner::Docx(doc) => doc.save(path)?,
            EditableInner::Xlsx(doc) => doc.save(path)?,
            EditableInner::Pptx(doc) => doc.save(path)?,
        }
        Ok(())
    }

    /// Write the edited document to any `Write + Seek` destination.
    pub fn write_to<W: Write + Seek>(&self, writer: W) -> Result<()> {
        match &self.inner {
            EditableInner::Docx(doc) => doc.write_to(writer)?,
            EditableInner::Xlsx(doc) => doc.write_to(writer)?,
            EditableInner::Pptx(doc) => doc.write_to(writer)?,
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use super::*;

    fn make_docx_bytes() -> Vec<u8> {
        let mut w = crate::docx::write::DocxWriter::new();
        w.add_paragraph("Hello world");
        let mut buf = Cursor::new(Vec::new());
        w.write_to(&mut buf).unwrap();
        buf.into_inner()
    }

    fn make_xlsx_bytes() -> Vec<u8> {
        let mut w = crate::xlsx::write::XlsxWriter::new();
        {
            let mut sheet = w.add_sheet("Sheet1");
            sheet.set_cell(0, 0, crate::xlsx::write::CellData::String("value".into()));
        }
        let mut buf = Cursor::new(Vec::new());
        w.write_to(&mut buf).unwrap();
        buf.into_inner()
    }

    fn make_pptx_bytes() -> Vec<u8> {
        let mut w = crate::pptx::write::PptxWriter::new();
        let slide = w.add_slide();
        slide.set_title("Slide 1");
        let mut buf = Cursor::new(Vec::new());
        w.write_to(&mut buf).unwrap();
        buf.into_inner()
    }

    #[test]
    fn docx_replace_text_roundtrip() {
        let data = make_docx_bytes();
        let mut doc =
            EditableDocument::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();
        let count = doc.replace_text("Hello", "Hi");
        assert!(count >= 1);
        let mut out = Cursor::new(Vec::new());
        doc.write_to(&mut out).unwrap();
        assert!(!out.into_inner().is_empty());
    }

    #[test]
    fn xlsx_set_cell_roundtrip() {
        let data = make_xlsx_bytes();
        let mut doc =
            EditableDocument::from_reader(Cursor::new(data), DocumentFormat::Xlsx).unwrap();
        doc.set_cell(0, "A1", crate::xlsx::edit::CellValue::String("updated".into()))
            .unwrap();
        let mut out = Cursor::new(Vec::new());
        doc.write_to(&mut out).unwrap();
        assert!(!out.into_inner().is_empty());
    }

    #[test]
    fn pptx_replace_text_roundtrip() {
        let data = make_pptx_bytes();
        let mut doc =
            EditableDocument::from_reader(Cursor::new(data), DocumentFormat::Pptx).unwrap();
        doc.replace_text("Slide 1", "Updated");
        let mut out = Cursor::new(Vec::new());
        doc.write_to(&mut out).unwrap();
        assert!(!out.into_inner().is_empty());
    }

    #[test]
    fn xlsx_replace_text_returns_zero() {
        let data = make_xlsx_bytes();
        let mut doc =
            EditableDocument::from_reader(Cursor::new(data), DocumentFormat::Xlsx).unwrap();
        assert_eq!(doc.replace_text("anything", "other"), 0);
    }

    #[test]
    fn set_cell_on_docx_returns_error() {
        let data = make_docx_bytes();
        let mut doc =
            EditableDocument::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();
        assert!(
            doc.set_cell(0, "A1", crate::xlsx::edit::CellValue::String("x".into()))
                .is_err()
        );
    }

    #[test]
    fn from_reader_unsupported_format_returns_error() {
        let data = vec![0u8; 16];
        assert!(EditableDocument::from_reader(Cursor::new(data), DocumentFormat::Doc).is_err());
    }
}
