//! Parse Markdown text into a `DocumentIR`.
//!
//! Handles the subset of Markdown that document extraction pipelines and
//! office_oxide's own `to_markdown()` produce: ATX headings, pipe tables,
//! bullet/numbered lists, thematic breaks, and paragraphs with bold/italic
//! inline spans.  This is not a full CommonMark implementation — it is
//! intentionally minimal so it carries no extra dependencies.

use crate::format::DocumentFormat;
use crate::ir::{
    DocumentIR, Element, Heading, InlineContent, List, ListItem, Metadata, Paragraph, Section,
    Table, TableCell, TableRow, TextSpan,
};

impl DocumentIR {
    /// Parse Markdown text into a `DocumentIR`.
    ///
    /// Sections are separated by `---` horizontal rules (common for page
    /// boundaries in extracted documents).  Each ATX heading that is not immediately inside
    /// a list or table also acts as a natural section boundary.
    ///
    /// # Example
    ///
    /// ```rust
    /// use office_oxide::ir::DocumentIR;
    /// use office_oxide::format::DocumentFormat;
    ///
    /// let md = "# Title\n\nHello **world**.\n\n- item one\n- item two\n";
    /// let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
    /// assert!(!ir.sections.is_empty());
    /// ```
    pub fn from_markdown(markdown: &str, format: DocumentFormat) -> Self {
        let mut parser = MarkdownParser::new(markdown);
        let sections = parser.parse_sections();
        DocumentIR {
            metadata: Metadata {
                format,
                title: None,
                ..Default::default()
            },
            sections,
        }
    }
}

// ---------------------------------------------------------------------------
// Parser state
// ---------------------------------------------------------------------------

struct MarkdownParser<'a> {
    lines: Vec<&'a str>,
    pos: usize,
}

impl<'a> MarkdownParser<'a> {
    fn new(src: &'a str) -> Self {
        Self {
            lines: src.lines().collect(),
            pos: 0,
        }
    }

    fn peek(&self) -> Option<&'a str> {
        self.lines.get(self.pos).copied()
    }

    fn advance(&mut self) -> Option<&'a str> {
        let line = self.lines.get(self.pos).copied();
        self.pos += 1;
        line
    }

    /// Parse the full document into sections, splitting on thematic `---`
    /// breaks and top-level H1/H2 headings.
    fn parse_sections(&mut self) -> Vec<Section> {
        let mut sections: Vec<Section> = Vec::new();
        let mut current = Section {
            title: None,
            elements: Vec::new(),
            ..Default::default()
        };

        while self.pos < self.lines.len() {
            let line = match self.peek() {
                Some(l) => l,
                None => break,
            };

            // Blank line
            if line.trim().is_empty() {
                self.advance();
                continue;
            }

            // Thematic break `---` / `***` / `___` starts a new section
            if is_thematic_break(line) {
                self.advance();
                if !current.elements.is_empty() || current.title.is_some() {
                    sections.push(current);
                    current = Section {
                        title: None,
                        elements: Vec::new(),
                        ..Default::default()
                    };
                }
                continue;
            }

            // ATX heading
            if let Some((level, text)) = parse_atx_heading(line) {
                self.advance();
                // H1 headings start a new section
                if level == 1 {
                    if !current.elements.is_empty() || current.title.is_some() {
                        sections.push(current);
                    }
                    current = Section {
                        title: Some(text.clone()),
                        elements: Vec::new(),
                        ..Default::default()
                    };
                } else {
                    current.elements.push(Element::Heading(Heading {
                        level,
                        content: parse_inline(&text),
                        ..Default::default()
                    }));
                }
                continue;
            }

            // Pipe table
            if line.trim_start().starts_with('|') {
                if let Some(table) = self.parse_table() {
                    current.elements.push(Element::Table(table));
                    continue;
                }
            }

            // Unordered list
            if is_unordered_list_marker(line) {
                let list = self.parse_list(false);
                current.elements.push(Element::List(list));
                continue;
            }

            // Ordered list
            if is_ordered_list_marker(line) {
                let list = self.parse_list(true);
                current.elements.push(Element::List(list));
                continue;
            }

            // Standalone image line: `![alt](data:image/...;base64,...)`.
            // Only data-URI sources carry embeddable bytes — external URLs
            // and local paths stay paragraph text.
            if let Some(image) = parse_image_element(line) {
                self.advance();
                current.elements.push(Element::Image(image));
                continue;
            }

            // Regular paragraph (accumulate until blank line or block element)
            let para = self.parse_paragraph();
            if !para.content.is_empty() {
                current.elements.push(Element::Paragraph(para));
            }
        }

        if !current.elements.is_empty() || current.title.is_some() {
            sections.push(current);
        }

        // Ensure at least one section
        if sections.is_empty() {
            sections.push(Section {
                title: None,
                elements: Vec::new(),
                ..Default::default()
            });
        }

        sections
    }

    // -----------------------------------------------------------------------
    // Block parsers
    // -----------------------------------------------------------------------

    fn parse_paragraph(&mut self) -> Paragraph {
        let mut lines: Vec<&str> = Vec::new();
        loop {
            match self.peek() {
                None => break,
                Some(line) => {
                    if line.trim().is_empty()
                        || parse_atx_heading(line).is_some()
                        || is_thematic_break(line)
                        || line.trim_start().starts_with('|')
                        || is_unordered_list_marker(line)
                        || is_ordered_list_marker(line)
                    {
                        break;
                    }
                    lines.push(line);
                    self.advance();
                },
            }
        }
        let text = lines.join(" ");
        Paragraph {
            content: parse_inline(&text),
            ..Default::default()
        }
    }

    fn parse_table(&mut self) -> Option<Table> {
        // Collect all consecutive pipe lines
        let mut raw: Vec<&'a str> = Vec::new();
        while let Some(line) = self.peek() {
            if line.trim_start().starts_with('|') {
                raw.push(line);
                self.advance();
            } else {
                break;
            }
        }

        if raw.is_empty() {
            return None;
        }

        // Filter out alignment rows (cells that look like `---`, `:---`, `---:`)
        let data_rows: Vec<&str> = raw
            .iter()
            .copied()
            .filter(|line| !is_table_separator_row(line))
            .collect();

        if data_rows.is_empty() {
            return None;
        }

        let mut rows: Vec<TableRow> = Vec::new();
        for (i, row_line) in data_rows.iter().enumerate() {
            let cells = split_pipe_row(row_line)
                .into_iter()
                .map(|cell_text| TableCell {
                    content: vec![Element::Paragraph(Paragraph {
                        content: parse_inline(cell_text.trim()),
                        ..Default::default()
                    })],
                    col_span: 1,
                    row_span: 1,
                    ..Default::default()
                })
                .collect();
            rows.push(TableRow {
                cells,
                is_header: i == 0,
                ..Default::default()
            });
        }

        Some(Table {
            rows,
            ..Default::default()
        })
    }

    fn parse_list(&mut self, ordered: bool) -> List {
        let mut items: Vec<ListItem> = Vec::new();
        loop {
            match self.peek() {
                None => break,
                Some(line) => {
                    if ordered && !is_ordered_list_marker(line) {
                        break;
                    }
                    if !ordered && !is_unordered_list_marker(line) {
                        break;
                    }
                    self.advance();
                    let content_str = strip_list_marker(line);
                    items.push(ListItem {
                        content: vec![Element::Paragraph(Paragraph {
                            content: parse_inline(content_str),
                            ..Default::default()
                        })],
                        nested: None,
                    });
                },
            }
        }
        List {
            ordered,
            items,
            ..Default::default()
        }
    }
}

// ---------------------------------------------------------------------------
// Inline parser (bold, italic, plain)
// ---------------------------------------------------------------------------

fn parse_inline(text: &str) -> Vec<InlineContent> {
    let mut out: Vec<InlineContent> = Vec::new();
    let bytes = text.as_bytes();
    let len = text.len();
    let mut plain_start = 0usize;

    macro_rules! flush_plain {
        ($end:expr) => {
            if plain_start < $end {
                let t = &text[plain_start..$end];
                if !t.is_empty() {
                    out.push(InlineContent::Text(TextSpan::plain(t)));
                }
            }
        };
    }

    let mut i = 0usize;
    while i < len {
        // Bold: **text** or __text__
        if i + 1 < len
            && ((bytes[i] == b'*' && bytes[i + 1] == b'*')
                || (bytes[i] == b'_' && bytes[i + 1] == b'_'))
        {
            let marker = &text[i..i + 2];
            if let Some(end) = text[i + 2..].find(marker) {
                flush_plain!(i);
                let inner = &text[i + 2..i + 2 + end];
                out.push(InlineContent::Text(TextSpan {
                    text: inner.to_string(),
                    bold: true,
                    ..Default::default()
                }));
                i += 2 + end + 2;
                plain_start = i;
                continue;
            }
        }

        // Italic: *text* or _text_
        if (bytes[i] == b'*' || bytes[i] == b'_') && i + 1 < len && bytes[i + 1] != bytes[i] {
            let marker = &text[i..i + 1];
            if let Some(end) = text[i + 1..].find(marker) {
                flush_plain!(i);
                let inner = &text[i + 1..i + 1 + end];
                out.push(InlineContent::Text(TextSpan {
                    text: inner.to_string(),
                    italic: true,
                    ..Default::default()
                }));
                i += 1 + end + 1;
                plain_start = i;
                continue;
            }
        }

        // Strikethrough: ~~text~~
        if i + 1 < len && bytes[i] == b'~' && bytes[i + 1] == b'~' {
            if let Some(end) = text[i + 2..].find("~~") {
                flush_plain!(i);
                let inner = &text[i + 2..i + 2 + end];
                out.push(InlineContent::Text(TextSpan {
                    text: inner.to_string(),
                    strikethrough: true,
                    ..Default::default()
                }));
                i += 2 + end + 2;
                plain_start = i;
                continue;
            }
        }

        // Inline code: `code` — strip backticks, treat as plain
        if bytes[i] == b'`' {
            if let Some(end) = text[i + 1..].find('`') {
                flush_plain!(i);
                let inner = &text[i + 1..i + 1 + end];
                out.push(InlineContent::Text(TextSpan::plain(inner)));
                i += 1 + end + 1;
                plain_start = i;
                continue;
            }
        }

        // Markdown link: [text](url)
        if bytes[i] == b'[' {
            if let Some(bracket_end) = text[i + 1..].find(']') {
                let after_bracket = i + 1 + bracket_end + 1;
                if after_bracket < len && bytes[after_bracket] == b'(' {
                    if let Some(paren_end) = text[after_bracket + 1..].find(')') {
                        flush_plain!(i);
                        let link_text = &text[i + 1..i + 1 + bracket_end];
                        let url = &text[after_bracket + 1..after_bracket + 1 + paren_end];
                        out.push(InlineContent::Text(TextSpan {
                            text: link_text.to_string(),
                            hyperlink: Some(url.to_string()),
                            ..Default::default()
                        }));
                        i = after_bracket + 1 + paren_end + 1;
                        plain_start = i;
                        continue;
                    }
                }
            }
        }

        i += text[i..].chars().next().map(|c| c.len_utf8()).unwrap_or(1);
    }

    flush_plain!(len);

    out
}

// ---------------------------------------------------------------------------
// Line classifiers
// ---------------------------------------------------------------------------

fn parse_atx_heading(line: &str) -> Option<(u8, String)> {
    let trimmed = line.trim_start();
    let hashes = trimmed.bytes().take_while(|&b| b == b'#').count();
    if hashes == 0 || hashes > 6 {
        return None;
    }
    let rest = &trimmed[hashes..];
    if rest.is_empty() || rest.starts_with(' ') || rest.starts_with('\t') {
        let text = rest.trim().trim_end_matches('#').trim().to_string();
        Some((hashes as u8, text))
    } else {
        None
    }
}

fn is_thematic_break(line: &str) -> bool {
    let t = line.trim();
    if t.len() < 3 {
        return false;
    }
    let Some(ch) = t.chars().next() else {
        return false;
    };
    if !matches!(ch, '-' | '*' | '_') {
        return false;
    }
    t.chars().all(|c| c == ch || c == ' ') && t.chars().filter(|&c| c == ch).count() >= 3
}

fn is_table_separator_row(line: &str) -> bool {
    let trimmed = line.trim().trim_matches('|');
    trimmed.split('|').all(|cell| {
        let c = cell.trim().trim_start_matches(':').trim_end_matches(':');
        !c.is_empty() && c.bytes().all(|b| b == b'-')
    })
}

fn split_pipe_row(line: &str) -> Vec<&str> {
    let inner = line.trim().trim_start_matches('|').trim_end_matches('|');
    inner.split('|').collect()
}

fn is_unordered_list_marker(line: &str) -> bool {
    let t = line.trim_start();
    (t.starts_with("- ") || t.starts_with("* ") || t.starts_with("+ ")) && !is_thematic_break(line)
}

fn is_ordered_list_marker(line: &str) -> bool {
    let t = line.trim_start();
    // e.g. "1. " "12. " "1) "
    let num_end = t.bytes().take_while(|b| b.is_ascii_digit()).count();
    if num_end == 0 {
        return false;
    }
    let after = &t[num_end..];
    after.starts_with(". ") || after.starts_with(") ")
}

fn strip_list_marker(line: &str) -> &str {
    let t = line.trim_start();
    if t.starts_with("- ") || t.starts_with("* ") || t.starts_with("+ ") {
        t[2..].trim_start()
    } else {
        // ordered: skip digits + ". " or ") "
        let num_end = t.bytes().take_while(|b| b.is_ascii_digit()).count();
        t[num_end + 2..].trim_start()
    }
}

// ---------------------------------------------------------------------------
// Image lines (`![alt](data:image/...;base64,...)`)
// ---------------------------------------------------------------------------

    /// Longest display width an embedded image may take: 6.0 in — fits the text
    /// column of Letter and A4 with default margins. Larger rasters are scaled
    /// down preserving aspect; smaller ones keep their natural 96-dpi size.
    const MAX_INLINE_IMAGE_WIDTH_EMU: u64 = 5_486_400;
    /// 96 dpi — the reference density for EMU conversion (`px * 9525`).
    const EMU_PER_PX: u64 = 9_525;

    /// Parse a standalone markdown image line into an IR image element.
    /// Returns `None` unless the whole line is one image whose source is a
    /// `data:image/...;base64,...` URI (the only source with embeddable bytes).
    fn parse_image_element(line: &str) -> Option<crate::ir::Image> {
        let trimmed = line.trim();
        let rest = trimmed.strip_prefix("![")?;
        let alt_end = rest.find("](")?;
        let alt = &rest[..alt_end];
        let src = rest[alt_end + 2..].strip_suffix(')')?;
        // The alt text may itself contain `](`-lookalikes only if the document
        // is pathological; the FIRST `](` closes the alt per CommonMark.
        if alt.contains(']') {
            return None;
        }

        let (format, data) = decode_data_image(src)?;
        let (pixel_w, pixel_h) = image_pixel_dimensions(&format, &data).unwrap_or((0, 0));
        let (display_w, display_h) = if pixel_w > 0 && pixel_h > 0 {
            let w = ((pixel_w as u64) * EMU_PER_PX).min(MAX_INLINE_IMAGE_WIDTH_EMU);
            let h = w * (pixel_h as u64) / (pixel_w as u64);
            (Some(w), Some(h))
        } else {
            (None, None)
        };

        Some(crate::ir::Image {
            alt_text: (!alt.is_empty()).then(|| alt.to_string()),
            data: Some(data),
            format: Some(format),
            display_width_emu: display_w,
            display_height_emu: display_h,
            pixel_width: (pixel_w > 0).then_some(pixel_w),
            pixel_height: (pixel_h > 0).then_some(pixel_h),
            ..Default::default()
        })
    }

    /// Decode a `data:image/<subtype>;base64,<payload>` URI into raw bytes plus
    /// the IR format we support embedding.
    fn decode_data_image(src: &str) -> Option<(crate::ir::ImageFormat, Vec<u8>)> {
        let rest = src.strip_prefix("data:image/")?;
        let (subtype, payload) = rest.split_once(';')?;
        let payload = payload.strip_prefix("base64,")?;
        let format = match subtype.to_ascii_lowercase().as_str() {
            "png" => crate::ir::ImageFormat::Png,
            "jpeg" | "jpg" => crate::ir::ImageFormat::Jpeg,
            _ => return None,
        };
        let data = base64_decode(payload)?;
        (!data.is_empty()).then(|| (format, data))
    }

    /// Pixel dimensions of a PNG or JPEG raster, read straight from the coded
    /// headers — no image-decoding dependency.
    fn image_pixel_dimensions(
        format: &crate::ir::ImageFormat,
        data: &[u8],
    ) -> Option<(u32, u32)> {
        match format {
            crate::ir::ImageFormat::Png => png_dimensions(data),
            crate::ir::ImageFormat::Jpeg => jpeg_dimensions(data),
            _ => None,
        }
    }

    /// PNG: 8-byte signature, then the IHDR chunk carries BE u32 width/height.
    fn png_dimensions(data: &[u8]) -> Option<(u32, u32)> {
        if data.len() < 24 || &data[..8] != b"\x89PNG\r\n\x1a\n" || &data[12..16] != b"IHDR" {
            return None;
        }
        let w = u32::from_be_bytes(data[16..20].try_into().ok()?);
        let h = u32::from_be_bytes(data[20..24].try_into().ok()?);
        (w > 0 && h > 0).then_some((w, h))
    }

    /// JPEG: walk the marker segments until a SOF0/1/2 frame header; its payload
    /// is precision(1) + height(2, BE) + width(2, BE).
    fn jpeg_dimensions(data: &[u8]) -> Option<(u32, u32)> {
        if data.len() < 4 || data[0] != 0xFF || data[1] != 0xD8 {
            return None;
        }
        let mut i = 2;
        while i + 4 <= data.len() {
            if data[i] != 0xFF {
                return None;
            }
            let marker = data[i + 1];
            i += 2;
            match marker {
                // Standalone markers carry no length segment.
                0xD8 | 0x01 | 0xD0..=0xD7 => continue,
                0xD9 => return None, // EOI before any frame header
                _ => {},
            }
            if i + 2 > data.len() {
                return None;
            }
            let seg_len = u16::from_be_bytes([data[i], data[i + 1]]) as usize;
            if seg_len < 2 || i + seg_len > data.len() {
                return None;
            }
            // SOF0 (baseline), SOF1 (extended), SOF2 (progressive).
            if (0xC0..=0xC2).contains(&marker) {
                if seg_len < 7 {
                    return None;
                }
                let h = u16::from_be_bytes([data[i + 3], data[i + 4]]) as u32;
                let w = u16::from_be_bytes([data[i + 5], data[i + 6]]) as u32;
                return (w > 0 && h > 0).then_some((w, h));
            }
            i += seg_len;
        }
        None
    }

    /// Standard base64 (RFC 4648, padded) → bytes. `None` on any invalid input.
    fn base64_decode(input: &str) -> Option<Vec<u8>> {
        fn value(b: u8) -> Option<u32> {
            match b {
                b'A'..=b'Z' => Some((b - b'A') as u32),
                b'a'..=b'z' => Some((b - b'a') as u32 + 26),
                b'0'..=b'9' => Some((b - b'0') as u32 + 52),
                b'+' => Some(62),
                b'/' => Some(63),
                _ => None,
            }
        }
        let cleaned: Vec<u8> = input
            .bytes()
            .filter(|b| !b.is_ascii_whitespace())
            .collect();
        let data_len = cleaned
            .iter()
            .position(|&b| b == b'=')
            .unwrap_or(cleaned.len());
        // Padding: at most two `=`, only at the very end.
        if cleaned.len() - data_len > 2 || cleaned[data_len..].iter().any(|&b| b != b'=') {
            return None;
        }
        let rem = data_len % 4;
        if rem == 1 {
            return None; // a 1-char quantum can never be valid base64
        }
        let mut out = Vec::with_capacity(data_len / 4 * 3 + 3);
        for chunk in cleaned[..data_len].chunks(4) {
            let mut n = chunk
                .iter()
                .try_fold(0u32, |acc, &b| value(b).map(|v| (acc << 6) | v))?;
            // Re-align partial quanta so the 24-bit window holds the bytes.
            n <<= 6 * (4 - chunk.len() as u32);
            match chunk.len() {
                4 => out.extend_from_slice(&[(n >> 16) as u8, (n >> 8) as u8, n as u8]),
                3 => out.extend_from_slice(&[(n >> 16) as u8, (n >> 8) as u8]),
                2 => out.push((n >> 16) as u8),
                _ => return None,
            }
        }
        Some(out)
    }

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use crate::format::DocumentFormat;

    #[test]
    fn parse_heading_paragraph() {
        let md = "# Hello\n\nSome text here.\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        assert_eq!(ir.sections.len(), 1);
        assert_eq!(ir.sections[0].title.as_deref(), Some("Hello"));
        assert!(matches!(ir.sections[0].elements[0], Element::Paragraph(_)));
    }

    #[test]
    fn parse_page_break_into_sections() {
        let md = "# Page 1\n\nText one.\n\n---\n\n# Page 2\n\nText two.\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        assert_eq!(ir.sections.len(), 2);
        assert_eq!(ir.sections[0].title.as_deref(), Some("Page 1"));
        assert_eq!(ir.sections[1].title.as_deref(), Some("Page 2"));
    }

    #[test]
    fn parse_unordered_list() {
        let md = "- apple\n- banana\n- cherry\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        let list = match &ir.sections[0].elements[0] {
            Element::List(l) => l,
            other => panic!("expected List, got {other:?}"),
        };
        assert!(!list.ordered);
        assert_eq!(list.items.len(), 3);
    }

    #[test]
    fn parse_ordered_list() {
        let md = "1. first\n2. second\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        let list = match &ir.sections[0].elements[0] {
            Element::List(l) => l,
            other => panic!("expected List, got {other:?}"),
        };
        assert!(list.ordered);
    }

    #[test]
    fn parse_pipe_table() {
        let md = "| Name | Age |\n|------|-----|\n| Alice | 30 |\n| Bob | 25 |\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        let table = match &ir.sections[0].elements[0] {
            Element::Table(t) => t,
            other => panic!("expected Table, got {other:?}"),
        };
        assert_eq!(table.rows.len(), 3); // header + 2 data rows (separator stripped)
        assert!(table.rows[0].is_header);
    }

    #[test]
    fn parse_bold_italic_inline() {
        let md = "Hello **world** and *rust*.\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        let para = match &ir.sections[0].elements[0] {
            Element::Paragraph(p) => p,
            other => panic!("expected Paragraph, got {other:?}"),
        };
        let spans: Vec<_> = para
            .content
            .iter()
            .filter_map(|c| match c {
                InlineContent::Text(s) => Some(s),
                _ => None,
            })
            .collect();
        assert!(spans.iter().any(|s| s.bold && s.text == "world"));
        assert!(spans.iter().any(|s| s.italic && s.text == "rust"));
    }

    #[test]
    fn parse_empty_markdown() {
        let ir = DocumentIR::from_markdown("", DocumentFormat::Docx);
        assert_eq!(ir.sections.len(), 1);
        assert!(ir.sections[0].elements.is_empty());
    }

    // 1×1 transparent PNG, base64 of the canonical 68-byte pixel.
    const TINY_PNG_B64: &str = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

    #[test]
    fn parse_standalone_data_uri_image() {
        let md = format!("# Report\n\n![Revenue chart](data:image/png;base64,{TINY_PNG_B64})\n\nAfter text.\n");
        let ir = DocumentIR::from_markdown(&md, DocumentFormat::Docx);
        assert_eq!(ir.sections[0].title.as_deref(), Some("Report"));
        let elems = &ir.sections[0].elements;
        let img = match &elems[0] {
            Element::Image(i) => i,
            other => panic!("expected Image, got {other:?}"),
        };
        assert_eq!(img.alt_text.as_deref(), Some("Revenue chart"));
        assert_eq!(img.format, Some(crate::ir::ImageFormat::Png));
        assert_eq!(img.pixel_width, Some(1));
        assert_eq!(img.pixel_height, Some(1));
        // 1 px → 9525 EMU display (below the 6-in cap).
        assert_eq!(img.display_width_emu, Some(9_525));
        assert_eq!(img.display_height_emu, Some(9_525));
        assert!(matches!(elems[1], Element::Paragraph(_)), "text after image");
    }

    #[test]
    fn parse_wide_image_fits_six_inch_column() {
        // 2000×1000 px → natural 19_050_000 EMU wide, capped at 5_486_400.
        let mut png = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR".to_vec();
        png.extend_from_slice(&2000u32.to_be_bytes());
        png.extend_from_slice(&1000u32.to_be_bytes());
        png.extend_from_slice(&[8, 6, 0, 0, 0]);
        let b64 = {
            const T: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
            let mut out = String::new();
            for chunk in png.chunks(3) {
                let n = ((chunk[0] as u32) << 16)
                    | ((chunk.get(1).copied().unwrap_or(0) as u32) << 8)
                    | chunk.get(2).copied().unwrap_or(0) as u32;
                out.push(T[(n >> 18) as usize & 63] as char);
                out.push(T[(n >> 12) as usize & 63] as char);
                if chunk.len() > 1 {
                    out.push(T[(n >> 6) as usize & 63] as char);
                }
                if chunk.len() > 2 {
                    out.push(T[n as usize & 63] as char);
                }
            }
            while out.len() % 4 != 0 {
                out.push('=');
            }
            out
        };
        let md = format!("![t](data:image/png;base64,{b64})\n");
        let ir = DocumentIR::from_markdown(&md, DocumentFormat::Docx);
        let img = match &ir.sections[0].elements[0] {
            Element::Image(i) => i,
            other => panic!("expected Image, got {other:?}"),
        };
        assert_eq!(img.display_width_emu, Some(5_486_400));
        assert_eq!(img.display_height_emu, Some(5_486_400 / 2));
    }

    #[test]
    fn external_image_source_stays_text() {
        let md = "![logo](https://example.com/x.png)\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        assert!(matches!(
            ir.sections[0].elements[0],
            Element::Paragraph(_)
        ));
    }

    #[test]
    fn jpeg_data_uri_image_parses() {
        // Minimal JPEG header: SOI + SOF0 (17-byte segment, 8×4 px).
        let mut jpg = vec![0xFF, 0xD8, 0xFF, 0xC0, 0x00, 0x11, 0x08, 0x00, 0x04, 0x00, 0x08, 0x03, 0x01, 0x22, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01, 0xFF, 0xD9];
        // Ensure the payload isn't empty after the marker walk.
        jpg.truncate(jpg.len());
        let b64 = {
            const T: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
            let mut out = String::new();
            for chunk in jpg.chunks(3) {
                let n = ((chunk[0] as u32) << 16)
                    | ((chunk.get(1).copied().unwrap_or(0) as u32) << 8)
                    | chunk.get(2).copied().unwrap_or(0) as u32;
                out.push(T[(n >> 18) as usize & 63] as char);
                out.push(T[(n >> 12) as usize & 63] as char);
                if chunk.len() > 1 {
                    out.push(T[(n >> 6) as usize & 63] as char);
                }
                if chunk.len() > 2 {
                    out.push(T[n as usize & 63] as char);
                }
            }
            while out.len() % 4 != 0 {
                out.push('=');
            }
            out
        };
        let md = format!("![photo](data:image/jpeg;base64,{b64})\n");
        let ir = DocumentIR::from_markdown(&md, DocumentFormat::Docx);
        let img = match &ir.sections[0].elements[0] {
            Element::Image(i) => i,
            other => panic!("expected Image, got {other:?}"),
        };
        assert_eq!(img.format, Some(crate::ir::ImageFormat::Jpeg));
        assert_eq!(img.pixel_width, Some(8));
        assert_eq!(img.pixel_height, Some(4));
    }

    #[test]
    fn base64_decode_rejects_garbage() {
        assert_eq!(base64_decode("aGVsbG8="), Some(b"hello".to_vec()));
        assert_eq!(base64_decode("aGVsbG8="), Some(b"hello".to_vec()));
        assert_eq!(base64_decode(""), Some(Vec::new()));
        assert_eq!(base64_decode("A"), None);
        assert_eq!(base64_decode("aGVsbG8*"), None);
        assert_eq!(base64_decode("a=b="), None);
    }
}
