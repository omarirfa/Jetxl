use crate::types::{CellValue, SheetData, WriteError};
use crate::styles::*;
use arrow_array::{Array, RecordBatch};
use arrow_schema::DataType;
use chrono::Timelike;
use std::collections::HashMap;

/// Official OOXML CT_Worksheet element order from the schema
// const WORKSHEET_ELEMENT_ORDER: &[&str] = &[
//     "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols",
//     "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges",
//     "scenarios", "autoFilter", "sortState", "mergeCells", "phoneticPr",
//     "conditionalFormatting", "dataValidations", "hyperlinks",
//     "printOptions", "pageMargins", "pageSetup", "headerFooter",
//     "rowBreaks", "colBreaks", "customSheetViews", "mergeCells",
//     "phoneticPr", "conditionalFormatting", "dataValidations",
//     "hyperlinks", "printOptions", "pageMargins", "pageSetup",
//     "headerFooter", "rowBreaks", "colBreaks", "customSheetViews",
//     "controls", "customProperties", "cellWatches", "ignoredErrors",
//     "smartTags", "drawing", "drawingHF", "picture", "oleObjects",
//     "activeXControls", "webPublishItems", "tableParts", "extLst"
// ];

pub fn generate_app_xml(sheet_names: &[&str]) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" \
xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\
<Application>Microsoft Excel</Application>\
<DocSecurity>0</DocSecurity>\
<ScaleCrop>false</ScaleCrop>\
<HeadingPairs><vt:vector size=\"2\" baseType=\"variant\">\
<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>\
<vt:variant><vt:i4>{}</vt:i4></vt:variant>\
</vt:vector></HeadingPairs>\
<TitlesOfParts><vt:vector size=\"{}\" baseType=\"lpstr\">{}</vt:vector></TitlesOfParts>\
<LinksUpToDate>false</LinksUpToDate>\
<SharedDoc>false</SharedDoc>\
<AppVersion>16.0300</AppVersion>\
</Properties>",
        sheet_names.len(),
        sheet_names.len(),
        sheet_names.iter().map(|n| format!("<vt:lpstr>{}</vt:lpstr>", xml_escape_str(n))).collect::<Vec<_>>().join("")
    )
}

pub fn generate_core_xml() -> &'static str {
    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" \
xmlns:dc=\"http://purl.org/dc/elements/1.1/\" \
xmlns:dcterms=\"http://purl.org/dc/terms/\" \
xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\
<dc:creator>jetxl</dc:creator>\
<cp:lastModifiedBy>jetxl</cp:lastModifiedBy>\
<dcterms:created xsi:type=\"dcterms:W3CDTF\">2020-01-01T00:00:00Z</dcterms:created>\
<dcterms:modified xsi:type=\"dcterms:W3CDTF\">2020-01-01T00:00:00Z</dcterms:modified>\
</cp:coreProperties>"
}

/// Zero-allocation column letter writing - returns length written
#[inline(always)]
pub fn write_col_letter(col: usize, buf: &mut [u8; 4]) -> usize {
    // Excel's maximum column is XFD (index 16383). Anything larger cannot be a
    // valid worksheet column and would also risk overrunning the 4-byte buffer
    // (XFD is 3 letters; the 4th slot is headroom). Assert in debug builds so
    // callers catch bad indices in tests; release builds stay branch-free.
    debug_assert!(col <= 16383, "column index {} exceeds Excel maximum (XFD=16383)", col);

    if col < 26 {
        buf[0] = b'A' + col as u8;
        return 1;
    }
    
    let mut col = col;
    let mut len = 0;
    let mut stack = [0u8; 4];
    let mut stack_len = 0;
    
    while col >= 26 {
        stack[stack_len] = b'A' + (col % 26) as u8;
        stack_len += 1;
        col = col / 26 - 1;
    }
    stack[stack_len] = b'A' + col as u8;
    stack_len += 1;
    
    for i in 0..stack_len {
        buf[i] = stack[stack_len - 1 - i];
        len += 1;
    }
    
    len
}

/// Write cell reference (e.g. "A1", "B2") to buffer
#[inline(always)]
fn write_cell_ref(col: usize, row: usize, buf: &mut Vec<u8>) {
    let mut col_buf = [0u8; 4];
    let col_len = write_col_letter(col, &mut col_buf);
    buf.extend_from_slice(&col_buf[..col_len]);
    buf.extend_from_slice(itoa::Buffer::new().format(row).as_bytes());
}

/// Write a `TopLeft:BottomRight` range with start/end normalized so the
/// top-left is always min and bottom-right always max. A caller who passes a
/// reversed range (e.g. end_row < start_row) previously produced an invalid
/// sqref like "B20:B2", which Excel and openpyxl reject -- silently discarding
/// the conditional-format or data-validation rule. Sorting the corners makes
/// any rectangle the caller supplies a valid A1 range.
fn write_normalized_range(start_col: usize, start_row: usize, end_col: usize, end_row: usize, buf: &mut Vec<u8>) {
    let (c0, c1) = if start_col <= end_col { (start_col, end_col) } else { (end_col, start_col) };
    let (r0, r1) = if start_row <= end_row { (start_row, end_row) } else { (end_row, start_row) };
    write_cell_ref(c0, r0, buf);
    buf.push(b':');
    write_cell_ref(c1, r1, buf);
}

#[inline(always)]
fn datetime_to_excel_serial(dt: &chrono::NaiveDateTime) -> f64 {
    // Excel's 1900 date system deliberately includes a non-existent 1900-02-29
    // (serial 60) for Lotus 1-2-3 compatibility, so serial numbers must account
    // for that phantom day:
    //     serial 1  = 1900-01-01
    //     serial 59 = 1900-02-28
    //     serial 60 = 1900-02-29  (does not exist; never emitted)
    //     serial 61 = 1900-03-01
    //
    // Compute plain elapsed days from 1899-12-31 (so 1900-01-01 == 1), then add
    // one day for any date on/after 1900-03-01 to jump the phantom leap day.
    // A single epoch of 1899-12-30 (the old approach) bakes that +1 in for ALL
    // dates, which is correct for modern dates but shifts every date on/before
    // 1900-02-28 one day too high. Real-world data is >= 1900-03-01 so this only
    // affected historical dates, but it was still wrong.
    let date = dt.date();
    let base_epoch = chrono::NaiveDate::from_ymd_opt(1899, 12, 31).unwrap();
    let phantom_cutoff = chrono::NaiveDate::from_ymd_opt(1900, 3, 1).unwrap();
    let mut days = (date - base_epoch).num_days();
    if date >= phantom_cutoff {
        days += 1;
    }
    let time_fraction = (dt.hour() * 3600 + dt.minute() * 60 + dt.second()) as f64 / 86400.0;
    days as f64 + time_fraction
}

/// SIMD-accelerated XML escaping for cell/element text.
///
/// Correctness contract (why this is more than "replace 5 chars"):
///   1. The five XML metacharacters `& < > " \'` are escaped so user data can
///      never break out of an attribute or element.
///   2. Bytes that are *illegal* in XML 1.0 regardless of escaping -- the C0
///      control range except TAB (0x09), LF (0x0A) and CR (0x0D) -- are dropped.
///      Excel writing such a byte verbatim into `<t>` yields a file every
///      conformant reader rejects ("unreadable content"). These bytes have no
///      valid XML representation at all, so stripping them is the only
///      lossless-as-possible option and matches xlsxwriter's behaviour.
///
/// Performance: the common case (no metacharacter, no control byte) still hits
/// a single memchr-guided fast path and copies the whole slice in one shot, so
/// the per-cell hot loop is unaffected for typical text.
#[inline(always)]
pub fn xml_escape_simd(input: &[u8], output: &mut Vec<u8>) {
    // Fast-path detection. Two things force the slow byte-by-byte path: the five
    // XML metacharacters, and any illegal C0 control byte (0x00..=0x1F except
    // TAB/LF/CR), which has no valid XML representation at all.
    //
    // The metacharacter test uses `memchr`, whose SIMD scan is the fastest way to
    // check for those bytes -- this is the original hot-path check, unchanged.
    // Only when NO metacharacter is present (the overwhelmingly common case for
    // numeric/plain-text cells) do we run the control-byte scan, so we never do
    // more work than the original on clean metachar-bearing data, and clean plain
    // data pays one extra linear scan that the compiler autovectorizes. This is
    // what makes the escaper correct about control bytes without a hot-path
    // regression -- and keeps `memchr` doing what it is good at.
    let has_meta = memchr::memchr3(b'&', b'<', b'>', input).is_some()
        || memchr::memchr2(b'"', b'\'', input).is_some();

    if !has_meta {
        // No metacharacters: bulk-copy unless an illegal control byte is present.
        if !input.iter().any(|&b| b < 0x20 && b != b'\t' && b != b'\n' && b != b'\r') {
            output.extend_from_slice(input);
            return;
        }
    }

    let mut last = 0;
    let mut pos = 0;

    while pos < input.len() {
        let byte = input[pos];
        let escape: &[u8] = match byte {
            b'&' => b"&amp;",
            b'<' => b"&lt;",
            b'>' => b"&gt;",
            b'"' => b"&quot;",
            b'\'' => b"&apos;",
            // Illegal control byte (excluding TAB/LF/CR): flush the pending run
            // and skip this byte entirely -- it has no valid XML form.
            b if b < 0x20 && b != b'\t' && b != b'\n' && b != b'\r' => {
                output.extend_from_slice(&input[last..pos]);
                pos += 1;
                last = pos;
                continue;
            }
            _ => {
                pos += 1;
                continue;
            }
        };

        output.extend_from_slice(&input[last..pos]);
        output.extend_from_slice(escape);
        pos += 1;
        last = pos;
    }

    if last < input.len() {
        output.extend_from_slice(&input[last..]);
    }
}

/// Escape a `&str` for the *cold* metadata paths (sheet names, chart titles,
/// table names, series names, etc.) that assemble a `String` rather than a byte
/// buffer. These run a handful of times per file -- never per cell -- so the
/// small scratch allocation is irrelevant. Routing them through the same
/// escaper closes the injection/corruption holes where user-controlled metadata
/// was previously concatenated raw into the XML.
#[inline]
pub fn xml_escape_str(input: &str) -> String {
    let mut out = Vec::with_capacity(input.len() + 8);
    xml_escape_simd(input.as_bytes(), &mut out);
    // The escaper only emits ASCII escapes, copies existing UTF-8 runs, or drops
    // whole ASCII control bytes -- it never splits a multibyte sequence -- so the
    // result is always valid UTF-8.
    String::from_utf8(out).unwrap_or_default()
}

/// Sanitize a string into a valid OOXML table identifier for the `name` /
/// `displayName` attributes.
///
/// The spreadsheet spec requires these to be defined-name identifiers: they may
/// NOT contain spaces, must not start with a digit, and cannot collide with a
/// cell reference. jetxl previously wrote the caller's string verbatim, so a
/// perfectly reasonable (and README-documented) value like `"My Data"` produced
/// `displayName="My Data"` -- which Excel and openpyxl both reject, making the
/// whole workbook fail to open. We conservatively map any character that isn't a
/// letter, digit, or underscore to `_`, and prefix `_` if the result is empty or
/// starts with a digit. Purely-ASCII identifiers (the common case) that are
/// already valid pass through unchanged, so this doesn't alter existing good
/// output.
pub fn sanitize_table_identifier(input: &str) -> String {
    let mut out = String::with_capacity(input.len());
    for ch in input.chars() {
        if ch == '_' || ch.is_ascii_alphanumeric() || (ch.is_alphabetic() && !ch.is_ascii()) {
            out.push(ch);
        } else {
            out.push('_');
        }
    }
    if out.is_empty() {
        return "_".to_string();
    }
    // Must not start with a digit.
    if out.chars().next().map(|c| c.is_ascii_digit()).unwrap_or(false) {
        out.insert(0, '_');
    }
    out
}

/// Write an inline-string cell payload: `<is><t ...>escaped</t></is>`.
///
/// Emits `xml:space="preserve"` *only* when the value begins or ends with
/// whitespace. Without it, conformant readers (and Excel on reload) collapse or
/// trim leading/trailing spaces, silently corrupting values like `" 007"` or
/// `"code "`. Making the attribute conditional keeps the common case -- text with
/// no edge whitespace -- byte-for-byte identical to before, so the per-cell hot
/// path pays nothing extra for the vast majority of cells.
#[inline(always)]
fn write_inline_string(text: &[u8], buf: &mut Vec<u8>) {
    let needs_preserve = matches!(text.first(), Some(b) if b.is_ascii_whitespace())
        || matches!(text.last(), Some(b) if b.is_ascii_whitespace());
    if needs_preserve {
        buf.extend_from_slice(b"<is><t xml:space=\"preserve\">");
    } else {
        buf.extend_from_slice(b"<is><t>");
    }
    xml_escape_simd(text, buf);
    buf.extend_from_slice(b"</t></is>");
}

#[allow(dead_code)]
pub fn generate_content_types(sheet_names: &[&str], tables_per_sheet: &[usize]) -> String {
    let total_tables: usize = tables_per_sheet.iter().sum();
    let mut xml = String::with_capacity(800 + sheet_names.len() * 150 + total_tables * 100);
    xml.push_str(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
<Default Extension=\"xml\" ContentType=\"application/xml\"/>\
<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\
<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\
<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\
<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>",
    );

    for i in 1..=sheet_names.len() {
        xml.push_str("<Override PartName=\"/xl/worksheets/sheet");
        xml.push_str(&i.to_string());
        xml.push_str(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
    }

    // Add table content types
    let mut table_id = 1;
    for &table_count in tables_per_sheet { 
        for _ in 0..table_count {
            xml.push_str("<Override PartName=\"/xl/tables/table");
            xml.push_str(&table_id.to_string());
            xml.push_str(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>");
            table_id += 1;
        }
    }

    xml.push_str("</Types>");
    xml
}

#[allow(dead_code)]  // superseded by _ext variant
pub fn generate_content_types_with_charts(
    sheet_names: &[&str], 
    tables_per_sheet: &[usize], 
    charts_per_sheet: &[usize],
    images_per_sheet: &[(&[ExcelImage], usize)]
) -> String {
    generate_content_types_with_charts_ext(sheet_names, tables_per_sheet, charts_per_sheet, images_per_sheet, false)
}

pub fn generate_content_types_with_charts_ext(
    sheet_names: &[&str], 
    tables_per_sheet: &[usize], 
    charts_per_sheet: &[usize],
    images_per_sheet: &[(&[ExcelImage], usize)],
    has_shared_strings: bool,
) -> String {
    let total_tables: usize = tables_per_sheet.iter().sum();
    let total_charts: usize = charts_per_sheet.iter().sum();
    
    // Collect unique image extensions
    let mut image_extensions = std::collections::HashSet::new();
    for (images, _) in images_per_sheet {
        for img in *images {
            image_extensions.insert(img.extension.as_str());
        }
    }
    
    let mut xml = String::with_capacity(1000 + sheet_names.len() * 150 + total_tables * 100 + total_charts * 100 + image_extensions.len() * 100);
    
    xml.push_str(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
<Default Extension=\"xml\" ContentType=\"application/xml\"/>",
    );
    
    // Add image extensions
    for ext in &image_extensions {
        let content_type = match *ext {
            "png" => "image/png",
            "jpg" | "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "tiff" | "tif" => "image/tiff",
            _ => "application/octet-stream",
        };
        xml.push_str(&format!("<Default Extension=\"{}\" ContentType=\"{}\"/>", ext, content_type));
    }
    
    xml.push_str(
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\
<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\
<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\
<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>",
    );

    for i in 1..=sheet_names.len() {
        xml.push_str("<Override PartName=\"/xl/worksheets/sheet");
        xml.push_str(&i.to_string());
        xml.push_str(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
    }

    let mut table_id = 1;
    for &table_count in tables_per_sheet {
        for _ in 0..table_count {
            xml.push_str("<Override PartName=\"/xl/tables/table");
            xml.push_str(&table_id.to_string());
            xml.push_str(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>");
            table_id += 1;
        }
    }
    
    let mut chart_id = 1;
    for &chart_count in charts_per_sheet {
        for _ in 0..chart_count {
            xml.push_str("<Override PartName=\"/xl/charts/chart");
            xml.push_str(&chart_id.to_string());
            xml.push_str(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/>");
            chart_id += 1;
        }
    }
    
    let mut drawing_id = 1;
    for &(_, drawing_count) in images_per_sheet {
        if drawing_count > 0 {
            xml.push_str("<Override PartName=\"/xl/drawings/drawing");
            xml.push_str(&drawing_id.to_string()); // Use drawing_id, not sheet index
            xml.push_str(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>");
            drawing_id += 1;
        }
    }

    if has_shared_strings {
        xml.push_str("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
    }

    xml.push_str("</Types>");
    xml
}

pub fn generate_rels() -> &'static str {
    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\
<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>\
<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>\
</Relationships>"
}

pub fn generate_workbook(sheet_names: &[&str]) -> String {
    let mut xml = String::with_capacity(500 + sheet_names.len() * 80);
    xml.push_str(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
<fileVersion appName=\"xl\" lastEdited=\"7\" lowestEdited=\"7\" rupBuild=\"22621\"/>\
<workbookPr defaultThemeVersion=\"166925\"/>\
<bookViews><workbookView xWindow=\"0\" yWindow=\"0\" windowWidth=\"28800\" windowHeight=\"12600\"/></bookViews>\
<sheets>",
    );

    for (i, name) in sheet_names.iter().enumerate() {
        let id = i + 1;
        xml.push_str("<sheet name=\"");
        // Sheet names may contain & or other metacharacters that pass
        // validate_sheet_name yet still need escaping to keep workbook.xml valid.
        xml.push_str(&xml_escape_str(name));
        xml.push_str("\" sheetId=\"");
        xml.push_str(&id.to_string());
        xml.push_str("\" r:id=\"rId");
        xml.push_str(&id.to_string());
        xml.push_str("\"/>");
    }

    xml.push_str("</sheets><calcPr calcId=\"191029\"/></workbook>");
    xml
}

#[allow(dead_code)]  // superseded by _ext variant
pub fn generate_workbook_rels(num_sheets: usize) -> String {
    generate_workbook_rels_ext(num_sheets, false)
}

pub fn generate_workbook_rels_ext(num_sheets: usize, has_shared_strings: bool) -> String {
    let mut xml = String::with_capacity(300 + num_sheets * 150);
    xml.push_str(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    );

    // Worksheet relationships: workbook.xml references each sheet as r:id="rId{i}"
    // (1-based), so these IDs must match exactly.
    for i in 1..=num_sheets {
        xml.push_str("<Relationship Id=\"rId");
        xml.push_str(&i.to_string());
        xml.push_str("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet");
        xml.push_str(&i.to_string());
        xml.push_str(".xml\"/>");
    }

    // Styles relationship. Per OPC (ECMA-376 Part 2), a relationship Id is an
    // xsd:ID and MUST be unique within the .rels part. The previous fixed value
    // "rId100" collided with sheet #100's "rId100" once a workbook had >= 100
    // sheets -- producing a duplicate ID (invalid package) AND making
    // <sheet r:id="rId100"> resolve to styles.xml instead of the worksheet.
    // Deriving the styles ID from the sheet count guarantees it can never
    // collide, no matter how many sheets there are.
    let styles_rid = num_sheets + 1;
    xml.push_str("<Relationship Id=\"rId");
    xml.push_str(&styles_rid.to_string());
    xml.push_str("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");

    // sharedStrings relationship, given the next unique id after styles.
    if has_shared_strings {
        let sst_rid = num_sheets + 2;
        xml.push_str("<Relationship Id=\"rId");
        xml.push_str(&sst_rid.to_string());
        xml.push_str("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
    }

    xml.push_str("</Relationships>");
    xml
}

/// Generate worksheet relationships (for hyperlinks)
#[allow(dead_code)]  // superseded by _ext variant
pub fn generate_worksheet_rels(hyperlinks: &[(String, usize)]) -> Option<String> {
    if hyperlinks.is_empty() {
        return None;
    }
    
    let mut xml = String::with_capacity(300 + hyperlinks.len() * 150);
    xml.push_str(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    );

    for (url, idx) in hyperlinks {
        xml.push_str("<Relationship Id=\"rId");
        xml.push_str(&idx.to_string());
        xml.push_str("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"");
        // URLs routinely contain `&` (query params) and other metacharacters that
        // corrupt the .rels part if written raw. This produced unopenable files.
        xml.push_str(&xml_escape_str(url));
        xml.push_str("\" TargetMode=\"External\"/>");
    }

    xml.push_str("</Relationships>");
    Some(xml)
}

/// Generate worksheet relationships with table support
#[allow(dead_code)]  // superseded by _ext variant
pub fn generate_worksheet_rels_with_tables(
    hyperlinks: &[(String, usize)],
    tables: &[(String, String)], // (rId, target)
) -> String {
    let mut xml = String::with_capacity(300 + hyperlinks.len() * 150 + tables.len() * 150);
    xml.push_str(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    );

    // Hyperlinks
    for (url, idx) in hyperlinks {
        xml.push_str("<Relationship Id=\"rId");
        xml.push_str(&idx.to_string());
        xml.push_str("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"");
        xml.push_str(&xml_escape_str(url));
        xml.push_str("\" TargetMode=\"External\"/>");
    }

    // Tables (no TargetMode for internal relationships)
    for (rid, target) in tables {
        xml.push_str("<Relationship Id=\"");
        xml.push_str(rid);
        xml.push_str("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"");
        xml.push_str(target);
        xml.push_str("\"/>");
    }

    xml.push_str("</Relationships>");
    xml
}

/// Generate table XML file
pub fn generate_table_xml(
    table: &ExcelTable,
    table_id: u32,
    column_names: &[String],
) -> String {
    let (start_row, start_col, end_row, end_col) = table.range;
    
    let mut xml = String::with_capacity(1000);
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    xml.push_str("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"");
    xml.push_str(&table_id.to_string());
    xml.push_str("\" name=\"");
    // Table name AND displayName must be unique across the whole workbook, or
    // Excel/openpyxl reject the file ("could not read workbook"). Appending the
    // globally-unique table_id guarantees uniqueness deterministically even when
    // a caller reuses a name across sheets, while keeping the user's chosen name
    // as a readable prefix.
    let unique_name = format!("{}_{}", sanitize_table_identifier(&table.name), table_id);
    let unique_display = format!("{}_{}", sanitize_table_identifier(&table.display_name), table_id);
    xml.push_str(&xml_escape_str(&unique_name));
    xml.push_str("\" displayName=\"");
    xml.push_str(&xml_escape_str(&unique_display));
    xml.push_str("\" ref=\"");
    
    // Write range reference
    let mut buf = Vec::with_capacity(32);
    write_cell_ref(start_col, start_row, &mut buf);
    buf.push(b':');
    write_cell_ref(end_col, end_row, &mut buf);
    xml.push_str(&String::from_utf8_lossy(&buf));
    
    xml.push_str("\" totalsRowShown=\"");
    xml.push_str(if table.show_totals_row { "1" } else { "0" });
    xml.push_str("\">");
    
    // AutoFilter (only if header row is shown and no totals row)
    if table.show_header_row {
        xml.push_str("<autoFilter ref=\"");
        buf.clear();
        write_cell_ref(start_col, start_row, &mut buf);
        buf.push(b':');
        write_cell_ref(end_col, end_row, &mut buf);
        xml.push_str(&String::from_utf8_lossy(&buf));
        xml.push_str("\"/>");
    }
    
    // Table columns
    let num_cols = end_col - start_col + 1;
    xml.push_str("<tableColumns count=\"");
    xml.push_str(&num_cols.to_string());
    xml.push_str("\">");
    
    for (idx, col_name) in column_names.iter().enumerate() {
        buf.clear();
        xml.push_str("<tableColumn id=\"");
        xml.push_str(&(idx + 1).to_string());
        xml.push_str("\" name=\"");
        xml_escape_simd(col_name.as_bytes(), &mut buf);
        xml.push_str(&String::from_utf8_lossy(&buf));
        xml.push_str("\"/>");
    }
    
    xml.push_str("</tableColumns>");
    
    // Table style
    if let Some(ref style) = table.style_name {
        xml.push_str("<tableStyleInfo name=\"");
        xml.push_str(&xml_escape_str(style));
        xml.push_str("\" showFirstColumn=\"");
        xml.push_str(if table.show_first_column { "1" } else { "0" });
        xml.push_str("\" showLastColumn=\"");
        xml.push_str(if table.show_last_column { "1" } else { "0" });
        xml.push_str("\" showRowStripes=\"");
        xml.push_str(if table.show_row_stripes { "1" } else { "0" });
        xml.push_str("\" showColumnStripes=\"");
        xml.push_str(if table.show_column_stripes { "1" } else { "0" });
        xml.push_str("\"/>");
    }
    
    xml.push_str("</table>");
    xml
}

/// Calculate exact XML buffer size for Arrow data
fn calculate_exact_xml_size(batches: &[RecordBatch]) -> Result<usize, WriteError> {
    if batches.is_empty() {
        return Ok(200);
    }

    let schema = batches[0].schema();
    let num_cols = schema.fields().len();
    let total_rows: usize = batches.iter().map(|b| b.num_rows()).sum();

    if num_cols == 0 {
        return Ok(200);
    }

    let mut size = 1500;
    size += 50;
    size += 20;
    
    for field in schema.fields().iter() {
        size += 50 + field.name().len();
    }

    for col_idx in 0..num_cols {
        let field = &schema.fields()[col_idx];
        
        if let Some(first_batch) = batches.first() {
            let array = first_batch.column(col_idx);
            let per_cell_size = estimate_cell_xml_size(array.as_ref(), field.data_type())?;
            size += per_cell_size * total_rows;
        }
    }

    size += total_rows * 20;
    size = (size as f64 * 1.3) as usize;

    Ok(size)
}

// ============================================================================
// SHARED STRINGS
// ============================================================================
//
// OOXML lets a workbook store text two ways (Open XML Explained, pp. 64-67):
//   * inline  -- `<c t="inlineStr"><is><t>text</t></is></c>` (self-contained)
//   * shared  -- one `xl/sharedStrings.xml` table holds each distinct string
//                once; cells reference it by index: `<c t="s"><v>7</v></c>`.
//
// Excel itself uses shared strings because, for columns that repeat the same
// values (categories, statuses, regions...), it drastically shrinks the file and
// speeds up the consumer. But for a column of all-distinct strings, a shared
// table is pure overhead: every value is inserted once and never reused, so you
// pay hashing + a second part for no benefit -- and it is measurably SLOWER than
// writing inline (benchmarked at ~9x slower on 1M unique values).
//
// Therefore jetxl decides PER COLUMN with a cheap cardinality probe: sample the
// leading rows, and only route a column through the shared table when its
// distinct-ratio is low enough to win. High-cardinality columns stay on the
// existing inline path, byte-for-byte unchanged, so this feature can never
// regress the unique-string case. The table is built once, before the parallel
// worksheet pass, and the per-cell hot path only ever does read-only lookups
// (thread-safe, ~28M cells/s measured).

// jetxl already depends on `rustc-hash`; reuse its well-tested FxHasher for the
// shared-string maps rather than hand-rolling one. The default SipHash is
// DoS-resistant but needlessly slow for our internal, trusted string keys.
pub use rustc_hash::FxHashMap;

/// Cardinality-probe controls.
const SST_PROBE_ROWS: usize = 1024;
/// Only share a column whose sampled distinct-ratio is at or below this. 0.5
/// means "at least half the sampled values repeat". Chosen from benchmarks: the
/// shared path wins comfortably below this and loses above it.
const SST_SHARE_RATIO: f64 = 0.5;
/// Never bother sharing a column with fewer rows than this -- the table + part
/// overhead isn't worth it and inline is already fast.
const SST_MIN_ROWS: usize = 64;

/// The global shared-string table for one workbook, plus the set of columns that
/// were chosen to use it. Built once (sequentially) before worksheet generation;
/// consulted read-only (and thread-safely) during the per-cell hot path.
#[derive(Default)]
pub struct SharedStrings {
    /// value -> index, for O(1) lookup while writing cells.
    pub map: FxHashMap<String, u32>,
    /// index -> value, preserving insertion order for the sst part.
    pub table: Vec<String>,
    /// Total number of shared-string *references* across all cells (the `count`
    /// attribute on <sst>; `uniqueCount` is `table.len()`).
    pub total_refs: u64,
}

impl SharedStrings {
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.table.is_empty()
    }

    /// Look up an already-interned string. Returns None if the column that owns
    /// this value was not selected for sharing (caller then writes it inline).
    #[inline]
    pub fn get(&self, s: &str) -> Option<u32> {
        self.map.get(s).copied()
    }
}

/// Probe a concrete Arrow string array. Returns true if it should be shared.
fn probe_string_column(values: &dyn Array) -> bool {
    use arrow_array::{StringArray, LargeStringArray};
    let num_rows = values.len();
    if num_rows < SST_MIN_ROWS {
        return false;
    }
    let sample = num_rows.min(SST_PROBE_ROWS);
    let mut seen: FxHashMap<&str, ()> = FxHashMap::default();
    seen.reserve(sample);

    macro_rules! probe {
        ($arr:expr) => {{
            let arr = $arr;
            for i in 0..sample {
                if arr.is_null(i) {
                    continue;
                }
                seen.insert(arr.value(i), ());
                // Early-out: once the distinct ratio is clearly too high, stop.
                if i >= 256 && (seen.len() as f64 / (i + 1) as f64) > SST_SHARE_RATIO {
                    return false;
                }
            }
        }};
    }

    if let Some(a) = values.as_any().downcast_ref::<StringArray>() {
        probe!(a);
    } else if let Some(a) = values.as_any().downcast_ref::<LargeStringArray>() {
        probe!(a);
    } else {
        return false;
    }

    (seen.len() as f64 / sample as f64) <= SST_SHARE_RATIO
}

/// Build the workbook-global shared-string table from all sheets' batches.
///
/// `sheets` yields, per sheet, its list of RecordBatches. Columns are examined
/// across the first batch of each sheet to decide sharing (schemas are stable
/// across a sheet's batches), then every value of every shared column is
/// interned. Runs once, sequentially, before the parallel worksheet pass.
///
/// Returns the table plus, per sheet, the set of column indices that are shared
/// (so the worksheet writer knows which columns to emit as `t="s"`).
pub fn build_shared_strings(
    sheets: &[&[RecordBatch]],
) -> (SharedStrings, Vec<Vec<bool>>) {
    use arrow_array::{StringArray, LargeStringArray};
    use rayon::prelude::*;

    // Phase 1 (parallel, per sheet): probe which columns should be shared, and
    // collect that sheet's DISTINCT shared-column strings plus its total
    // reference count. Each sheet dedups independently against a local set, so
    // the expensive hashing work is spread across cores. Strings are collected as
    // owned `String`s here (the borrow can't outlive the parallel closure), which
    // is the one unavoidable allocation; the serial merge below reuses them.
    //
    // Deterministic order: strings within a sheet are pushed in first-seen order,
    // and sheets are processed in index order during the merge, so the global
    // table indices are identical run-to-run (reproducible output).
    struct SheetLocal {
        shared_cols: Vec<bool>,
        distinct: Vec<String>,
        total_refs: u64,
    }

    let locals: Vec<SheetLocal> = sheets
        .par_iter()
        .map(|batches| {
            if batches.is_empty() {
                return SheetLocal { shared_cols: Vec::new(), distinct: Vec::new(), total_refs: 0 };
            }
            let schema = batches[0].schema();
            let num_cols = schema.fields().len();
            let mut shared_cols = vec![false; num_cols];
            for col_idx in 0..num_cols {
                if matches!(schema.field(col_idx).data_type(), DataType::Utf8 | DataType::LargeUtf8)
                    && probe_string_column(batches[0].column(col_idx).as_ref())
                {
                    shared_cols[col_idx] = true;
                }
            }

            // Local dedup set (first-seen order preserved via `distinct`).
            let mut local_map: FxHashMap<&str, ()> = FxHashMap::default();
            let mut distinct: Vec<String> = Vec::new();
            let mut total_refs: u64 = 0;

            for batch in *batches {
                for col_idx in 0..num_cols {
                    if !shared_cols[col_idx] {
                        continue;
                    }
                    let array = batch.column(col_idx);
                    macro_rules! collect {
                        ($arr:expr) => {{
                            let arr = $arr;
                            for i in 0..arr.len() {
                                if arr.is_null(i) {
                                    continue;
                                }
                                let v = arr.value(i);
                                total_refs += 1;
                                if !local_map.contains_key(v) {
                                    local_map.insert(v, ());
                                    distinct.push(v.to_string());
                                }
                            }
                        }};
                    }
                    if let Some(a) = array.as_any().downcast_ref::<StringArray>() {
                        collect!(a);
                    } else if let Some(a) = array.as_any().downcast_ref::<LargeStringArray>() {
                        collect!(a);
                    }
                }
            }

            SheetLocal { shared_cols, distinct, total_refs }
        })
        .collect();

    // Phase 2 (serial): merge the per-sheet distinct lists into one global table
    // with stable indices. Only genuinely-new strings are inserted; duplicates
    // across sheets collapse. This pass touches each distinct string once and does
    // no per-cell work, so it is far cheaper than the original all-cells-serial
    // interning it replaces.
    let mut sst = SharedStrings::default();
    let mut per_sheet_shared: Vec<Vec<bool>> = Vec::with_capacity(sheets.len());

    for local in locals {
        sst.total_refs += local.total_refs;
        for s in local.distinct {
            if !sst.map.contains_key(s.as_str()) {
                let idx = sst.table.len() as u32;
                sst.map.insert(s.clone(), idx);
                sst.table.push(s);
            }
        }
        per_sheet_shared.push(local.shared_cols);
    }

    (sst, per_sheet_shared)
}

/// Serialize the shared-string table to the `xl/sharedStrings.xml` part.
pub fn generate_shared_strings_xml(sst: &SharedStrings) -> Vec<u8> {
    // Rough sizing: header + per-entry markup + payload.
    let payload: usize = sst.table.iter().map(|s| s.len()).sum();
    let mut buf = Vec::with_capacity(128 + sst.table.len() * 16 + payload + payload / 8);

    buf.extend_from_slice(
        b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"",
    );
    buf.extend_from_slice(itoa::Buffer::new().format(sst.total_refs).as_bytes());
    buf.extend_from_slice(b"\" uniqueCount=\"");
    buf.extend_from_slice(itoa::Buffer::new().format(sst.table.len()).as_bytes());
    buf.extend_from_slice(b"\">");

    for s in &sst.table {
        // Preserve significant leading/trailing whitespace, matching the inline
        // path's behaviour and what Excel itself emits. NOTE: the strict ECMA-376
        // transitional schema types <si><t> as a bare xsd:string (ST_Xstring) and
        // technically disallows the xml:space attribute here, but every real
        // consumer (Excel, LibreOffice, openpyxl) both accepts and requires it to
        // avoid silently trimming values like " 007". We follow real-world Excel
        // behaviour, identical to the inline <is><t> path.
        let bytes = s.as_bytes();
        let needs_preserve = matches!(bytes.first(), Some(b) if b.is_ascii_whitespace())
            || matches!(bytes.last(), Some(b) if b.is_ascii_whitespace());
        if needs_preserve {
            buf.extend_from_slice(b"<si><t xml:space=\"preserve\">");
        } else {
            buf.extend_from_slice(b"<si><t>");
        }
        xml_escape_simd(bytes, &mut buf);
        buf.extend_from_slice(b"</t></si>");
    }

    buf.extend_from_slice(b"</sst>");
    buf
}

/// Validate, ONCE per sheet, that every column has a data type the per-cell
/// writer supports. This lets the hot loop keep using infallible
/// `downcast_ref().unwrap()` (any supported type is guaranteed to downcast) while
/// still giving the user a clear, catchable error for an unsupported column
/// instead of either a silent empty column or a `panic = "abort"` process kill.
///
/// The check is O(num_cols), not O(cells), so it costs nothing measurable.
fn validate_arrow_schema(schema: &arrow_schema::Schema) -> Result<(), WriteError> {
    for field in schema.fields() {
        if !is_supported_data_type(field.data_type()) {
            return Err(WriteError::Validation(format!(
                "Column '{}' has unsupported Arrow type {:?}. Supported types: \
                 integers, floats, boolean, Utf8/LargeUtf8, Date32/64, \
                 Time32/64, and Timestamp. Cast the column before writing.",
                field.name(),
                field.data_type()
            )));
        }
    }
    Ok(())
}

/// The exact set of types handled by `write_arrow_cell_to_xml_optimized`. Keep
/// this in lockstep with that function's match arms.
fn is_supported_data_type(dt: &DataType) -> bool {
    matches!(
        dt,
        DataType::Null            // all-None column -> every cell empty
            | DataType::Utf8
            | DataType::LargeUtf8
            | DataType::Int8
            | DataType::Int16
            | DataType::Int32
            | DataType::Int64
            | DataType::UInt8
            | DataType::UInt16
            | DataType::UInt32
            | DataType::UInt64
            | DataType::Float32
            | DataType::Float64
            | DataType::Boolean
            | DataType::Date32
            | DataType::Date64
            | DataType::Time32(_)
            | DataType::Time64(_)
            | DataType::Timestamp(_, _)
    )
}

fn estimate_cell_xml_size(array: &dyn Array, data_type: &DataType) -> Result<usize, WriteError> {
    use arrow_array::*;
    
    match data_type {
        DataType::Utf8 => {
            let arr = array.as_any().downcast_ref::<StringArray>()
                .ok_or_else(|| WriteError::Validation("Type mismatch".to_string()))?;
            
            let num_rows = arr.len();
            if num_rows == 0 {
                return Ok(25);
            }
            
            let total_string_bytes = get_string_array_total_bytes(arr);
            let avg_string_len = total_string_bytes / num_rows.max(1);
            
            Ok(55 + avg_string_len + (avg_string_len / 10))
        }
        
        DataType::LargeUtf8 => {
            let arr = array.as_any().downcast_ref::<LargeStringArray>()
                .ok_or_else(|| WriteError::Validation("Type mismatch".to_string()))?;
            
            let num_rows = arr.len();
            if num_rows == 0 {
                return Ok(25);
            }
            
            let total_string_bytes = get_large_string_array_total_bytes(arr);
            let avg_string_len = total_string_bytes / num_rows.max(1);
            Ok(55 + avg_string_len + (avg_string_len / 10))
        }
        DataType::Int8 | DataType::Int16 | DataType::Int32 | DataType::Int64 |
        DataType::UInt8 | DataType::UInt16 | DataType::UInt32 | DataType::UInt64 => {
            Ok(33)
        }
        DataType::Float32 | DataType::Float64 => {
            Ok(35)
        }
        DataType::Boolean => {
            Ok(28)
        }
        DataType::Date32 | DataType::Date64 | DataType::Timestamp(_, _) => {
            Ok(35)
        }
        _ => {
            Ok(20)
        }
    }
}

fn get_string_array_total_bytes(arr: &arrow_array::StringArray) -> usize {
    use arrow_array::Array;

    if arr.len() == 0 {
        return 0;
    }
    // Arrow keeps all string bytes in one contiguous values buffer, and the
    // offsets are monotonic. The total payload is therefore just
    // last_offset - first_offset, an O(1) read instead of an O(rows) loop of
    // per-value length calls. (Nulls occupy zero-width offset ranges, so they
    // don't inflate the total.)
    let offsets = arr.offsets();
    let first = offsets[0] as usize;
    let last = offsets[arr.len()] as usize;
    last - first
}

fn get_large_string_array_total_bytes(arr: &arrow_array::LargeStringArray) -> usize {
    use arrow_array::Array;

    if arr.len() == 0 {
        return 0;
    }
    let offsets = arr.offsets();
    let first = offsets[0] as usize;
    let last = offsets[arr.len()] as usize;
    last - first
}

/// Generate drawing XML for chart positioning
pub fn generate_drawing_xml(charts: &[ExcelChart]) -> String {
    let mut xml = String::with_capacity(2000 + charts.len() * 1000);
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    xml.push_str("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" ");
    xml.push_str("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n");
    
    for (idx, chart) in charts.iter().enumerate() {
        let chart_id = idx + 1;
        xml.push_str("<xdr:twoCellAnchor>\n");
        
        // From marker
        xml.push_str("<xdr:from>\n");
        xml.push_str(&format!("<xdr:col>{}</xdr:col>\n", chart.position.from_col));
        xml.push_str("<xdr:colOff>0</xdr:colOff>\n");
        xml.push_str(&format!("<xdr:row>{}</xdr:row>\n", chart.position.from_row));
        xml.push_str("<xdr:rowOff>0</xdr:rowOff>\n");
        xml.push_str("</xdr:from>\n");
        
        // To marker
        xml.push_str("<xdr:to>\n");
        xml.push_str(&format!("<xdr:col>{}</xdr:col>\n", chart.position.to_col));
        xml.push_str("<xdr:colOff>0</xdr:colOff>\n");
        xml.push_str(&format!("<xdr:row>{}</xdr:row>\n", chart.position.to_row));
        xml.push_str("<xdr:rowOff>0</xdr:rowOff>\n");
        xml.push_str("</xdr:to>\n");
        
        // Graphic frame
        xml.push_str("<xdr:graphicFrame macro=\"\">\n");
        xml.push_str("<xdr:nvGraphicFramePr>\n");
        xml.push_str(&format!("<xdr:cNvPr id=\"{}\" name=\"Chart {}\"/>\n", chart_id + 1000, chart_id));
        xml.push_str("<xdr:cNvGraphicFramePr/>\n");
        xml.push_str("</xdr:nvGraphicFramePr>\n");
        xml.push_str("<xdr:xfrm>\n");
        xml.push_str("<a:off x=\"0\" y=\"0\"/>\n");
        xml.push_str("<a:ext cx=\"0\" cy=\"0\"/>\n");
        xml.push_str("</xdr:xfrm>\n");
        xml.push_str("<a:graphic>\n");
        xml.push_str("<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\n");
        xml.push_str(&format!("<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdChart{}\"/>\n", chart_id));
        xml.push_str("</a:graphicData>\n");
        xml.push_str("</a:graphic>\n");
        xml.push_str("</xdr:graphicFrame>\n");
        xml.push_str("<xdr:clientData/>\n");
        xml.push_str("</xdr:twoCellAnchor>\n");
    }
    
    xml.push_str("</xdr:wsDr>");
    xml
}

fn get_column_letter(col: usize) -> String {
    let mut buf = [0u8; 4];
    let len = write_col_letter(col, &mut buf);
    std::str::from_utf8(&buf[..len]).unwrap().to_string()
}

/// Generate chart XML
pub fn generate_chart_xml(chart: &ExcelChart, sheet_name: &str) -> String {
    let mut xml = String::with_capacity(8000);
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    xml.push_str("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" ");
    xml.push_str("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" ");
    xml.push_str("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ");
    xml.push_str("xmlns:c16r2=\"http://schemas.microsoft.com/office/drawing/2015/06/chart\">");
    
    xml.push_str("<c:date1904 val=\"0\"/>\n");
    xml.push_str("<c:lang val=\"en-US\"/>\n");
    xml.push_str("<c:roundedCorners val=\"0\"/>\n");
    
    // Chart style
    if let Some(style) = chart.chart_style {
        xml.push_str("<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">");
        xml.push_str(&format!("<mc:Choice Requires=\"c14\" xmlns:c14=\"http://schemas.microsoft.com/office/drawing/2007/8/2/chart\"><c14:style val=\"{}\"/></mc:Choice>", style));
        xml.push_str(&format!("<mc:Fallback><c:style val=\"{}\"/></mc:Fallback>", if style >= 100 { style - 100 } else { style }));
        xml.push_str("</mc:AlternateContent>\n");
    }
    
    xml.push_str("<c:chart>\n");
    
    // Title with formatting
    if let Some(ref title) = chart.title {
        xml.push_str("<c:title>\n");
        xml.push_str("<c:tx><c:rich>\n");
        xml.push_str("<a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n");
        xml.push_str("<a:lstStyle/>\n");
        xml.push_str("<a:p><a:pPr>\n");
        
        let font_size = chart.title_font_size.unwrap_or(1400).clamp(100, 400000);
        xml.push_str(&format!("<a:defRPr sz=\"{}\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" spc=\"0\" baseline=\"0\">\n", font_size));
        
        if let Some(ref color) = chart.title_color {
            if let Some(c) = crate::styles::normalize_color_rgb(color) { xml.push_str(&format!("<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>\n", c)); }
        } else {
            xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill>\n");
        }
        
        xml.push_str("<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\n");
        xml.push_str("</a:defRPr>\n");
        xml.push_str("</a:pPr>\n");
        xml.push_str("<a:r>\n");
        xml.push_str("<a:rPr lang=\"en-US\"");
        if chart.title_bold {
            xml.push_str(" b=\"1\"");
        }
        xml.push_str("/>\n");
        xml.push_str(&format!("<a:t>{}</a:t>\n", xml_escape_str(title)));
        xml.push_str("</a:r>\n");
        xml.push_str("</a:p>\n");
        xml.push_str("</c:rich></c:tx>\n");
        xml.push_str("<c:overlay val=\"0\"/>\n");
        xml.push_str("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>\n");
        xml.push_str("<c:txPr>\n");
        xml.push_str("<a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n");
        xml.push_str("<a:lstStyle/>\n");
        xml.push_str("<a:p><a:pPr>\n");
        xml.push_str(&format!("<a:defRPr sz=\"{}\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" spc=\"0\" baseline=\"0\">\n", font_size));
        
        if let Some(ref color) = chart.title_color {
            if let Some(c) = crate::styles::normalize_color_rgb(color) { xml.push_str(&format!("<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>\n", c)); }
        } else {
            xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill>\n");
        }
        
        xml.push_str("<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\n");
        xml.push_str("</a:defRPr>\n");
        xml.push_str("</a:pPr>\n");
        xml.push_str("<a:endParaRPr lang=\"en-US\"/>\n");
        xml.push_str("</a:p>\n");
        xml.push_str("</c:txPr>\n");
        xml.push_str("</c:title>\n");
    }
    
    xml.push_str("<c:autoTitleDeleted val=\"0\"/>\n");
    
    // Plot area
    xml.push_str("<c:plotArea>\n");
    xml.push_str("<c:layout/>\n");
    
    // Chart-specific content
    match chart.chart_type {
        ChartType::Column => generate_column_chart_content(&mut xml, chart, sheet_name),
        ChartType::Bar => generate_bar_chart_content(&mut xml, chart, sheet_name),
        ChartType::Line => generate_line_chart_content(&mut xml, chart, sheet_name),
        ChartType::Pie => generate_pie_chart_content(&mut xml, chart, sheet_name),
        ChartType::Scatter => generate_scatter_chart_content(&mut xml, chart, sheet_name),
        ChartType::Area => generate_area_chart_content(&mut xml, chart, sheet_name),
    }
    
    xml.push_str("</c:plotArea>\n");
    
    // Legend with styling
    if chart.show_legend && !matches!(chart.legend_position, LegendPosition::None) {
        xml.push_str("<c:legend>\n");
        xml.push_str(&format!("<c:legendPos val=\"{}\"/>\n", match chart.legend_position {
            LegendPosition::Right => "r",
            LegendPosition::Left => "l",
            LegendPosition::Top => "t",
            LegendPosition::Bottom => "b",
            LegendPosition::None => "r",
        }));
        xml.push_str("<c:overlay val=\"0\"/>\n");
        xml.push_str("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>\n");
        xml.push_str("<c:txPr>\n");
        xml.push_str("<a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n");
        xml.push_str("<a:lstStyle/>\n");
        xml.push_str("<a:p><a:pPr>\n");
        
        let legend_size = chart.legend_font_size.unwrap_or(900).clamp(100, 400000);
        xml.push_str(&format!("<a:defRPr sz=\"{}\"", legend_size));
        if chart.legend_bold {
            xml.push_str(" b=\"1\"");
        } else {
            xml.push_str(" b=\"0\"");
        }
        xml.push_str(" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n");
        xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill>\n");
        xml.push_str("<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\n");
        xml.push_str("</a:defRPr>\n");
        xml.push_str("</a:pPr><a:endParaRPr lang=\"en-US\"/></a:p>\n");
        xml.push_str("</c:txPr>\n");
        xml.push_str("</c:legend>\n");
    }
    
    xml.push_str("<c:plotVisOnly val=\"1\"/>\n");
    // Area charts use "zero" for dispBlanksAs, other charts use "gap"
    let disp_blanks = if matches!(chart.chart_type, ChartType::Area) { "zero" } else { "gap" };
    xml.push_str(&format!("<c:dispBlanksAs val=\"{}\"/>\n", disp_blanks));
    xml.push_str("<c:showDLblsOverMax val=\"0\"/>\n");
    xml.push_str("</c:chart>\n");
    
    xml.push_str("<c:spPr>\n");
    xml.push_str("<a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill>\n");
    xml.push_str("<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">\n");
    xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill>\n");
    xml.push_str("<a:round/></a:ln>\n");
    xml.push_str("<a:effectLst/>\n");
    xml.push_str("</c:spPr>\n");
    
    xml.push_str("<c:txPr><a:bodyPr/><a:lstStyle/>\n");
    xml.push_str("<a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang=\"en-US\"/></a:p>\n");
    xml.push_str("</c:txPr>\n");
    
    xml.push_str("<c:printSettings>\n");
    xml.push_str("<c:headerFooter/>\n");
    xml.push_str("<c:pageMargins b=\"0.75\" l=\"0.7\" r=\"0.7\" t=\"0.75\" header=\"0.3\" footer=\"0.3\"/>\n");
    xml.push_str("<c:pageSetup/>\n");
    xml.push_str("</c:printSettings>\n");
    
    xml.push_str("</c:chartSpace>");
    xml
}



// Helper function for axis styling
fn write_axis_title(xml: &mut String, title: &str, chart: &ExcelChart) {
    xml.push_str("<c:title>\n");
    xml.push_str("<c:overlay val=\"0\"/>\n");
    xml.push_str("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>\n");
    xml.push_str("<c:txPr>\n");
    xml.push_str("<a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n");
    xml.push_str("<a:lstStyle/>\n");
    xml.push_str("<a:p>\n");
    xml.push_str("<a:pPr>\n");
    
    let font_size = chart.axis_title_font_size.unwrap_or(1000).clamp(100, 400000);
    xml.push_str(&format!("<a:defRPr sz=\"{}\"", font_size));
    if chart.axis_title_bold {
        xml.push_str(" b=\"1\"");
    } else {
        xml.push_str(" b=\"0\"");
    }
    xml.push_str(" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n");
    
    if let Some(ref color) = chart.axis_title_color {
        if let Some(c) = crate::styles::normalize_color_rgb(color) { xml.push_str(&format!("<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>\n", c)); }
    } else {
        xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill>\n");
    }
    
    xml.push_str("<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\n");
    xml.push_str("</a:defRPr>\n");
    xml.push_str("</a:pPr>\n");
    xml.push_str("<a:r>\n");
    xml.push_str("<a:rPr lang=\"en-US\"/>\n");
    xml.push_str(&format!("<a:t>{}</a:t>\n", xml_escape_str(title)));
    xml.push_str("</a:r>\n");
    xml.push_str("<a:endParaRPr lang=\"en-US\"/>\n");
    xml.push_str("</a:p>\n");
    xml.push_str("</c:txPr>\n");
    xml.push_str("</c:title>\n");
}

fn write_data_labels(xml: &mut String, show_values: bool) {
    xml.push_str("<c:dLbls>\n");
    xml.push_str("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>\n");
    xml.push_str("<c:txPr>\n");
    xml.push_str("<a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" lIns=\"38100\" tIns=\"19050\" rIns=\"38100\" bIns=\"19050\" anchor=\"ctr\" anchorCtr=\"1\"><a:spAutoFit/></a:bodyPr>\n");
    xml.push_str("<a:lstStyle/>\n");
    xml.push_str("<a:p>\n");
    xml.push_str("<a:pPr>\n");
    xml.push_str("<a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n");
    xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"75000\"/><a:lumOff val=\"25000\"/></a:schemeClr></a:solidFill>\n");
    xml.push_str("<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\n");
    xml.push_str("</a:defRPr>\n");
    xml.push_str("</a:pPr>\n");
    xml.push_str("<a:endParaRPr lang=\"en-US\"/>\n");
    xml.push_str("</a:p>\n");
    xml.push_str("</c:txPr>\n");
    xml.push_str("<c:dLblPos val=\"ctr\"/>\n");
    xml.push_str("<c:showLegendKey val=\"0\"/>\n");
    xml.push_str(&format!("<c:showVal val=\"{}\"/>\n", if show_values { "1" } else { "0" }));
    xml.push_str("<c:showCatName val=\"0\"/>\n");
    xml.push_str("<c:showSerName val=\"0\"/>\n");
    xml.push_str("<c:showPercent val=\"0\"/>\n");
    xml.push_str("<c:showBubbleSize val=\"0\"/>\n");
    xml.push_str("<c:showLeaderLines val=\"0\"/>\n");
    xml.push_str("<c:extLst><c:ext uri=\"{CE6537A1-D6FC-4f65-9D91-7224C49458BB}\" xmlns:c15=\"http://schemas.microsoft.com/office/drawing/2012/chart\">");
    xml.push_str("<c15:showLeaderLines val=\"1\"/>");
    xml.push_str("<c15:leaderLines><c:spPr>");
    xml.push_str("<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">");
    xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"35000\"/><a:lumOff val=\"65000\"/></a:schemeClr></a:solidFill>");
    xml.push_str("<a:round/></a:ln>");
    xml.push_str("<a:effectLst/></c:spPr></c15:leaderLines>");
    xml.push_str("</c:ext></c:extLst>\n");
    xml.push_str("</c:dLbls>\n");
}

// Common axis styling components
fn write_category_axis_styling(xml: &mut String) {
    xml.push_str("<c:spPr><a:noFill/>\n");
    xml.push_str("<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">\n");
    xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill>\n");
    xml.push_str("<a:round/></a:ln>\n");
    xml.push_str("<a:effectLst/></c:spPr>\n");
    xml.push_str("<c:txPr>\n");
    xml.push_str("<a:bodyPr rot=\"-60000000\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n");
    xml.push_str("<a:lstStyle/>\n");
    xml.push_str("<a:p><a:pPr>\n");
    xml.push_str("<a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n");
    xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill>\n");
    xml.push_str("<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\n");
    xml.push_str("</a:defRPr>\n");
    xml.push_str("</a:pPr><a:endParaRPr lang=\"en-US\"/></a:p>\n");
    xml.push_str("</c:txPr>\n");
}

fn write_value_axis_styling(xml: &mut String) {
    xml.push_str("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>\n");
    xml.push_str("<c:txPr>\n");
    xml.push_str("<a:bodyPr rot=\"-60000000\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n");
    xml.push_str("<a:lstStyle/>\n");
    xml.push_str("<a:p><a:pPr>\n");
    xml.push_str("<a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n");
    xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill>\n");
    xml.push_str("<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\n");
    xml.push_str("</a:defRPr>\n");
    xml.push_str("</a:pPr><a:endParaRPr lang=\"en-US\"/></a:p>\n");
    xml.push_str("</c:txPr>\n");
}

fn write_major_gridlines(xml: &mut String) {
    xml.push_str("<c:majorGridlines>\n");
    xml.push_str("<c:spPr>\n");
    xml.push_str("<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">\n");
    xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill>\n");
    xml.push_str("<a:round/></a:ln>\n");
    xml.push_str("<a:effectLst/>\n");
    xml.push_str("</c:spPr>\n");
    xml.push_str("</c:majorGridlines>\n");
}

fn generate_column_chart_content(xml: &mut String, chart: &ExcelChart, sheet_name: &str) {
    xml.push_str("<c:barChart>\n");
    xml.push_str("<c:barDir val=\"col\"/>\n");
    xml.push_str(&format!("<c:grouping val=\"{}\"/>\n", 
        if chart.percent_stacked { "percentStacked" } else if chart.stacked { "stacked" } else { "clustered" }));
    xml.push_str("<c:varyColors val=\"0\"/>\n");
    
    let (start_row, start_col, end_row, end_col) = chart.data_range;
    let category_col = chart.category_col.unwrap_or(start_col);
    
    let accent_colors = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
    let tint_shade_values = [("tint", "65000"), ("", ""), ("shade", "65000")];
    
    let mut actual_series_idx = 0;
    for col in start_col..=end_col {
        if Some(col) == chart.category_col {
            continue;
        }
        
        let series_name = chart.series_names.get(actual_series_idx).map(|s| s.as_str()).unwrap_or("Series");
        let accent_color = accent_colors[actual_series_idx % accent_colors.len()];
        let (modifier, value) = tint_shade_values[actual_series_idx % tint_shade_values.len()];
        
        xml.push_str(&format!("<c:ser>\n<c:idx val=\"{}\"/>\n<c:order val=\"{}\"/>\n", actual_series_idx, actual_series_idx));
        
        // Series name
        xml.push_str("<c:tx>\n<c:strRef>\n<c:f>");
        xml.push_str(&format!("{}!${}$1", sheet_name, get_column_letter(col)));
        xml.push_str("</c:f>\n<c:strCache>\n<c:ptCount val=\"1\"/>\n<c:pt idx=\"0\">\n");
        xml.push_str(&format!("<c:v>{}</c:v>\n", xml_escape_str(series_name)));
        xml.push_str("</c:pt>\n</c:strCache>\n</c:strRef>\n</c:tx>\n");
        
        // Series styling with scheme colors and tint/shade
        xml.push_str("<c:spPr>\n");
        xml.push_str(&format!("<a:solidFill><a:schemeClr val=\"{}\">", accent_color));
        if !modifier.is_empty() {
            xml.push_str(&format!("<a:{} val=\"{}\"/>", modifier, value));
        }
        xml.push_str("</a:schemeClr></a:solidFill>\n");
        xml.push_str("<a:ln><a:noFill/></a:ln>\n");
        xml.push_str("<a:effectLst/>\n");
        xml.push_str("</c:spPr>\n");
        xml.push_str("<c:invertIfNegative val=\"0\"/>\n");
        
        // Data labels per series for stacked charts
        if chart.stacked || chart.percent_stacked {
            write_data_labels(xml, chart.show_data_labels.unwrap_or(false));
        }
        
        // Category axis data
        xml.push_str("<c:cat>\n<c:strRef>\n<c:f>");
        xml.push_str(&format!("{}!${}${}:${}${}", 
            sheet_name, get_column_letter(category_col), start_row + 1, 
            get_column_letter(category_col), end_row + 1));
        xml.push_str("</c:f>\n</c:strRef>\n</c:cat>\n");
        
        // Values
        xml.push_str("<c:val>\n<c:numRef>\n<c:f>");
        xml.push_str(&format!("{}!${}${}:${}${}", 
            sheet_name, get_column_letter(col), start_row + 1, 
            get_column_letter(col), end_row + 1));
        xml.push_str("</c:f>\n</c:numRef>\n</c:val>\n");
        
        // Add extLst with uniqueId for modern Excel compatibility
        xml.push_str("<c:extLst><c:ext uri=\"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\">");
        xml.push_str(&format!("<c16:uniqueId val=\"{{0000000{}-6E8F-43DD-B1F6-30AC1D0140EF}}\"/>", actual_series_idx));
        xml.push_str("</c:ext></c:extLst>\n");
        
        xml.push_str("</c:ser>\n");
        actual_series_idx += 1;
    }
    
    // Chart-level data labels
    if !chart.stacked && !chart.percent_stacked {
        write_data_labels(xml, chart.show_data_labels.unwrap_or(false));
    }
    
    xml.push_str("<c:gapWidth val=\"150\"/>\n");
    if chart.stacked || chart.percent_stacked {
        xml.push_str("<c:overlap val=\"100\"/>\n");
    }
    xml.push_str("<c:axId val=\"100000001\"/>\n");
    xml.push_str("<c:axId val=\"100000002\"/>\n");
    xml.push_str("</c:barChart>\n");
    
    // Category axis
    xml.push_str("<c:catAx>\n");
    xml.push_str("<c:axId val=\"100000001\"/>\n");
    xml.push_str("<c:scaling><c:orientation val=\"minMax\"/></c:scaling>\n");
    xml.push_str("<c:delete val=\"0\"/>\n");
    xml.push_str("<c:axPos val=\"b\"/>\n");
    if let Some(ref x_title) = chart.x_axis_title {
        write_axis_title(xml, x_title, chart);
    }
    xml.push_str("<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>\n");
    xml.push_str("<c:majorTickMark val=\"none\"/>\n");
    xml.push_str("<c:minorTickMark val=\"none\"/>\n");
    xml.push_str("<c:tickLblPos val=\"nextTo\"/>\n");
    xml.push_str("<c:spPr><a:noFill/>\n");
    xml.push_str("<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">\n");
    xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill>\n");
    xml.push_str("<a:round/></a:ln>\n");
    xml.push_str("<a:effectLst/></c:spPr>\n");
    xml.push_str("<c:txPr>\n");
    xml.push_str("<a:bodyPr rot=\"-60000000\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n");
    xml.push_str("<a:lstStyle/>\n");
    xml.push_str("<a:p><a:pPr>\n");
    xml.push_str("<a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n");
    xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill>\n");
    xml.push_str("<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\n");
    xml.push_str("</a:defRPr>\n");
    xml.push_str("</a:pPr><a:endParaRPr lang=\"en-US\"/></a:p>\n");
    xml.push_str("</c:txPr>\n");
    xml.push_str("<c:crossAx val=\"100000002\"/>\n");
    xml.push_str("<c:crosses val=\"autoZero\"/>\n");
    xml.push_str("<c:auto val=\"1\"/>\n");
    xml.push_str("<c:lblAlgn val=\"ctr\"/>\n");
    xml.push_str("<c:lblOffset val=\"100\"/>\n");
    xml.push_str("<c:noMultiLvlLbl val=\"0\"/>\n");
    xml.push_str("</c:catAx>\n");
    
    // Value axis
    xml.push_str("<c:valAx>\n");
    xml.push_str("<c:axId val=\"100000002\"/>\n");
    xml.push_str("<c:scaling>\n");
    xml.push_str("<c:orientation val=\"minMax\"/>\n");
    if let Some(min) = chart.axis_min {
        xml.push_str(&format!("<c:min val=\"{}\"/>\n", min));
    }
    if let Some(max) = chart.axis_max {
        xml.push_str(&format!("<c:max val=\"{}\"/>\n", max));
    }
    xml.push_str("</c:scaling>\n");
    xml.push_str("<c:delete val=\"0\"/>\n");
    xml.push_str("<c:axPos val=\"l\"/>\n");
    xml.push_str("<c:majorGridlines>\n");
    xml.push_str("<c:spPr>\n");
    xml.push_str("<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">\n");
    xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill>\n");
    xml.push_str("<a:round/></a:ln>\n");
    xml.push_str("<a:effectLst/>\n");
    xml.push_str("</c:spPr>\n");
    xml.push_str("</c:majorGridlines>\n");
    if let Some(ref y_title) = chart.y_axis_title {
        xml.push_str("<c:title>\n");
        xml.push_str("<c:overlay val=\"0\"/>\n");
        xml.push_str("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>\n");
        xml.push_str("<c:txPr>\n");
        xml.push_str("<a:bodyPr rot=\"-5400000\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n");
        xml.push_str("<a:lstStyle/>\n");
        xml.push_str("<a:p>\n");
        xml.push_str("<a:pPr>\n");
        
        let font_size = chart.axis_title_font_size.unwrap_or(1000).clamp(100, 400000);
        xml.push_str(&format!("<a:defRPr sz=\"{}\"", font_size));
        if chart.axis_title_bold {
            xml.push_str(" b=\"1\"");
        } else {
            xml.push_str(" b=\"0\"");
        }
        xml.push_str(" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n");
        
        if let Some(ref color) = chart.axis_title_color {
            if let Some(c) = crate::styles::normalize_color_rgb(color) { xml.push_str(&format!("<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>\n", c)); }
        } else {
            xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill>\n");
        }
        
        xml.push_str("<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\n");
        xml.push_str("</a:defRPr>\n");
        xml.push_str("</a:pPr>\n");
        xml.push_str("<a:r>\n");
        xml.push_str("<a:rPr lang=\"en-US\"/>\n");
        xml.push_str(&format!("<a:t>{}</a:t>\n", xml_escape_str(y_title)));
        xml.push_str("</a:r>\n");
        xml.push_str("<a:endParaRPr lang=\"en-US\"/>\n");
        xml.push_str("</a:p>\n");
        xml.push_str("</c:txPr>\n");
        xml.push_str("</c:title>\n");
    }
    
    // Format code for percentage stacked charts
    let format_code = if chart.percent_stacked { "0%" } else { "General" };
    xml.push_str(&format!("<c:numFmt formatCode=\"{}\" sourceLinked=\"1\"/>\n", format_code));
    xml.push_str("<c:majorTickMark val=\"none\"/>\n");
    xml.push_str("<c:minorTickMark val=\"none\"/>\n");
    xml.push_str("<c:tickLblPos val=\"nextTo\"/>\n");
    xml.push_str("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>\n");
    xml.push_str("<c:txPr>\n");
    xml.push_str("<a:bodyPr rot=\"-60000000\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>\n");
    xml.push_str("<a:lstStyle/>\n");
    xml.push_str("<a:p><a:pPr>\n");
    xml.push_str("<a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n");
    xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill>\n");
    xml.push_str("<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\n");
    xml.push_str("</a:defRPr>\n");
    xml.push_str("</a:pPr><a:endParaRPr lang=\"en-US\"/></a:p>\n");
    xml.push_str("</c:txPr>\n");
    xml.push_str("<c:crossAx val=\"100000001\"/>\n");
    xml.push_str("<c:crosses val=\"autoZero\"/>\n");
    xml.push_str("<c:crossBetween val=\"between\"/>\n");
    xml.push_str("</c:valAx>\n");
    xml.push_str("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>\n");
}

// ============================================================================
// BAR CHART (Horizontal bars - barDir="bar")
// ============================================================================
fn generate_bar_chart_content(xml: &mut String, chart: &ExcelChart, sheet_name: &str) {
    xml.push_str("<c:barChart>\n");
    xml.push_str("<c:barDir val=\"bar\"/>\n");
    xml.push_str(&format!("<c:grouping val=\"{}\"/>\n", 
        if chart.percent_stacked { "percentStacked" } else if chart.stacked { "stacked" } else { "clustered" }));
    xml.push_str("<c:varyColors val=\"0\"/>\n");
    
    let (start_row, start_col, end_row, end_col) = chart.data_range;
    let category_col = chart.category_col.unwrap_or(start_col);
    let accent_colors = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
    let tint_shade_values = [("tint", "65000"), ("", ""), ("shade", "65000")];
    
    let mut actual_series_idx = 0;
    for col in start_col..=end_col {
        if Some(col) == chart.category_col {
            continue;
        }
        
        let series_name = chart.series_names.get(actual_series_idx).map(|s| s.as_str()).unwrap_or("Series");
        let accent_color = accent_colors[actual_series_idx % accent_colors.len()];
        let (modifier, value) = tint_shade_values[actual_series_idx % tint_shade_values.len()];
        
        xml.push_str(&format!("<c:ser>\n<c:idx val=\"{}\"/>\n<c:order val=\"{}\"/>\n", actual_series_idx, actual_series_idx));
        
        xml.push_str("<c:tx>\n<c:strRef>\n<c:f>");
        xml.push_str(&format!("{}!${}$1", sheet_name, get_column_letter(col)));
        xml.push_str("</c:f>\n<c:strCache>\n<c:ptCount val=\"1\"/>\n<c:pt idx=\"0\">\n");
        xml.push_str(&format!("<c:v>{}</c:v>\n", xml_escape_str(series_name)));
        xml.push_str("</c:pt>\n</c:strCache>\n</c:strRef>\n</c:tx>\n");
        
        xml.push_str("<c:spPr>\n");
        xml.push_str(&format!("<a:solidFill><a:schemeClr val=\"{}\">", accent_color));
        if !modifier.is_empty() {
            xml.push_str(&format!("<a:{} val=\"{}\"/>", modifier, value));
        }
        xml.push_str("</a:schemeClr></a:solidFill>\n");
        xml.push_str("<a:ln><a:noFill/></a:ln>\n");
        xml.push_str("<a:effectLst/>\n");
        xml.push_str("</c:spPr>\n");
        xml.push_str("<c:invertIfNegative val=\"0\"/>\n");
        
        if chart.stacked || chart.percent_stacked {
            write_data_labels(xml, chart.show_data_labels.unwrap_or(false));
        }
        
        xml.push_str("<c:cat>\n<c:strRef>\n<c:f>");
        xml.push_str(&format!("{}!${}${}:${}${}", 
            sheet_name, get_column_letter(category_col), start_row + 1, 
            get_column_letter(category_col), end_row + 1));
        xml.push_str("</c:f>\n</c:strRef>\n</c:cat>\n");
        
        xml.push_str("<c:val>\n<c:numRef>\n<c:f>");
        xml.push_str(&format!("{}!${}${}:${}${}", 
            sheet_name, get_column_letter(col), start_row + 1, 
            get_column_letter(col), end_row + 1));
        xml.push_str("</c:f>\n</c:numRef>\n</c:val>\n");
        
        xml.push_str("<c:extLst><c:ext uri=\"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\">");
        xml.push_str(&format!("<c16:uniqueId val=\"{{0000000{}-6E8F-43DD-B1F6-30AC1D0140EF}}\"/>", actual_series_idx));
        xml.push_str("</c:ext></c:extLst>\n");
        
        xml.push_str("</c:ser>\n");
        actual_series_idx += 1;
    }
    
    if !chart.stacked && !chart.percent_stacked {
        write_data_labels(xml, chart.show_data_labels.unwrap_or(false));
    }
    
    xml.push_str("<c:gapWidth val=\"150\"/>\n");
    if chart.stacked || chart.percent_stacked {
        xml.push_str("<c:overlap val=\"100\"/>\n");
    }
    xml.push_str("<c:axId val=\"100000001\"/>\n");
    xml.push_str("<c:axId val=\"100000002\"/>\n");
    xml.push_str("</c:barChart>\n");
    
    xml.push_str("<c:catAx>\n");
    xml.push_str("<c:axId val=\"100000001\"/>\n");
    xml.push_str("<c:scaling><c:orientation val=\"minMax\"/></c:scaling>\n");
    xml.push_str("<c:delete val=\"0\"/>\n");
    xml.push_str("<c:axPos val=\"l\"/>\n");
    if let Some(ref x_title) = chart.x_axis_title {
        write_axis_title(xml, x_title, chart);
    }
    xml.push_str("<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>\n");
    xml.push_str("<c:majorTickMark val=\"none\"/>\n");
    xml.push_str("<c:minorTickMark val=\"none\"/>\n");
    xml.push_str("<c:tickLblPos val=\"nextTo\"/>\n");
    write_category_axis_styling(xml);
    xml.push_str("<c:crossAx val=\"100000002\"/>\n");
    xml.push_str("<c:crosses val=\"autoZero\"/>\n");
    xml.push_str("<c:auto val=\"1\"/>\n");
    xml.push_str("<c:lblAlgn val=\"ctr\"/>\n");
    xml.push_str("<c:lblOffset val=\"100\"/>\n");
    xml.push_str("<c:noMultiLvlLbl val=\"0\"/>\n");
    xml.push_str("</c:catAx>\n");
    
    xml.push_str("<c:valAx>\n");
    xml.push_str("<c:axId val=\"100000002\"/>\n");
    xml.push_str("<c:scaling>\n");
    xml.push_str("<c:orientation val=\"minMax\"/>\n");
    if let Some(min) = chart.axis_min {
        xml.push_str(&format!("<c:min val=\"{}\"/>\n", min));
    }
    if let Some(max) = chart.axis_max {
        xml.push_str(&format!("<c:max val=\"{}\"/>\n", max));
    }
    xml.push_str("</c:scaling>\n");
    xml.push_str("<c:delete val=\"0\"/>\n");
    xml.push_str("<c:axPos val=\"b\"/>\n");
    write_major_gridlines(xml);
    if let Some(ref y_title) = chart.y_axis_title {
        write_axis_title(xml, y_title, chart);
    }
    let format_code = if chart.percent_stacked { "0%" } else { "General" };
    xml.push_str(&format!("<c:numFmt formatCode=\"{}\" sourceLinked=\"1\"/>\n", format_code));
    xml.push_str("<c:majorTickMark val=\"none\"/>\n");
    xml.push_str("<c:minorTickMark val=\"none\"/>\n");
    xml.push_str("<c:tickLblPos val=\"nextTo\"/>\n");
    write_value_axis_styling(xml);
    xml.push_str("<c:crossAx val=\"100000001\"/>\n");
    xml.push_str("<c:crosses val=\"autoZero\"/>\n");
    xml.push_str("<c:crossBetween val=\"between\"/>\n");
    xml.push_str("</c:valAx>\n");
    xml.push_str("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>\n");
}

// ============================================================================
// LINE CHART
// ============================================================================
fn generate_line_chart_content(xml: &mut String, chart: &ExcelChart, sheet_name: &str) {
    xml.push_str("<c:lineChart>\n");
    xml.push_str(&format!("<c:grouping val=\"{}\"/>\n", 
        if chart.percent_stacked { "percentStacked" } else if chart.stacked { "stacked" } else { "standard" }));
    xml.push_str("<c:varyColors val=\"0\"/>\n");
    
    let (start_row, start_col, end_row, end_col) = chart.data_range;
    let category_col = chart.category_col.unwrap_or(start_col);
    let accent_colors = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
    let tint_shade_values = [("tint", "65000"), ("", ""), ("shade", "65000")];
    
    let mut actual_series_idx = 0;
    for col in start_col..=end_col {
        if Some(col) == chart.category_col {
            continue;
        }
        
        let series_name = chart.series_names.get(actual_series_idx).map(|s| s.as_str()).unwrap_or("Series");
        let accent_color = accent_colors[actual_series_idx % accent_colors.len()];
        let (modifier, value) = tint_shade_values[actual_series_idx % tint_shade_values.len()];
        
        xml.push_str(&format!("<c:ser>\n<c:idx val=\"{}\"/>\n<c:order val=\"{}\"/>\n", actual_series_idx, actual_series_idx));
        
        xml.push_str("<c:tx>\n<c:strRef>\n<c:f>");
        xml.push_str(&format!("{}!${}$1", sheet_name, get_column_letter(col)));
        xml.push_str("</c:f>\n<c:strCache>\n<c:ptCount val=\"1\"/>\n<c:pt idx=\"0\">\n");
        xml.push_str(&format!("<c:v>{}</c:v>\n", xml_escape_str(series_name)));
        xml.push_str("</c:pt>\n</c:strCache>\n</c:strRef>\n</c:tx>\n");
        
        xml.push_str("<c:spPr>\n");
        xml.push_str("<a:ln w=\"28575\" cap=\"rnd\">\n");
        xml.push_str(&format!("<a:solidFill><a:schemeClr val=\"{}\">", accent_color));
        if !modifier.is_empty() {
            xml.push_str(&format!("<a:{} val=\"{}\"/>", modifier, value));
        }
        xml.push_str("</a:schemeClr></a:solidFill>\n");
        xml.push_str("<a:round/></a:ln>\n");
        xml.push_str("<a:effectLst/>\n");
        xml.push_str("</c:spPr>\n");
        xml.push_str("<c:marker><c:symbol val=\"none\"/></c:marker>\n");
        
        if chart.stacked || chart.percent_stacked {
            write_data_labels(xml, chart.show_data_labels.unwrap_or(false));
        }
        
        xml.push_str("<c:cat>\n<c:strRef>\n<c:f>");
        xml.push_str(&format!("{}!${}${}:${}${}", 
            sheet_name, get_column_letter(category_col), start_row + 1, 
            get_column_letter(category_col), end_row + 1));
        xml.push_str("</c:f>\n</c:strRef>\n</c:cat>\n");
        
        xml.push_str("<c:val>\n<c:numRef>\n<c:f>");
        xml.push_str(&format!("{}!${}${}:${}${}", 
            sheet_name, get_column_letter(col), start_row + 1, 
            get_column_letter(col), end_row + 1));
        xml.push_str("</c:f>\n</c:numRef>\n</c:val>\n");
        
        xml.push_str("<c:smooth val=\"0\"/>\n");
        
        xml.push_str("<c:extLst><c:ext uri=\"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\">");
        xml.push_str(&format!("<c16:uniqueId val=\"{{0000000{}-6E8F-43DD-B1F6-30AC1D0140EF}}\"/>", actual_series_idx));
        xml.push_str("</c:ext></c:extLst>\n");
        
        xml.push_str("</c:ser>\n");
        actual_series_idx += 1;
    }
    
    if !chart.stacked && !chart.percent_stacked {
        write_data_labels(xml, chart.show_data_labels.unwrap_or(false));
    }
    xml.push_str("<c:smooth val=\"0\"/>\n");
    
    xml.push_str("<c:axId val=\"100000001\"/>\n");
    xml.push_str("<c:axId val=\"100000002\"/>\n");
    xml.push_str("</c:lineChart>\n");
    
    xml.push_str("<c:catAx>\n");
    xml.push_str("<c:axId val=\"100000001\"/>\n");
    xml.push_str("<c:scaling><c:orientation val=\"minMax\"/></c:scaling>\n");
    xml.push_str("<c:delete val=\"0\"/>\n");
    xml.push_str("<c:axPos val=\"b\"/>\n");
    if let Some(ref x_title) = chart.x_axis_title {
        write_axis_title(xml, x_title, chart);
    }
    xml.push_str("<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>\n");
    xml.push_str("<c:majorTickMark val=\"none\"/>\n");
    xml.push_str("<c:minorTickMark val=\"none\"/>\n");
    xml.push_str("<c:tickLblPos val=\"nextTo\"/>\n");
    write_category_axis_styling(xml);
    xml.push_str("<c:crossAx val=\"100000002\"/>\n");
    xml.push_str("<c:crosses val=\"autoZero\"/>\n");
    xml.push_str("<c:auto val=\"1\"/>\n");
    xml.push_str("<c:lblAlgn val=\"ctr\"/>\n");
    xml.push_str("<c:lblOffset val=\"100\"/>\n");
    xml.push_str("<c:noMultiLvlLbl val=\"0\"/>\n");
    xml.push_str("</c:catAx>\n");
    
    xml.push_str("<c:valAx>\n");
    xml.push_str("<c:axId val=\"100000002\"/>\n");
    xml.push_str("<c:scaling>\n");
    xml.push_str("<c:orientation val=\"minMax\"/>\n");
    if let Some(min) = chart.axis_min {
        xml.push_str(&format!("<c:min val=\"{}\"/>\n", min));
    }
    if let Some(max) = chart.axis_max {
        xml.push_str(&format!("<c:max val=\"{}\"/>\n", max));
    }
    xml.push_str("</c:scaling>\n");
    xml.push_str("<c:delete val=\"0\"/>\n");
    xml.push_str("<c:axPos val=\"l\"/>\n");
    write_major_gridlines(xml);
    if let Some(ref y_title) = chart.y_axis_title {
        write_axis_title(xml, y_title, chart);
    }
    let format_code = if chart.percent_stacked { "0%" } else { "General" };
    xml.push_str(&format!("<c:numFmt formatCode=\"{}\" sourceLinked=\"1\"/>\n", format_code));
    xml.push_str("<c:majorTickMark val=\"none\"/>\n");
    xml.push_str("<c:minorTickMark val=\"none\"/>\n");
    xml.push_str("<c:tickLblPos val=\"nextTo\"/>\n");
    write_value_axis_styling(xml);
    xml.push_str("<c:crossAx val=\"100000001\"/>\n");
    xml.push_str("<c:crosses val=\"autoZero\"/>\n");
    xml.push_str("<c:crossBetween val=\"between\"/>\n");
    xml.push_str("</c:valAx>\n");
    xml.push_str("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>\n");
}

fn generate_pie_chart_content(xml: &mut String, chart: &ExcelChart, sheet_name: &str) {
    xml.push_str("<c:pieChart>\n");
    xml.push_str("<c:varyColors val=\"1\"/>\n");
    
    let (start_row, start_col, end_row, _end_col) = chart.data_range;
    let category_col = chart.category_col.unwrap_or(start_col);
    
    // Pie charts typically show one series
    let data_col = if start_col == category_col { start_col + 1 } else { start_col };
    
    xml.push_str("<c:ser>\n<c:idx val=\"0\"/>\n<c:order val=\"0\"/>\n");
    
    xml.push_str("<c:cat>\n<c:strRef>\n<c:f>");
    xml.push_str(&format!("'{}'!${}${}:${}${}", 
        sheet_name, get_column_letter(category_col), start_row + 1, 
        get_column_letter(category_col), end_row + 1));
    xml.push_str("</c:f>\n</c:strRef>\n</c:cat>\n");
    
    xml.push_str("<c:val>\n<c:numRef>\n<c:f>");
    xml.push_str(&format!("'{}'!${}${}:${}${}", 
        sheet_name, get_column_letter(data_col), start_row + 1, 
        get_column_letter(data_col), end_row + 1));
    xml.push_str("</c:f>\n</c:numRef>\n</c:val>\n");
    
    xml.push_str("<c:extLst><c:ext uri=\"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\">");
    xml.push_str("<c16:uniqueId val=\"{00000000-6E8F-43DD-B1F6-30AC1D0140EF}\"/>");
    xml.push_str("</c:ext></c:extLst>\n");
    
    xml.push_str("</c:ser>\n");
    
    if chart.show_data_labels.unwrap_or(false) {
        xml.push_str("<c:dLbls>\n");
        xml.push_str("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>\n");
        xml.push_str("<c:txPr>\n");
        xml.push_str("<a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" lIns=\"38100\" tIns=\"19050\" rIns=\"38100\" bIns=\"19050\" anchor=\"ctr\" anchorCtr=\"1\"><a:spAutoFit/></a:bodyPr>\n");
        xml.push_str("<a:lstStyle/>\n");
        xml.push_str("<a:p><a:pPr>\n");
        xml.push_str("<a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n");
        xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"75000\"/><a:lumOff val=\"25000\"/></a:schemeClr></a:solidFill>\n");
        xml.push_str("<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\n");
        xml.push_str("</a:defRPr>\n");
        xml.push_str("</a:pPr><a:endParaRPr lang=\"en-US\"/></a:p>\n");
        xml.push_str("</c:txPr>\n");
        xml.push_str("<c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"1\"/><c:showBubbleSize val=\"0\"/>\n");
        xml.push_str("<c:showLeaderLines val=\"1\"/>\n");
        xml.push_str("</c:dLbls>\n");
    } else {
        xml.push_str("<c:dLbls><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"1\"/><c:showBubbleSize val=\"0\"/></c:dLbls>\n");
    }
    
    xml.push_str("</c:pieChart>\n");
}

fn generate_scatter_chart_content(xml: &mut String, chart: &ExcelChart, sheet_name: &str) {
    xml.push_str("<c:scatterChart>\n");
    xml.push_str("<c:scatterStyle val=\"lineMarker\"/>\n");
    
    let (start_row, start_col, end_row, end_col) = chart.data_range;
    let accent_colors = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
    let tint_shade_values = [("tint", "65000"), ("", ""), ("shade", "65000")];
    
    for (series_idx, col) in (start_col + 1..=end_col).enumerate() {
        let accent_color = accent_colors[series_idx % accent_colors.len()];
        let (modifier, value) = tint_shade_values[series_idx % tint_shade_values.len()];
        
        xml.push_str(&format!("<c:ser>\n<c:idx val=\"{}\"/>\n<c:order val=\"{}\"/>\n", series_idx, series_idx));
        
        xml.push_str("<c:spPr>\n");
        xml.push_str("<a:ln w=\"28575\" cap=\"rnd\">\n");
        xml.push_str(&format!("<a:solidFill><a:schemeClr val=\"{}\">", accent_color));
        if !modifier.is_empty() {
            xml.push_str(&format!("<a:{} val=\"{}\"/>", modifier, value));
        }
        xml.push_str("</a:schemeClr></a:solidFill>\n");
        xml.push_str("<a:round/></a:ln>\n");
        xml.push_str("<a:effectLst/>\n");
        xml.push_str("</c:spPr>\n");
        
        xml.push_str("<c:xVal>\n<c:numRef>\n<c:f>");
        xml.push_str(&format!("'{}'!${}${}:${}${}", 
            sheet_name, get_column_letter(start_col), start_row + 1, 
            get_column_letter(start_col), end_row + 1));
        xml.push_str("</c:f>\n</c:numRef>\n</c:xVal>\n");
        
        xml.push_str("<c:yVal>\n<c:numRef>\n<c:f>");
        xml.push_str(&format!("'{}'!${}${}:${}${}", 
            sheet_name, get_column_letter(col), start_row + 1, 
            get_column_letter(col), end_row + 1));
        xml.push_str("</c:f>\n</c:numRef>\n</c:yVal>\n");
        
        xml.push_str("<c:extLst><c:ext uri=\"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\">");
        xml.push_str(&format!("<c16:uniqueId val=\"{{0000000{}-6E8F-43DD-B1F6-30AC1D0140EF}}\"/>", series_idx));
        xml.push_str("</c:ext></c:extLst>\n");
        
        xml.push_str("</c:ser>\n");
    }
    
    write_data_labels(xml, chart.show_data_labels.unwrap_or(false));
    
    xml.push_str("<c:axId val=\"100000001\"/>\n");
    xml.push_str("<c:axId val=\"100000002\"/>\n");
    xml.push_str("</c:scatterChart>\n");
    
    xml.push_str("<c:valAx>\n");
    xml.push_str("<c:axId val=\"100000001\"/>\n");
    xml.push_str("<c:scaling>\n");
    xml.push_str("<c:orientation val=\"minMax\"/>\n");
    if let Some(min) = chart.axis_min {
        xml.push_str(&format!("<c:min val=\"{}\"/>\n", min));
    }
    if let Some(max) = chart.axis_max {
        xml.push_str(&format!("<c:max val=\"{}\"/>\n", max));
    }
    xml.push_str("</c:scaling>\n");
    xml.push_str("<c:delete val=\"0\"/>\n");
    xml.push_str("<c:axPos val=\"b\"/>\n");
    if let Some(ref x_title) = chart.x_axis_title {
        write_axis_title(xml, x_title, chart);
    }
    xml.push_str("<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>\n");
    xml.push_str("<c:majorTickMark val=\"none\"/>\n");
    xml.push_str("<c:minorTickMark val=\"none\"/>\n");
    xml.push_str("<c:tickLblPos val=\"nextTo\"/>\n");
    xml.push_str("<c:crossAx val=\"100000002\"/>\n");
    xml.push_str("<c:crosses val=\"autoZero\"/>\n");
    xml.push_str("</c:valAx>\n");
    
    xml.push_str("<c:valAx>\n");
    xml.push_str("<c:axId val=\"100000002\"/>\n");
    xml.push_str("<c:scaling>\n");
    xml.push_str("<c:orientation val=\"minMax\"/>\n");
    if let Some(min) = chart.axis_min {
        xml.push_str(&format!("<c:min val=\"{}\"/>\n", min));
    }
    if let Some(max) = chart.axis_max {
        xml.push_str(&format!("<c:max val=\"{}\"/>\n", max));
    }
    xml.push_str("</c:scaling>\n");
    xml.push_str("<c:delete val=\"0\"/>\n");
    xml.push_str("<c:axPos val=\"l\"/>\n");
    // EG_AxShared order requires majorGridlines BEFORE title (which is before
    // numFmt). Emitting the title first produced a schema-invalid axis that
    // Excel would try to repair. Order: axPos, majorGridlines, title, numFmt.
    xml.push_str("<c:majorGridlines/>\n");
    if let Some(ref y_title) = chart.y_axis_title {
        write_axis_title(xml, y_title, chart);
    }
    xml.push_str("<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>\n");
    xml.push_str("<c:majorTickMark val=\"none\"/>\n");
    xml.push_str("<c:minorTickMark val=\"none\"/>\n");
    xml.push_str("<c:tickLblPos val=\"nextTo\"/>\n");
    xml.push_str("<c:crossAx val=\"100000001\"/>\n");
    xml.push_str("<c:crosses val=\"autoZero\"/>\n");
    xml.push_str("</c:valAx>\n");
}
// ============================================================================
// AREA CHART
// ============================================================================
fn generate_area_chart_content(xml: &mut String, chart: &ExcelChart, sheet_name: &str) {
    xml.push_str("<c:areaChart>\n");
    xml.push_str(&format!("<c:grouping val=\"{}\"/>\n", 
        if chart.percent_stacked { "percentStacked" } else if chart.stacked { "stacked" } else { "standard" }));
    xml.push_str("<c:varyColors val=\"0\"/>\n");
    
    let (start_row, start_col, end_row, end_col) = chart.data_range;
    let category_col = chart.category_col.unwrap_or(start_col);
    let accent_colors = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
    let tint_shade_values = [("tint", "65000"), ("", ""), ("shade", "65000")];
    
    let mut actual_series_idx = 0;
    for col in start_col..=end_col {
        if Some(col) == chart.category_col {
            continue;
        }
        
        let series_name = chart.series_names.get(actual_series_idx).map(|s| s.as_str()).unwrap_or("Series");
        let accent_color = accent_colors[actual_series_idx % accent_colors.len()];
        let (modifier, value) = tint_shade_values[actual_series_idx % tint_shade_values.len()];
        
        xml.push_str(&format!("<c:ser>\n<c:idx val=\"{}\"/>\n<c:order val=\"{}\"/>\n", actual_series_idx, actual_series_idx));
        
        xml.push_str("<c:tx>\n<c:strRef>\n<c:f>");
        xml.push_str(&format!("{}!${}$1", sheet_name, get_column_letter(col)));
        xml.push_str("</c:f>\n<c:strCache>\n<c:ptCount val=\"1\"/>\n<c:pt idx=\"0\">\n");
        xml.push_str(&format!("<c:v>{}</c:v>\n", xml_escape_str(series_name)));
        xml.push_str("</c:pt>\n</c:strCache>\n</c:strRef>\n</c:tx>\n");
        
        xml.push_str("<c:spPr>\n");
        xml.push_str(&format!("<a:solidFill><a:schemeClr val=\"{}\">", accent_color));
        if !modifier.is_empty() {
            xml.push_str(&format!("<a:{} val=\"{}\"/>", modifier, value));
        }
        xml.push_str("</a:schemeClr></a:solidFill>\n");
        xml.push_str("<a:ln><a:noFill/></a:ln>\n");
        xml.push_str("<a:effectLst/>\n");
        xml.push_str("</c:spPr>\n");
        
        xml.push_str("<c:cat>\n<c:strRef>\n<c:f>");
        xml.push_str(&format!("{}!${}${}:${}${}", 
            sheet_name, get_column_letter(category_col), start_row + 1, 
            get_column_letter(category_col), end_row + 1));
        xml.push_str("</c:f>\n</c:strRef>\n</c:cat>\n");
        
        xml.push_str("<c:val>\n<c:numRef>\n<c:f>");
        xml.push_str(&format!("{}!${}${}:${}${}", 
            sheet_name, get_column_letter(col), start_row + 1, 
            get_column_letter(col), end_row + 1));
        xml.push_str("</c:f>\n</c:numRef>\n</c:val>\n");
        
        xml.push_str("<c:extLst><c:ext uri=\"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\">");
        xml.push_str(&format!("<c16:uniqueId val=\"{{0000000{}-6E8F-43DD-B1F6-30AC1D0140EF}}\"/>", actual_series_idx));
        xml.push_str("</c:ext></c:extLst>\n");
        
        xml.push_str("</c:ser>\n");
        actual_series_idx += 1;
    }
    
    // Area charts always have dLbls after all series
    write_data_labels(xml, chart.show_data_labels.unwrap_or(false));
    
    xml.push_str("<c:axId val=\"100000001\"/>\n");
    xml.push_str("<c:axId val=\"100000002\"/>\n");
    xml.push_str("</c:areaChart>\n");
    
    xml.push_str("<c:catAx>\n");
    xml.push_str("<c:axId val=\"100000001\"/>\n");
    xml.push_str("<c:scaling><c:orientation val=\"minMax\"/></c:scaling>\n");
    xml.push_str("<c:delete val=\"0\"/>\n");
    xml.push_str("<c:axPos val=\"b\"/>\n");
    if let Some(ref x_title) = chart.x_axis_title {
        write_axis_title(xml, x_title, chart);
    }
    xml.push_str("<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>\n");
    xml.push_str("<c:majorTickMark val=\"out\"/>\n");
    xml.push_str("<c:minorTickMark val=\"none\"/>\n");
    xml.push_str("<c:tickLblPos val=\"nextTo\"/>\n");
    write_category_axis_styling(xml);
    xml.push_str("<c:crossAx val=\"100000002\"/>\n");
    xml.push_str("<c:crosses val=\"autoZero\"/>\n");
    xml.push_str("<c:auto val=\"1\"/>\n");
    xml.push_str("<c:lblAlgn val=\"ctr\"/>\n");
    xml.push_str("<c:lblOffset val=\"100\"/>\n");
    xml.push_str("<c:noMultiLvlLbl val=\"0\"/>\n");
    xml.push_str("</c:catAx>\n");
    
    xml.push_str("<c:valAx>\n");
    xml.push_str("<c:axId val=\"100000002\"/>\n");
    xml.push_str("<c:scaling>\n");
    xml.push_str("<c:orientation val=\"minMax\"/>\n");
    if let Some(min) = chart.axis_min {
        xml.push_str(&format!("<c:min val=\"{}\"/>\n", min));
    }
    if let Some(max) = chart.axis_max {
        xml.push_str(&format!("<c:max val=\"{}\"/>\n", max));
    }
    xml.push_str("</c:scaling>\n");
    xml.push_str("<c:delete val=\"0\"/>\n");
    xml.push_str("<c:axPos val=\"l\"/>\n");
    write_major_gridlines(xml);
    if let Some(ref y_title) = chart.y_axis_title {
        write_axis_title(xml, y_title, chart);
    }
    let format_code = if chart.percent_stacked { "0%" } else { "General" };
    xml.push_str(&format!("<c:numFmt formatCode=\"{}\" sourceLinked=\"1\"/>\n", format_code));
    xml.push_str("<c:majorTickMark val=\"none\"/>\n");
    xml.push_str("<c:minorTickMark val=\"none\"/>\n");
    xml.push_str("<c:tickLblPos val=\"nextTo\"/>\n");
    write_value_axis_styling(xml);
    xml.push_str("<c:crossAx val=\"100000001\"/>\n");
    xml.push_str("<c:crosses val=\"autoZero\"/>\n");
    xml.push_str("<c:crossBetween val=\"midCat\"/>\n");
    xml.push_str("</c:valAx>\n");
    xml.push_str("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>\n");
}

/// Generate drawing relationships
#[allow(dead_code)]  // superseded by _ext variant
pub fn generate_drawing_rels(num_charts: usize) -> String {
    let mut xml = String::with_capacity(300 + num_charts * 150);
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    xml.push_str("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
    
    for i in 1..=num_charts {
        xml.push_str(&format!("<Relationship Id=\"rIdChart{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart{}.xml\"/>\n", i, i));
    }
    
    xml.push_str("</Relationships>");
    xml
}


/// A column whose concrete Arrow array type has been resolved once, so the
/// per-cell hot path can skip `match data_type()` + `downcast_ref` on every
/// single cell. Only the high-volume primitive/string/bool types are given a
/// dedicated variant; everything else (dates, timestamps, decimals, ...) maps to
/// `Other` and takes the original generic per-cell path. The borrowed concrete
/// arrays live as long as the batch being written.
enum ColumnView<'a> {
    Utf8(&'a arrow_array::StringArray),
    LargeUtf8(&'a arrow_array::LargeStringArray),
    Bool(&'a arrow_array::BooleanArray),
    I8(&'a arrow_array::Int8Array),
    I16(&'a arrow_array::Int16Array),
    I32(&'a arrow_array::Int32Array),
    I64(&'a arrow_array::Int64Array),
    U8(&'a arrow_array::UInt8Array),
    U16(&'a arrow_array::UInt16Array),
    U32(&'a arrow_array::UInt32Array),
    U64(&'a arrow_array::UInt64Array),
    F32(&'a arrow_array::Float32Array),
    F64(&'a arrow_array::Float64Array),
    /// Any type without a fast variant; handled by the generic per-cell writer.
    Other,
}

impl<'a> ColumnView<'a> {
    /// Resolve a column's concrete type once. Unknown/rare types become `Other`.
    #[inline]
    fn resolve(array: &'a dyn arrow_array::Array) -> ColumnView<'a> {
        use arrow_array::*;
        match array.data_type() {
            DataType::Utf8 => ColumnView::Utf8(array.as_any().downcast_ref::<StringArray>().unwrap()),
            DataType::LargeUtf8 => ColumnView::LargeUtf8(array.as_any().downcast_ref::<LargeStringArray>().unwrap()),
            DataType::Boolean => ColumnView::Bool(array.as_any().downcast_ref::<BooleanArray>().unwrap()),
            DataType::Int8 => ColumnView::I8(array.as_any().downcast_ref::<Int8Array>().unwrap()),
            DataType::Int16 => ColumnView::I16(array.as_any().downcast_ref::<Int16Array>().unwrap()),
            DataType::Int32 => ColumnView::I32(array.as_any().downcast_ref::<Int32Array>().unwrap()),
            DataType::Int64 => ColumnView::I64(array.as_any().downcast_ref::<Int64Array>().unwrap()),
            DataType::UInt8 => ColumnView::U8(array.as_any().downcast_ref::<UInt8Array>().unwrap()),
            DataType::UInt16 => ColumnView::U16(array.as_any().downcast_ref::<UInt16Array>().unwrap()),
            DataType::UInt32 => ColumnView::U32(array.as_any().downcast_ref::<UInt32Array>().unwrap()),
            DataType::UInt64 => ColumnView::U64(array.as_any().downcast_ref::<UInt64Array>().unwrap()),
            DataType::Float32 => ColumnView::F32(array.as_any().downcast_ref::<Float32Array>().unwrap()),
            DataType::Float64 => ColumnView::F64(array.as_any().downcast_ref::<Float64Array>().unwrap()),
            _ => ColumnView::Other,
        }
    }

    /// Write cell `row_idx` with no style/overlay. Returns `true` if it handled
    /// the cell, `false` for `Other` (caller uses the generic path). Nulls are
    /// written as empty cells to match the generic path's behavior.
    #[inline]
    fn write_fast(
        &self,
        row_idx: usize,
        cell_ref: &[u8],
        buf: &mut Vec<u8>,
        ryu_buf: &mut zmij::Buffer,
        int_buf: &mut itoa::Buffer,
    ) -> bool {
        use arrow_array::Array;
        match self {
            ColumnView::Utf8(a) => {
                if a.is_null(row_idx) { return true; }
                let offsets = a.offsets();
                let values = a.values();
                let start = offsets[row_idx] as usize;
                let end = offsets[row_idx + 1] as usize;
                let str_bytes = &values.as_ref()[start..end];
                if str_bytes.is_empty() { return true; }
                buf.extend_from_slice(b"<c r=\"");
                buf.extend_from_slice(cell_ref);
                buf.extend_from_slice(b"\" t=\"inlineStr\">");
                write_inline_string(str_bytes, buf);
                buf.extend_from_slice(b"</c>");
            }
            ColumnView::LargeUtf8(a) => {
                if a.is_null(row_idx) { return true; }
                let offsets = a.offsets();
                let values = a.values();
                let start = offsets[row_idx] as usize;
                let end = offsets[row_idx + 1] as usize;
                let str_bytes = &values.as_ref()[start..end];
                if str_bytes.is_empty() { return true; }
                buf.extend_from_slice(b"<c r=\"");
                buf.extend_from_slice(cell_ref);
                buf.extend_from_slice(b"\" t=\"inlineStr\">");
                write_inline_string(str_bytes, buf);
                buf.extend_from_slice(b"</c>");
            }
            ColumnView::Bool(a) => {
                if a.is_null(row_idx) { return true; }
                // Boolean cells are t="b" with <v>0</v> or <v>1</v>.
                buf.extend_from_slice(b"<c r=\"");
                buf.extend_from_slice(cell_ref);
                buf.extend_from_slice(b"\" t=\"b\"><v>");
                buf.push(if a.value(row_idx) { b'1' } else { b'0' });
                buf.extend_from_slice(b"</v></c>");
            }
            ColumnView::I8(a)  => { if a.is_null(row_idx) { return true; } write_number_cell_int(a.value(row_idx) as i64, cell_ref, None, buf, int_buf); }
            ColumnView::I16(a) => { if a.is_null(row_idx) { return true; } write_number_cell_int(a.value(row_idx) as i64, cell_ref, None, buf, int_buf); }
            ColumnView::I32(a) => { if a.is_null(row_idx) { return true; } write_number_cell_int(a.value(row_idx) as i64, cell_ref, None, buf, int_buf); }
            ColumnView::I64(a) => { if a.is_null(row_idx) { return true; } write_number_cell_int(a.value(row_idx), cell_ref, None, buf, int_buf); }
            ColumnView::U8(a)  => { if a.is_null(row_idx) { return true; } write_number_cell_int(a.value(row_idx) as i64, cell_ref, None, buf, int_buf); }
            ColumnView::U16(a) => { if a.is_null(row_idx) { return true; } write_number_cell_int(a.value(row_idx) as i64, cell_ref, None, buf, int_buf); }
            ColumnView::U32(a) => { if a.is_null(row_idx) { return true; } write_number_cell_int(a.value(row_idx) as i64, cell_ref, None, buf, int_buf); }
            ColumnView::U64(a) => {
                if a.is_null(row_idx) { return true; }
                let v = a.value(row_idx);
                // u64 values above i64::MAX wrap to negative when cast `as i64`,
                // silently corrupting large IDs/counts. Excel stores every number
                // as an f64 anyway, so emit large values through the float path
                // (exact up to 2^53, same precision Excel itself provides) while
                // keeping the fast integer path for the common in-range case.
                if v <= i64::MAX as u64 {
                    write_number_cell_int(v as i64, cell_ref, None, buf, int_buf);
                } else {
                    write_number_cell(v as f64, cell_ref, None, buf, ryu_buf, int_buf);
                }
            }
            ColumnView::F32(a) => { if a.is_null(row_idx) { return true; } write_number_cell(a.value(row_idx) as f64, cell_ref, None, buf, ryu_buf, int_buf); }
            ColumnView::F64(a) => { if a.is_null(row_idx) { return true; } write_number_cell(a.value(row_idx), cell_ref, None, buf, ryu_buf, int_buf); }
            ColumnView::Other => return false,
        }
        true
    }
}

/// Generate complete sheet XML with all enhanced features
/// Element order: dimension → sheetViews → sheetFormatPr → cols → sheetData → 
///                autoFilter → mergeCells → conditionalFormatting → dataValidations → 
///                hyperlinks → drawing → tableParts
/// Write a single data row (`<row>…</row>`) into `buf`. Extracted so the serial
/// and the chunked-parallel row generators emit byte-for-byte identical output --
/// they both call exactly this. Everything it reads is either passed by value
/// (the row number / row index) or borrowed read-only (arrays, maps, config), so
/// it is safe to run concurrently on disjoint row ranges into disjoint buffers.
#[allow(clippy::too_many_arguments)]
#[inline(always)]
fn write_one_data_row(
    row_num: usize,
    row_idx: usize,
    batch: &RecordBatch,
    col_views: &[ColumnView],
    num_cols: usize,
    col_letters: &[([u8; 4], usize)],
    row_spans: &[u8],
    config: &StyleConfig,
    col_format_map: &HashMap<usize, u32>,
    cell_style_map: &HashMap<(usize, usize), u32>,
    hyperlink_map: &HashMap<(usize, usize), &Hyperlink>,
    formula_map: &HashMap<(usize, usize), &Formula>,
    sst: &SharedStrings,
    col_is_shared: &[bool],
    any_shared: bool,
    has_overlays: bool,
    has_row_heights: bool,
    has_hidden_rows: bool,
    buf: &mut Vec<u8>,
    ryu_buf: &mut zmij::Buffer,
    cell_int_buf: &mut itoa::Buffer,
) -> Result<(), WriteError> {
    let mut int_buf = itoa::Buffer::new();
    let mut cell_ref = [0u8; 16];

    let row_str = int_buf.format(row_num);
    let row_bytes = row_str.as_bytes();

    buf.extend_from_slice(b"<row r=\"");
    buf.extend_from_slice(row_bytes);
    buf.push(b'\"');
    buf.extend_from_slice(row_spans);

    if has_row_heights {
        if let Some(height) = config.row_heights.as_ref().unwrap().get(&row_num) {
            buf.extend_from_slice(b" ht=\"");
            buf.extend_from_slice(zmij::Buffer::new().format(*height).as_bytes());
            buf.extend_from_slice(b"\" customHeight=\"1\"");
        }
    }

    if has_hidden_rows && config.hidden_rows.contains(&row_num) {
        buf.extend_from_slice(b" hidden=\"1\"");
    }

    buf.push(b'>');

    for col_idx in 0..num_cols {
        let array = batch.column(col_idx);
        let (col_letter, col_len) = &col_letters[col_idx];

        let cell_ref_len = {
            cell_ref[..*col_len].copy_from_slice(&col_letter[..*col_len]);
            cell_ref[*col_len..*col_len + row_bytes.len()].copy_from_slice(row_bytes);
            *col_len + row_bytes.len()
        };
        let cell_ref_slice = &cell_ref[..cell_ref_len];

        let custom_style_id = cell_style_map.get(&(row_num, col_idx)).copied();
        let default_style_id = col_format_map.get(&col_idx).copied();
        let style_id = custom_style_id.or(default_style_id);

        let hyperlink = hyperlink_map.get(&(row_num, col_idx));
        let formula = formula_map.get(&(row_num, col_idx));

        let col_shared = any_shared && col_is_shared[col_idx];

        let cv = &col_views[col_idx];
        if !has_overlays && !col_shared && cv.write_fast(row_idx, cell_ref_slice, buf, ryu_buf, cell_int_buf) {
            continue;
        }

        write_arrow_cell_to_xml_optimized(
            array.as_ref(),
            row_idx,
            cell_ref_slice,
            style_id,
            hyperlink,
            formula,
            buf,
            ryu_buf,
            cell_int_buf,
            if col_shared { Some(sst) } else { None },
        )?;
    }

    buf.extend_from_slice(b"</row>");
    Ok(())
}

pub fn generate_sheet_xml_from_arrow(
    batches: &[RecordBatch],
    config: &StyleConfig,
    col_format_map: &HashMap<usize, u32>,
    cell_style_map: &HashMap<(usize, usize), u32>,
    sst: &SharedStrings,
    shared_cols: &[bool],
) -> Result<Vec<u8>, WriteError> {
    if batches.is_empty() {
        return Ok(b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
<dimension ref=\"A1\"/><sheetData/></worksheet>".to_vec());
    }

    let schema = batches[0].schema();
    let num_cols = schema.fields().len();
    let total_rows: usize = batches.iter().map(|b| b.num_rows()).sum();

    if num_cols == 0 {
        return Ok(b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
<dimension ref=\"A1\"/><sheetData/></worksheet>".to_vec());
    }

    // Reject unsupported column types up front (once), so the per-cell hot path
    // can keep its infallible downcasts and never panic on a surprise type.
    validate_arrow_schema(&schema)?;

    // Determine where DataFrame data actually starts. This must be computed
    // BEFORE the <dimension> element so the dimension's row extent matches the
    // rows we actually emit. It depends only on config, not on the data.
    let data_start = if config.write_header_row {
        config.data_start_row.max(1)
    } else {
        // Excel/OOXML rows are 1-based. Without a header row the first data row
        // must still land on row 1 (or the caller-requested data_start_row if
        // they deliberately offset it), never row 0 -- a `<row r="0">` is
        // invalid and readers silently drop it, losing the first data row.
        config.data_start_row.max(1)
    };

    // Build map of table header rows that need to be inserted. A header row is
    // inserted only when a table starts strictly below the DataFrame's own
    // header (start_row > data_start); a table anchored at data_start reuses the
    // existing header. This is the SAME predicate used later when the rows are
    // actually written, so `num_inserted_headers` (which feeds <dimension>) can
    // never disagree with the real output. The previous code computed this twice
    // with two different predicates (`> 1` here, `> data_start` below), so a
    // custom data_start_row produced a dimension that overstated the row count.
    let mut table_header_rows: HashMap<usize, (usize, usize)> = HashMap::new();
    let mut num_inserted_headers = 0;
    for table in &config.tables {
        let (start_row, start_col, _, end_col) = table.range;
        if start_row > data_start {
            table_header_rows.insert(start_row, (start_col, end_col));
            num_inserted_headers += 1;
        }
    }

    let exact_size = calculate_exact_xml_size(batches)?;
    let mut buf = Vec::with_capacity(exact_size);

    buf.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

    // SheetPr (tab color - must come before dimension)
    if let Some(ref color) = config.tab_color {
        buf.extend_from_slice(b"<sheetPr><tabColor rgb=\"");
        // User-supplied colors are never validated upstream, so an stray `&`,
        // `"` or `<` would corrupt the whole worksheet. Escape defensively.
        xml_escape_simd(color.as_bytes(), &mut buf);
        buf.extend_from_slice(b"\"/></sheetPr>");
    }

    // Dimension
    buf.extend_from_slice(b"<dimension ref=\"");
    if total_rows > 0 {
        buf.extend_from_slice(b"A1:");
        let mut col_buf = [0u8; 4];
        let col_len = write_col_letter(num_cols - 1, &mut col_buf);
        buf.extend_from_slice(&col_buf[..col_len]);
        
        let mut row_buf = itoa::Buffer::new();
        buf.extend_from_slice(row_buf.format(total_rows + 1 + num_inserted_headers).as_bytes());
    } else {
        buf.extend_from_slice(b"A1");
    }
    buf.extend_from_slice(b"\"/>");

    // SheetViews (with gridlines, zoom, RTL, and optional freeze panes)
    buf.extend_from_slice(b"<sheetViews><sheetView workbookViewId=\"0\"");
    
    // Add showGridLines if disabled
    if !config.gridlines_visible {
        buf.extend_from_slice(b" showGridLines=\"0\"");
    }
    
    // Add zoom scale
    if let Some(zoom) = config.zoom_scale {
        buf.extend_from_slice(b" zoomScale=\"");
        buf.extend_from_slice(itoa::Buffer::new().format(zoom).as_bytes());
        buf.push(b'\"');
    }
    
    // Add right-to-left
    if config.right_to_left {
        buf.extend_from_slice(b" rightToLeft=\"1\"");
    }
    
    if config.freeze_rows > 0 || config.freeze_cols > 0 {
        buf.push(b'>');
        buf.extend_from_slice(b"<pane ");
        
        if config.freeze_cols > 0 {
            buf.extend_from_slice(b"xSplit=\"");
            buf.extend_from_slice(itoa::Buffer::new().format(config.freeze_cols).as_bytes());
            buf.extend_from_slice(b"\" ");
        }
        
        if config.freeze_rows > 0 {
            buf.extend_from_slice(b"ySplit=\"");
            buf.extend_from_slice(itoa::Buffer::new().format(config.freeze_rows).as_bytes());
            buf.extend_from_slice(b"\" ");
        }
        
        buf.extend_from_slice(b"topLeftCell=\"");
        write_cell_ref(config.freeze_cols, config.freeze_rows + 1, &mut buf);
        buf.extend_from_slice(b"\" activePane=\"bottomRight\" state=\"frozen\"/>");
        buf.extend_from_slice(b"</sheetView></sheetViews>");
    } else {
        buf.extend_from_slice(b"/></sheetViews>");
    }

    // SheetFormatPr (default row height)
    buf.extend_from_slice(b"<sheetFormatPr defaultRowHeight=\"");
    let default_height = config.default_row_height.unwrap_or(15.0);
    buf.extend_from_slice(zmij::Buffer::new().format(default_height).as_bytes());
    buf.push(b'\"');
    if config.default_row_height.is_some() {
        buf.extend_from_slice(b" customHeight=\"1\"");
    }
    buf.extend_from_slice(b"/>");

    // Cols (column widths and hidden columns)
    if config.auto_width || config.column_widths.is_some() || !config.hidden_columns.is_empty() {
        buf.extend_from_slice(b"<cols>");
        
        for (col_idx, field) in schema.fields().iter().enumerate() {
            let width = if let Some(widths) = &config.column_widths {
                if let Some(col_width) = widths.get(field.name()) {
                    match col_width {
                        ColumnWidth::Characters(w) => *w,
                        ColumnWidth::Pixels(px) => px / 7.0,  // Calibri 11pt MDW
                        ColumnWidth::Auto => calculate_column_width(
                            batches[0].column(col_idx).as_ref(),
                            field.name(), 100, config.data_start_row
                        ),
                    }
                } else if config.auto_width {
                    calculate_column_width(batches[0].column(col_idx).as_ref(),
                                        field.name(), 100, config.data_start_row)
                } else {
                    8.43
                }
            } else if config.auto_width {
                calculate_column_width(batches[0].column(col_idx).as_ref(),
                                    field.name(), 100, config.data_start_row)
            } else {
                8.43
            };
            
            buf.extend_from_slice(b"<col min=\"");
            buf.extend_from_slice(itoa::Buffer::new().format(col_idx + 1).as_bytes());
            buf.extend_from_slice(b"\" max=\"");
            buf.extend_from_slice(itoa::Buffer::new().format(col_idx + 1).as_bytes());
            buf.extend_from_slice(b"\" width=\"");
            buf.extend_from_slice(zmij::Buffer::new().format(width).as_bytes());
            buf.extend_from_slice(b"\" customWidth=\"1\"");
            
            // Hidden column
            if config.hidden_columns.contains(&col_idx) {
                buf.extend_from_slice(b" hidden=\"1\"");
            }
            
            buf.extend_from_slice(b"/>");
        }
        
        buf.extend_from_slice(b"</cols>");
    }


    // SheetData (all cell data)
    buf.extend_from_slice(b"<sheetData>");

    let col_letters: Vec<([u8; 4], usize)> = (0..num_cols)
        .map(|i| {
            let mut col_buf = [0u8; 4];
            let len = write_col_letter(i, &mut col_buf);
            (col_buf, len)
        })
        .collect();

    let mut ryu_buf = zmij::Buffer::new();
    let mut int_buf = itoa::Buffer::new();
    let mut cell_int_buf = itoa::Buffer::new();

    // Precompute, per column, whether it is served by the shared-string table.
    // A fixed-length Vec indexed by col_idx removes a bounds-checked Option probe
    // from the per-cell path. When nothing is shared (the common case, and always
    // for all-unique strings), `any_shared` is false and the writer takes the
    // original inline path with zero added work.
    let any_shared = !sst.is_empty() && shared_cols.iter().any(|&b| b);
    let col_is_shared: Vec<bool> = (0..num_cols)
        .map(|i| shared_cols.get(i).copied().unwrap_or(false))
        .collect();

    let hyperlink_map: HashMap<(usize, usize), &Hyperlink> = config.hyperlinks
        .iter()
        .map(|h| ((h.row, h.col), h))
        .collect();
    
    let formula_map: HashMap<(usize, usize), &Formula> = config.formulas
        .iter()
        .map(|f| ((f.row, f.col), f))
        .collect();

    // (data_start is computed earlier, before <dimension>, so the dimension row
    // count and the emitted rows stay consistent.)

    // Write header_content rows (arbitrary content before DataFrame data)
    if !config.header_content.is_empty() {
        let mut rows_map: HashMap<usize, Vec<(usize, String)>> = HashMap::new();
        for (row, col, text) in &config.header_content {
            rows_map.entry(*row).or_insert_with(Vec::new).push((*col, text.clone()));
        }
        
        let mut sorted_rows: Vec<_> = rows_map.keys().copied().collect();
        sorted_rows.sort();
        
        for row_num in sorted_rows {
            if row_num >= data_start { break; }
            
            let row_str = int_buf.format(row_num);
            let row_bytes = row_str.as_bytes();
            
            // Start row tag with row number
            buf.extend_from_slice(b"<row r=\"");
            buf.extend_from_slice(row_bytes);
            buf.push(b'\"');  // CRITICAL: Always close the r attribute
            
            // Add optional row height
            if let Some(heights) = &config.row_heights {
                if let Some(height) = heights.get(&row_num) {
                    buf.extend_from_slice(b" ht=\"");  // Note: leading space for separate attribute
                    buf.extend_from_slice(zmij::Buffer::new().format(*height).as_bytes());
                    buf.extend_from_slice(b"\" customHeight=\"1\"");
                }
            }
            
            // Add hidden attribute if needed
            if config.hidden_rows.contains(&row_num) {
                buf.extend_from_slice(b" hidden=\"1\"");
            }
            
            buf.push(b'>');  // Close row opening tag
            
            // Write cells in this row
            if let Some(cells) = rows_map.get(&row_num) {
                for (col_idx, text) in cells {
                    let (col_letter, col_len) = &col_letters[*col_idx];
                    
                    // Cell reference (e.g., "A2")
                    buf.extend_from_slice(b"<c r=\"");
                    buf.extend_from_slice(&col_letter[..*col_len]);
                    buf.extend_from_slice(row_bytes);
                    buf.push(b'\"');  // Close r attribute
                    
                    // Apply custom cell style if defined
                    if let Some(style_id) = cell_style_map.get(&(row_num, *col_idx)) {
                        buf.extend_from_slice(b" s=\"");
                        buf.extend_from_slice(itoa::Buffer::new().format(*style_id).as_bytes());
                        buf.push(b'\"');  // Close s attribute
                    }
                    
                    // Write inline string content
                    buf.extend_from_slice(b" t=\"inlineStr\">");
                    write_inline_string(text.as_bytes(), &mut buf);
                    buf.extend_from_slice(b"</c>");
                }
            }
            
            buf.extend_from_slice(b"</row>");
        }
    }

    // OOXML row `spans` optimization (Open XML Explained, markup sample 78):
    // Excel emits `spans="1:N"` on each row to tell the consumer, up front, the
    // column extent of that row. A reader can then pre-size its in-memory row
    // structures in one shot instead of growing them cell-by-cell as it parses.
    // Our data rows are dense across all `num_cols` columns, so the span is a
    // single constant we compute ONCE here and reuse for every row -- no
    // per-cell or per-row cost beyond appending a few cached bytes. Columns are
    // 1-based in the spans attribute, hence "1:num_cols".
    let row_spans: Vec<u8> = {
        let mut s = Vec::with_capacity(16);
        s.extend_from_slice(b" spans=\"1:");
        s.extend_from_slice(itoa::Buffer::new().format(num_cols).as_bytes());
        s.push(b'\"');
        s
    };

    // Write DataFrame header row at data_start (only if enabled)
    if config.write_header_row {
        let header_row_height = config.row_heights.as_ref().and_then(|h| h.get(&data_start));
        buf.extend_from_slice(b"<row r=\"");
        buf.extend_from_slice(itoa::Buffer::new().format(data_start).as_bytes());
        buf.push(b'\"');
        // Header row spans all columns too.
        buf.extend_from_slice(&row_spans);
        if let Some(height) = header_row_height {
            buf.extend_from_slice(b" ht=\"");
            buf.extend_from_slice(zmij::Buffer::new().format(*height).as_bytes());
            buf.extend_from_slice(b"\" customHeight=\"1\"");
        }
        // Hidden row check for header
        if config.hidden_rows.contains(&data_start) {
            buf.extend_from_slice(b" hidden=\"1\"");
        }
        buf.push(b'>');
        
        for (col_idx, field) in schema.fields().iter().enumerate() {
            let (col_letter, col_len) = &col_letters[col_idx];
            
            let style_id = if config.styled_headers { 2 } else { 0 };
            
            buf.extend_from_slice(b"<c r=\"");
            buf.extend_from_slice(&col_letter[..*col_len]);
            buf.extend_from_slice(itoa::Buffer::new().format(data_start).as_bytes());
            if style_id > 0 {
                buf.extend_from_slice(b"\" s=\"");
                buf.extend_from_slice(int_buf.format(style_id).as_bytes());
            }
            buf.extend_from_slice(b"\" t=\"inlineStr\">");
            write_inline_string(field.name().as_bytes(), &mut buf);
            buf.extend_from_slice(b"</c>");
        }
        buf.extend_from_slice(b"</row>");
    }

    let mut current_row = if config.write_header_row { data_start + 1 } else { data_start };

    // `table_header_rows` was already built once (before <dimension>) using the
    // correct `start_row > data_start` predicate, so we reuse it directly here
    // instead of recomputing it with a divergent rule.

    // Cache feature flags to avoid repeated checks
    let has_table_headers = !table_header_rows.is_empty();
    let has_row_heights = config.row_heights.is_some();
    let has_hidden_rows = !config.hidden_rows.is_empty();

    // Whether any per-cell overlay (styles/formats/hyperlinks/formulas) exists.
    // When none do -- the overwhelmingly common case for bulk data -- the cell
    // writer can take a branch-free fast path.
    let has_overlays = !cell_style_map.is_empty()
        || !col_format_map.is_empty()
        || !hyperlink_map.is_empty()
        || !formula_map.is_empty();

    // Total data rows across all batches.
    let total_data_rows: usize = batches.iter().map(|b| b.num_rows()).sum();

    // The data-row region starts here when there are no table-header insertions.
    let data_row_start = current_row;

    // Decide whether to generate rows in parallel. This is a pure speed
    // optimization for the common bulk-data case; it is only taken when it cannot
    // change a single output byte:
    //   * no table-header insertions (those shift row numbers non-linearly),
    //   * the sheet is large enough that thread spawn/join pays off,
    //   * more than one worker thread is actually available.
    // In every other case we fall through to the serial loop below, unchanged.
    const PARALLEL_ROW_THRESHOLD: usize = 50_000;
    let nthreads = rayon::current_num_threads();
    let use_parallel = !has_table_headers
        && total_data_rows >= PARALLEL_ROW_THRESHOLD
        && nthreads > 1;

    if use_parallel {
        use rayon::prelude::*;

        // Build a flat index of every data row as (batch_idx, row_in_batch) so a
        // chunk can be described by a contiguous slice of global row positions.
        // With no table headers, global row r maps to sheet row data_row_start + r.
        let mut batch_offsets = Vec::with_capacity(batches.len() + 1);
        let mut acc = 0usize;
        batch_offsets.push(0usize);
        for b in batches {
            acc += b.num_rows();
            batch_offsets.push(acc);
        }

        // Pre-resolve each batch's column views once (shared read-only across
        // threads). ColumnView borrows the Arc<dyn Array>, which is Sync.
        let per_batch_cols: Vec<Vec<ColumnView>> = batches
            .iter()
            .map(|batch| {
                (0..num_cols)
                    .map(|c| ColumnView::resolve(batch.column(c).as_ref()))
                    .collect()
            })
            .collect();

        // Split the global row range into `nthreads` contiguous chunks. Each chunk
        // produces its own buffer; chunks are concatenated in order afterwards so
        // the byte stream is identical to the serial path and fully deterministic.
        let chunk_size = total_data_rows.div_ceil(nthreads);
        let chunk_starts: Vec<usize> = (0..total_data_rows).step_by(chunk_size).collect();

        let chunks: Result<Vec<Vec<u8>>, WriteError> = chunk_starts
            .par_iter()
            .map(|&chunk_start| {
                let chunk_end = (chunk_start + chunk_size).min(total_data_rows);
                // Estimate ~48 bytes/cell to size the buffer and avoid regrowth.
                let mut local = Vec::with_capacity((chunk_end - chunk_start) * num_cols * 48 + 64);
                let mut ryu_buf = zmij::Buffer::new();
                let mut cell_int_buf = itoa::Buffer::new();

                // Which batch does chunk_start fall in? Walk batch offsets.
                let mut bi = match batch_offsets.binary_search(&chunk_start) {
                    Ok(i) => i,
                    Err(i) => i - 1,
                };
                for global_r in chunk_start..chunk_end {
                    while global_r >= batch_offsets[bi + 1] {
                        bi += 1;
                    }
                    let row_idx = global_r - batch_offsets[bi];
                    let row_num = data_row_start + global_r;
                    write_one_data_row(
                        row_num, row_idx, &batches[bi], &per_batch_cols[bi], num_cols,
                        &col_letters, &row_spans, config, col_format_map, cell_style_map,
                        &hyperlink_map, &formula_map, sst, &col_is_shared, any_shared,
                        has_overlays, has_row_heights, has_hidden_rows,
                        &mut local, &mut ryu_buf, &mut cell_int_buf,
                    )?;
                }
                Ok(local)
            })
            .collect();

        for chunk in chunks? {
            buf.extend_from_slice(&chunk);
        }
    } else {
        // Serial path (fallback): identical output, used for small sheets, when
        // table headers shift row numbers, or on a single-threaded pool.
        for batch in batches {
            let batch_rows = batch.num_rows();

            // Resolve each column's concrete Arrow type ONCE per batch.
            let cols: Vec<_> = (0..num_cols).map(|c| batch.column(c).clone()).collect();
            let col_views: Vec<ColumnView> = cols
                .iter()
                .map(|c| ColumnView::resolve(c.as_ref()))
                .collect();

            for row_idx in 0..batch_rows {
                // Check if we need to insert table header row before this data row
                if has_table_headers {
                    if let Some(&(start_col, end_col)) = table_header_rows.get(&current_row) {
                        let row_str = int_buf.format(current_row);
                        let row_bytes = row_str.as_bytes();

                        buf.extend_from_slice(b"<row r=\"");
                        buf.extend_from_slice(row_bytes);
                        buf.push(b'\"');

                        if has_row_heights {
                            if let Some(height) = config.row_heights.as_ref().unwrap().get(&current_row) {
                                buf.extend_from_slice(b" ht=\"");
                                buf.extend_from_slice(zmij::Buffer::new().format(*height).as_bytes());
                                buf.extend_from_slice(b"\" customHeight=\"1\"");
                            }
                        }

                        if has_hidden_rows && config.hidden_rows.contains(&current_row) {
                            buf.extend_from_slice(b" hidden=\"1\"");
                        }

                        buf.push(b'>');

                        // Write header cells for table columns
                        for col_idx in start_col..=end_col {
                            let (col_letter, col_len) = &col_letters[col_idx];
                            let field_name = schema.fields()[col_idx].name();

                            let mut header_cell_ref = Vec::with_capacity(16);
                            header_cell_ref.extend_from_slice(&col_letter[..*col_len]);
                            header_cell_ref.extend_from_slice(row_bytes);

                            let custom_style_id = cell_style_map.get(&(current_row, col_idx)).copied();

                            buf.extend_from_slice(b"<c r=\"");
                            buf.extend_from_slice(&header_cell_ref);
                            if let Some(sid) = custom_style_id {
                                buf.extend_from_slice(b"\" s=\"");
                                buf.extend_from_slice(itoa::Buffer::new().format(sid).as_bytes());
                            }
                            buf.extend_from_slice(b"\" t=\"inlineStr\">");
                            write_inline_string(field_name.as_bytes(), &mut buf);
                            buf.extend_from_slice(b"</c>");
                        }

                        buf.extend_from_slice(b"</row>");
                        current_row += 1;
                    }
                }

                write_one_data_row(
                    current_row, row_idx, batch, &col_views, num_cols,
                    &col_letters, &row_spans, config, col_format_map, cell_style_map,
                    &hyperlink_map, &formula_map, sst, &col_is_shared, any_shared,
                    has_overlays, has_row_heights, has_hidden_rows,
                    &mut buf, &mut ryu_buf, &mut cell_int_buf,
                )?;
                current_row += 1;
            }
        }
    }

    buf.extend_from_slice(b"</sheetData>");

    // AutoFilter - only if no table covers the entire range from A1
    let has_full_table = config.tables.iter().any(|t| {
        let (start_row, start_col, end_row, end_col) = t.range;
        start_row == 1 && start_col == 0 && end_row >= total_rows && end_col >= num_cols - 1
    });
    // AutoFilter
    if config.auto_filter && total_rows > 0 && !has_full_table {
        buf.extend_from_slice(b"<autoFilter ref=\"A1:");
        let mut col_buf = [0u8; 4];
        let col_len = write_col_letter(num_cols - 1, &mut col_buf);
        buf.extend_from_slice(&col_buf[..col_len]);
        buf.extend_from_slice(int_buf.format(total_rows + 1).as_bytes());
        buf.extend_from_slice(b"\"/>");
    }

    // MergeCells
    if !config.merge_cells.is_empty() {
        // A merge ref touching row 0 (e.g. "A0:C0") is invalid -- Excel rows are
        // 1-based -- and a single bad ref makes the ENTIRE worksheet unreadable.
        // Skip any such range rather than corrupting the whole file for one bad
        // input. (merge coordinates are 1-based, matching hyperlinks.)
        let valid_merges: Vec<&crate::styles::MergeRange> = config
            .merge_cells
            .iter()
            .filter(|m| m.start_row >= 1 && m.end_row >= 1)
            .collect();

        if !valid_merges.is_empty() {
            buf.extend_from_slice(b"<mergeCells count=\"");
            buf.extend_from_slice(itoa::Buffer::new().format(valid_merges.len()).as_bytes());
            buf.extend_from_slice(b"\">");

            for merge in valid_merges {
                buf.extend_from_slice(b"<mergeCell ref=\"");
                write_normalized_range(merge.start_col, merge.start_row, merge.end_col, merge.end_row, &mut buf);
                buf.extend_from_slice(b"\"/>");
            }

            buf.extend_from_slice(b"</mergeCells>");
        }
    }

    // ConditionalFormatting
    if !config.conditional_formats.is_empty() {
        write_conditional_formatting(&mut buf, &config.conditional_formats, config);
    }

    // DataValidations
    if !config.data_validations.is_empty() {
        buf.extend_from_slice(b"<dataValidations count=\"");
        buf.extend_from_slice(itoa::Buffer::new().format(config.data_validations.len()).as_bytes());
        buf.extend_from_slice(b"\">");
        
        for validation in &config.data_validations {
            buf.extend_from_slice(b"<dataValidation sqref=\"");
            write_normalized_range(validation.start_col, validation.start_row, validation.end_col, validation.end_row, &mut buf);
            buf.extend_from_slice(b"\" ");
            
            match &validation.validation_type {
                ValidationType::List(_items) => {
                    buf.extend_from_slice(b"type=\"list\" showDropDown=\"");
                    buf.push(if validation.show_dropdown { b'0' } else { b'1' });
                    buf.extend_from_slice(b"\"");
                }
                ValidationType::WholeNumber { .. } => {
                    buf.extend_from_slice(b"type=\"whole\" operator=\"between\"");
                }
                ValidationType::Decimal { .. } => {
                    buf.extend_from_slice(b"type=\"decimal\" operator=\"between\"");
                }
                ValidationType::TextLength { .. } => {
                    buf.extend_from_slice(b"type=\"textLength\" operator=\"between\"");
                }
            }
            
            if let Some(title) = &validation.error_title {
                buf.extend_from_slice(b" errorTitle=\"");
                xml_escape_simd(title.as_bytes(), &mut buf);
                buf.push(b'\"');
            }
            if let Some(msg) = &validation.error_message {
                buf.extend_from_slice(b" error=\"");
                xml_escape_simd(msg.as_bytes(), &mut buf);
                buf.push(b'\"');
            }
            
            buf.push(b'>');
            
            match &validation.validation_type {
                ValidationType::List(items) => {
                    buf.extend_from_slice(b"<formula1>\"");
                    for (i, item) in items.iter().enumerate() {
                        if i > 0 { buf.push(b','); }
                        xml_escape_simd(item.as_bytes(), &mut buf);
                    }
                    buf.extend_from_slice(b"\"</formula1>");
                }
                ValidationType::WholeNumber { min, max } => {
                    buf.extend_from_slice(b"<formula1>");
                    buf.extend_from_slice(itoa::Buffer::new().format(*min).as_bytes());
                    buf.extend_from_slice(b"</formula1><formula2>");
                    buf.extend_from_slice(itoa::Buffer::new().format(*max).as_bytes());
                    buf.extend_from_slice(b"</formula2>");
                }
                ValidationType::Decimal { min, max } => {
                    buf.extend_from_slice(b"<formula1>");
                    buf.extend_from_slice(zmij::Buffer::new().format(*min).as_bytes());
                    buf.extend_from_slice(b"</formula1><formula2>");
                    buf.extend_from_slice(zmij::Buffer::new().format(*max).as_bytes());
                    buf.extend_from_slice(b"</formula2>");
                }
                ValidationType::TextLength { min, max } => {
                    buf.extend_from_slice(b"<formula1>");
                    buf.extend_from_slice(itoa::Buffer::new().format(*min).as_bytes());
                    buf.extend_from_slice(b"</formula1><formula2>");
                    buf.extend_from_slice(itoa::Buffer::new().format(*max).as_bytes());
                    buf.extend_from_slice(b"</formula2>");
                }
            }
            
            buf.extend_from_slice(b"</dataValidation>");
        }
        
        buf.extend_from_slice(b"</dataValidations>");
    }

    // Hyperlinks
    if !config.hyperlinks.is_empty() {
        buf.extend_from_slice(b"<hyperlinks>");
        
        for (idx, hyperlink) in config.hyperlinks.iter().enumerate() {
            buf.extend_from_slice(b"<hyperlink ref=\"");
            // Rows are 1-based; a ref like "A0" is invalid and makes the whole
            // sheet unreadable. Clamp row 0 to 1 rather than corrupt the file.
            // (r:id numbering is untouched, so it stays in sync with the rels.)
            let mut hl_row = hyperlink.row.max(1);
            let mut hl_col = hyperlink.col;
            // A hyperlink whose ref lands INSIDE a merged range but not on its
            // top-left anchor makes readers (openpyxl, Excel) choke -- the
            // non-anchor cells don't exist as addressable cells. Relocate the ref
            // to the merge's anchor so the link stays valid. With overlapping
            // merges a single move can land inside another merge, so iterate to a
            // fixed point (bounded by the merge count to avoid any cycle).
            for _ in 0..config.merge_cells.len() + 1 {
                let mut moved = false;
                for m in &config.merge_cells {
                    if m.start_row < 1 || m.end_row < 1 {
                        continue;
                    }
                    let (r0, r1) = (m.start_row.min(m.end_row), m.start_row.max(m.end_row));
                    let (c0, c1) = (m.start_col.min(m.end_col), m.start_col.max(m.end_col));
                    if hl_row >= r0 && hl_row <= r1 && hl_col >= c0 && hl_col <= c1
                        && !(hl_row == r0 && hl_col == c0)
                    {
                        hl_row = r0;
                        hl_col = c0;
                        moved = true;
                    }
                }
                if !moved {
                    break;
                }
            }
            write_cell_ref(hl_col, hl_row, &mut buf);
            buf.extend_from_slice(b"\" r:id=\"rId");
            buf.extend_from_slice(itoa::Buffer::new().format(idx + 1).as_bytes());
            buf.extend_from_slice(b"\"/>");
        }
        
        buf.extend_from_slice(b"</hyperlinks>");
    }

    // Drawing (for charts and images)
    if !config.charts.is_empty() || !config.images.is_empty() {
        buf.extend_from_slice(b"<drawing r:id=\"rIdDraw1\"/>");
    }

    // TableParts (MUST be after drawing)
    if !config.tables.is_empty() {
        buf.extend_from_slice(b"<tableParts count=\"");
        buf.extend_from_slice(itoa::Buffer::new().format(config.tables.len()).as_bytes());
        buf.extend_from_slice(b"\">");
        
        for idx in 0..config.tables.len() {
            buf.extend_from_slice(b"<tablePart r:id=\"rIdTable");
            buf.extend_from_slice(itoa::Buffer::new().format(idx + 1).as_bytes());
            buf.extend_from_slice(b"\"/>");
        }
        
        buf.extend_from_slice(b"</tableParts>");
    }

    buf.extend_from_slice(b"</worksheet>");
    
    Ok(buf)
}


/// Write conditional formatting section
fn write_conditional_formatting(buf: &mut Vec<u8>, formats: &[ConditionalFormat], config: &StyleConfig) {
    for (idx, format) in formats.iter().enumerate() {
        buf.extend_from_slice(b"<conditionalFormatting sqref=\"");
        write_normalized_range(format.start_col, format.start_row, format.end_col, format.end_row, buf);
        buf.extend_from_slice(b"\">");
        
        buf.extend_from_slice(b"<cfRule type=\"");
        
        match &format.rule {
            ConditionalRule::CellValue { operator, value } => {
                // Get DXF ID from the properly built map
                if let Some(&dxf_id) = config.cond_format_dxf_ids.get(&idx) {
                    buf.extend_from_slice(b"cellIs\" dxfId=\"");
                    buf.extend_from_slice(itoa::Buffer::new().format(dxf_id).as_bytes());
                    buf.extend_from_slice(b"\" operator=\"");
                } else {
                    buf.extend_from_slice(b"cellIs\" operator=\"");
                }
                let op_str = match operator {
                    ComparisonOperator::GreaterThan => "greaterThan",
                    ComparisonOperator::LessThan => "lessThan",
                    ComparisonOperator::Equal => "equal",
                    ComparisonOperator::NotEqual => "notEqual",
                    ComparisonOperator::GreaterThanOrEqual => "greaterThanOrEqual",
                    ComparisonOperator::LessThanOrEqual => "lessThanOrEqual",
                    ComparisonOperator::Between => "between",
                };
                buf.extend_from_slice(op_str.as_bytes());
                buf.extend_from_slice(b"\" priority=\"");
                buf.extend_from_slice(itoa::Buffer::new().format(format.priority).as_bytes());
                buf.extend_from_slice(b"\"><formula>");
                xml_escape_simd(value.as_bytes(), buf);
                buf.extend_from_slice(b"</formula></cfRule>");
            }
            ConditionalRule::ColorScale { min_color, max_color, mid_color } => {
                buf.extend_from_slice(b"colorScale\" priority=\"");
                buf.extend_from_slice(itoa::Buffer::new().format(format.priority).as_bytes());
                buf.extend_from_slice(b"\"><colorScale><cfvo type=\"min\"/>");
                if mid_color.is_some() {
                    buf.extend_from_slice(b"<cfvo type=\"percentile\" val=\"50\"/>");
                }
                buf.extend_from_slice(b"<cfvo type=\"max\"/>");
                // A colorScale requires exactly one <color> per <cfvo>, so an
                // invalid color must fall back to a valid default rather than be
                // dropped (which would unbalance the pairing and corrupt the file).
                let min_c = crate::styles::normalize_color(min_color).unwrap_or_else(|| "FFFF0000".to_string());
                buf.extend_from_slice(b"<color rgb=\"");
                buf.extend_from_slice(min_c.as_bytes());
                buf.extend_from_slice(b"\"/>");
                if let Some(mid) = mid_color {
                    let mid_c = crate::styles::normalize_color(mid).unwrap_or_else(|| "FFFFFF00".to_string());
                    buf.extend_from_slice(b"<color rgb=\"");
                    buf.extend_from_slice(mid_c.as_bytes());
                    buf.extend_from_slice(b"\"/>");
                }
                let max_c = crate::styles::normalize_color(max_color).unwrap_or_else(|| "FF00FF00".to_string());
                buf.extend_from_slice(b"<color rgb=\"");
                buf.extend_from_slice(max_c.as_bytes());
                buf.extend_from_slice(b"\"/>");
                buf.extend_from_slice(b"</colorScale></cfRule>");
            }
            ConditionalRule::DataBar { color, show_value } => {
                buf.extend_from_slice(b"dataBar\" priority=\"");
                buf.extend_from_slice(itoa::Buffer::new().format(format.priority).as_bytes());
                buf.extend_from_slice(b"\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"");
                let bar_c = crate::styles::normalize_color(color).unwrap_or_else(|| "FF638EC6".to_string());
                buf.extend_from_slice(bar_c.as_bytes());
                buf.extend_from_slice(b"\"/>");
                if !show_value {
                    buf.extend_from_slice(b"<showValue val=\"0\"/>");
                }
                buf.extend_from_slice(b"</dataBar></cfRule>");
            }
            ConditionalRule::Top10 { rank, bottom } => {
                if let Some(&dxf_id) = config.cond_format_dxf_ids.get(&idx) {
                    buf.extend_from_slice(b"top10\" dxfId=\"");
                    buf.extend_from_slice(itoa::Buffer::new().format(dxf_id).as_bytes());
                    buf.extend_from_slice(b"\" priority=\"");
                } else {
                    buf.extend_from_slice(b"top10\" priority=\"");
                }
                buf.extend_from_slice(itoa::Buffer::new().format(format.priority).as_bytes());
                buf.extend_from_slice(b"\" rank=\"");
                buf.extend_from_slice(itoa::Buffer::new().format(*rank).as_bytes());
                if *bottom {
                    buf.extend_from_slice(b"\" bottom=\"1\"/>");
                } else {
                    buf.extend_from_slice(b"\"/>");
                }
            }
        }
        
        buf.extend_from_slice(b"</conditionalFormatting>");
    }
}

/// Write a single Arrow cell with formula and hyperlink support
#[inline(always)]
fn write_arrow_cell_to_xml_optimized(
    array: &dyn Array,
    row_idx: usize,
    cell_ref: &[u8],
    style_id: Option<u32>,
    hyperlink: Option<&&Hyperlink>,
    formula: Option<&&Formula>,
    buf: &mut Vec<u8>,
    ryu_buf: &mut zmij::Buffer,
    int_buf: &mut itoa::Buffer,
    // Some(sst) iff this column was selected for shared strings. When set, Utf8
    // cells emit `t="s"` + index instead of an inline string. None keeps the
    // original inline path byte-for-byte.
    sst: Option<&SharedStrings>,
) -> Result<(), WriteError> {
    use arrow_array::*;
    
    if let Some(f) = formula {
        buf.extend_from_slice(b"<c r=\"");
        buf.extend_from_slice(cell_ref);
        if let Some(sid) = style_id {
            buf.extend_from_slice(b"\" s=\"");
            buf.extend_from_slice(int_buf.format(sid).as_bytes());
        }
        buf.extend_from_slice(b"\"><f>");
        // In OOXML the <f> element content must NOT include the leading '='
        // (it's implicit). If the caller passes "=A1+B1" (natural in Excel),
        // writing it verbatim yields "==A1+B1" when reopened. Strip one leading
        // '=' so both "=A1+B1" and "A1+B1" produce a correct formula.
        let formula_body = f.formula.strip_prefix('=').unwrap_or(&f.formula);
        xml_escape_simd(formula_body.as_bytes(), buf);
        buf.extend_from_slice(b"</f>");
        
        if let Some(ref cached) = f.cached_value {
            buf.extend_from_slice(b"<v>");
            xml_escape_simd(cached.as_bytes(), buf);
            buf.extend_from_slice(b"</v>");
        }
        
        buf.extend_from_slice(b"</c>");
        return Ok(());
    }
    
    if let Some(hl) = hyperlink {
        let display_text = hl.display.as_ref().map(|s| s.as_str()).unwrap_or(&hl.url);
        
        buf.extend_from_slice(b"<c r=\"");
        buf.extend_from_slice(cell_ref);
        buf.extend_from_slice(b"\" s=\"9\" t=\"inlineStr\">");
        write_inline_string(display_text.as_bytes(), buf);
        buf.extend_from_slice(b"</c>");
        return Ok(());
    }
    
    if array.is_null(row_idx) {
        buf.extend_from_slice(b"<c r=\"");
        buf.extend_from_slice(cell_ref);
        if let Some(sid) = style_id {
            buf.extend_from_slice(b"\" s=\"");
            buf.extend_from_slice(int_buf.format(sid).as_bytes());
        }
        buf.extend_from_slice(b"\"/>");
        return Ok(());
    }

    match array.data_type() {
        DataType::Utf8 => {
            let arr = array.as_any().downcast_ref::<StringArray>().unwrap();
            
            let offsets = arr.offsets();
            let values = arr.values();
            let start = offsets[row_idx] as usize;
            let end = offsets[row_idx + 1] as usize;
            let str_bytes = &values.as_ref()[start..end];
            
            // Skip empty strings entirely to allow text overflow
            if str_bytes.is_empty() && style_id.is_none() && hyperlink.is_none() && formula.is_none() {
                return Ok(());
            }

            // Shared-string fast path: emit `t="s"` + table index. Only reached
            // for columns the pre-pass selected; the string is guaranteed to be
            // in the table (built from these same arrays), so lookup is a hit.
            // SAFETY: Arrow StringArray values are guaranteed valid UTF-8, so the
            // unchecked conversion avoids a redundant per-cell validation scan.
            if let Some(sst) = sst {
                let s = unsafe { std::str::from_utf8_unchecked(str_bytes) };
                if let Some(idx) = sst.get(s) {
                    buf.extend_from_slice(b"<c r=\"");
                    buf.extend_from_slice(cell_ref);
                    if let Some(sid) = style_id {
                        buf.extend_from_slice(b"\" s=\"");
                        buf.extend_from_slice(int_buf.format(sid).as_bytes());
                    }
                    buf.extend_from_slice(b"\" t=\"s\"><v>");
                    buf.extend_from_slice(int_buf.format(idx).as_bytes());
                    buf.extend_from_slice(b"</v></c>");
                    return Ok(());
                }
            }

            buf.extend_from_slice(b"<c r=\"");
            buf.extend_from_slice(cell_ref);
            if let Some(sid) = style_id {
                buf.extend_from_slice(b"\" s=\"");
                buf.extend_from_slice(int_buf.format(sid).as_bytes());
            }
            buf.extend_from_slice(b"\" t=\"inlineStr\">");
            write_inline_string(str_bytes, buf);
            buf.extend_from_slice(b"</c>");
        }
        DataType::LargeUtf8 => {
            let arr = array.as_any().downcast_ref::<LargeStringArray>().unwrap();
            
            let offsets = arr.offsets();
            let values = arr.values();
            let start = offsets[row_idx] as usize;
            let end = offsets[row_idx + 1] as usize;
            let str_bytes = &values.as_ref()[start..end];

            // Skip empty strings entirely to allow text overflow
            if str_bytes.is_empty() && style_id.is_none() && hyperlink.is_none() && formula.is_none() {
                return Ok(());
            }

            if let Some(sst) = sst {
                // SAFETY: Arrow LargeStringArray values are valid UTF-8.
                let s = unsafe { std::str::from_utf8_unchecked(str_bytes) };
                if let Some(idx) = sst.get(s) {
                    buf.extend_from_slice(b"<c r=\"");
                    buf.extend_from_slice(cell_ref);
                    if let Some(sid) = style_id {
                        buf.extend_from_slice(b"\" s=\"");
                        buf.extend_from_slice(int_buf.format(sid).as_bytes());
                    }
                    buf.extend_from_slice(b"\" t=\"s\"><v>");
                    buf.extend_from_slice(int_buf.format(idx).as_bytes());
                    buf.extend_from_slice(b"</v></c>");
                    return Ok(());
                }
            }
            
            buf.extend_from_slice(b"<c r=\"");
            buf.extend_from_slice(cell_ref);
            if let Some(sid) = style_id {
                buf.extend_from_slice(b"\" s=\"");
                buf.extend_from_slice(int_buf.format(sid).as_bytes());
            }
            buf.extend_from_slice(b"\" t=\"inlineStr\">");
            write_inline_string(str_bytes, buf);
            buf.extend_from_slice(b"</c>");
        }
        DataType::Int8 => {
            let arr = array.as_any().downcast_ref::<Int8Array>().unwrap();
            write_number_cell_int(arr.value(row_idx) as i64, cell_ref, style_id, buf, int_buf);
        }
        DataType::Int16 => {
            let arr = array.as_any().downcast_ref::<Int16Array>().unwrap();
            write_number_cell_int(arr.value(row_idx) as i64, cell_ref, style_id, buf, int_buf);
        }
        DataType::Int32 => {
            let arr = array.as_any().downcast_ref::<Int32Array>().unwrap();
            write_number_cell_int(arr.value(row_idx) as i64, cell_ref, style_id, buf, int_buf);
        }
        DataType::Int64 => {
            let arr = array.as_any().downcast_ref::<Int64Array>().unwrap();
            write_number_cell_int(arr.value(row_idx), cell_ref, style_id, buf, int_buf);
        }
        DataType::UInt8 => {
            let arr = array.as_any().downcast_ref::<UInt8Array>().unwrap();
            write_number_cell_int(arr.value(row_idx) as i64, cell_ref, style_id, buf, int_buf);
        }
        DataType::UInt16 => {
            let arr = array.as_any().downcast_ref::<UInt16Array>().unwrap();
            write_number_cell_int(arr.value(row_idx) as i64, cell_ref, style_id, buf, int_buf);
        }
        DataType::UInt32 => {
            let arr = array.as_any().downcast_ref::<UInt32Array>().unwrap();
            write_number_cell_int(arr.value(row_idx) as i64, cell_ref, style_id, buf, int_buf);
        }
        DataType::UInt64 => {
            let arr = array.as_any().downcast_ref::<UInt64Array>().unwrap();
            let v = arr.value(row_idx);
            // See the note on the U64 arm of the optimized writer: values above
            // i64::MAX must go through the f64 path or they wrap to negative.
            if v <= i64::MAX as u64 {
                write_number_cell_int(v as i64, cell_ref, style_id, buf, int_buf);
            } else {
                write_number_cell(v as f64, cell_ref, style_id, buf, ryu_buf, int_buf);
            }
        }
        DataType::Float32 => {
            let arr = array.as_any().downcast_ref::<Float32Array>().unwrap();
            write_number_cell(arr.value(row_idx) as f64, cell_ref, style_id, buf, ryu_buf, int_buf);
        }
        DataType::Float64 => {
            let arr = array.as_any().downcast_ref::<Float64Array>().unwrap();
            write_number_cell(arr.value(row_idx), cell_ref, style_id, buf, ryu_buf, int_buf);
        }
        DataType::Boolean => {
            let arr = array.as_any().downcast_ref::<BooleanArray>().unwrap();
            buf.extend_from_slice(b"<c r=\"");
            buf.extend_from_slice(cell_ref);
            if let Some(sid) = style_id {
                buf.extend_from_slice(b"\" s=\"");
                buf.extend_from_slice(int_buf.format(sid).as_bytes());
            }
            buf.extend_from_slice(b"\" t=\"b\"><v>");
            buf.push(if arr.value(row_idx) { b'1' } else { b'0' });
            buf.extend_from_slice(b"</v></c>");
        }
        DataType::Date32 => {
            let arr = array.as_any().downcast_ref::<Date32Array>().unwrap();
            let days = arr.value(row_idx);
            let date = chrono::NaiveDate::from_ymd_opt(1970, 1, 1)
                .ok_or_else(|| WriteError::Validation("Invalid base date".to_string()))?
                .checked_add_signed(chrono::Duration::days(days as i64))
                .ok_or_else(|| WriteError::Validation("Date out of range".to_string()))?;
            let dt = date.and_hms_opt(0, 0, 0).unwrap();
            write_date_cell(&dt, cell_ref, style_id.or(Some(10)), buf, ryu_buf);
        }
        DataType::Date64 => {
            let arr = array.as_any().downcast_ref::<Date64Array>().unwrap();
            let millis = arr.value(row_idx);
            let datetime = chrono::DateTime::from_timestamp_millis(millis)
                .ok_or_else(|| WriteError::Validation("Invalid timestamp".to_string()))?;
            write_date_cell(&datetime.naive_utc(), cell_ref, style_id.or(Some(10)), buf, ryu_buf); // Date-only format
        }
       DataType::Time32(unit) => {
            use arrow_schema::TimeUnit;
            let seconds = match unit {
                TimeUnit::Second => {
                    let arr = array.as_any().downcast_ref::<Time32SecondArray>().unwrap();
                    arr.value(row_idx) as f64
                }
                TimeUnit::Millisecond => {
                    let arr = array.as_any().downcast_ref::<Time32MillisecondArray>().unwrap();
                    arr.value(row_idx) as f64 / 1000.0
                }
                _ => 0.0,
            };
            let time_fraction = seconds / 86400.0;
            write_number_cell(time_fraction, cell_ref, style_id, buf, ryu_buf, int_buf);
        }
        DataType::Time64(unit) => {
            use arrow_schema::TimeUnit;
            let seconds = match unit {
                TimeUnit::Microsecond => {
                    let arr = array.as_any().downcast_ref::<Time64MicrosecondArray>().unwrap();
                    arr.value(row_idx) as f64 / 1_000_000.0
                }
                TimeUnit::Nanosecond => {
                    let arr = array.as_any().downcast_ref::<Time64NanosecondArray>().unwrap();
                    arr.value(row_idx) as f64 / 1_000_000_000.0
                }
                _ => 0.0,
            };
            let time_fraction = seconds / 86400.0;
            write_number_cell(time_fraction, cell_ref, style_id, buf, ryu_buf, int_buf);
        }
        DataType::Timestamp(unit, _) => {
            use arrow_schema::TimeUnit;
            let dt = match unit {
                TimeUnit::Second => {
                    let arr = array.as_any().downcast_ref::<TimestampSecondArray>().unwrap();
                    let value = arr.value(row_idx);
                    chrono::DateTime::from_timestamp(value, 0)
                        .ok_or_else(|| WriteError::Validation("Invalid timestamp".to_string()))?
                        .naive_utc()
                }
                TimeUnit::Millisecond => {
                    let arr = array.as_any().downcast_ref::<TimestampMillisecondArray>().unwrap();
                    let value = arr.value(row_idx);
                    chrono::DateTime::from_timestamp_millis(value)
                        .ok_or_else(|| WriteError::Validation("Invalid timestamp".to_string()))?
                        .naive_utc()
                }
                TimeUnit::Microsecond => {
                    let arr = array.as_any().downcast_ref::<TimestampMicrosecondArray>().unwrap();
                    let value = arr.value(row_idx);
                    chrono::DateTime::from_timestamp_micros(value)
                        .ok_or_else(|| WriteError::Validation("Invalid timestamp".to_string()))?
                        .naive_utc()
                }
                TimeUnit::Nanosecond => {
                    let arr = array.as_any().downcast_ref::<TimestampNanosecondArray>().unwrap();
                    let value = arr.value(row_idx);
                    let secs = value / 1_000_000_000;
                    let nsecs = (value % 1_000_000_000) as u32;
                    chrono::DateTime::from_timestamp(secs, nsecs)
                        .ok_or_else(|| WriteError::Validation("Invalid timestamp".to_string()))?
                        .naive_utc()
                }
            };
            write_date_cell(&dt, cell_ref, style_id.or(Some(1)), buf, ryu_buf);
        }
        _ => {
            buf.extend_from_slice(b"<c r=\"");
            buf.extend_from_slice(cell_ref);
            if let Some(sid) = style_id {
                buf.extend_from_slice(b"\" s=\"");
                buf.extend_from_slice(int_buf.format(sid).as_bytes());
            }
            buf.extend_from_slice(b"\"/>");
        }
    }
    
    Ok(())
}

#[inline(always)]
fn write_number_cell_int(
    n: i64,
    cell_ref: &[u8],
    style_id: Option<u32>,
    buf: &mut Vec<u8>,
    int_buf: &mut itoa::Buffer,
) {
    buf.extend_from_slice(b"<c r=\"");
    buf.extend_from_slice(cell_ref);
    if let Some(sid) = style_id {
        buf.extend_from_slice(b"\" s=\"");
        buf.extend_from_slice(int_buf.format(sid).as_bytes());
    }
    buf.extend_from_slice(b"\"><v>");
    buf.extend_from_slice(int_buf.format(n).as_bytes());
    buf.extend_from_slice(b"</v></c>");
}

#[inline(always)]
fn write_number_cell(
    n: f64,
    cell_ref: &[u8],
    style_id: Option<u32>,
    buf: &mut Vec<u8>,
    ryu_buf: &mut zmij::Buffer,
    int_buf: &mut itoa::Buffer,
) {

    // Excel can't handle NaN or inf - write empty cell instead
    if !n.is_finite() {
        buf.extend_from_slice(b"<c r=\"");
        buf.extend_from_slice(cell_ref);
        if let Some(sid) = style_id {
            buf.extend_from_slice(b"\" s=\"");
            buf.extend_from_slice(int_buf.format(sid).as_bytes());
        }
        buf.extend_from_slice(b"\"/>");
        return;
    }

    buf.extend_from_slice(b"<c r=\"");
    buf.extend_from_slice(cell_ref);
    if let Some(sid) = style_id {
        buf.extend_from_slice(b"\" s=\"");
        buf.extend_from_slice(int_buf.format(sid).as_bytes());
    }
    buf.extend_from_slice(b"\"><v>");
    
    let abs = n.abs();
    if n.fract() == 0.0 && abs < 9007199254740992.0 && abs > 0.0 {
        buf.extend_from_slice(int_buf.format(n as i64).as_bytes());
    } else {
        buf.extend_from_slice(ryu_buf.format(n).as_bytes());
    }
    
    buf.extend_from_slice(b"</v></c>");
}

#[inline(always)]
fn write_date_cell(
    dt: &chrono::NaiveDateTime,
    cell_ref: &[u8],
    style_id: Option<u32>,
    buf: &mut Vec<u8>,
    ryu_buf: &mut zmij::Buffer,
) {
    let serial = datetime_to_excel_serial(dt);
    // Excel's date system starts at 1900-01-01 (serial 1). A date before that
    // yields a serial < 1 (often negative), which Excel can't display as a date
    // -- the cell shows ###### / is flagged invalid. Preserve the value by
    // writing it as an ISO-8601 inline string instead of a broken date serial.
    if serial < 1.0 {
        buf.extend_from_slice(b"<c r=\"");
        buf.extend_from_slice(cell_ref);
        buf.extend_from_slice(b"\" t=\"inlineStr\"><is><t>");
        let iso = dt.format("%Y-%m-%dT%H:%M:%S").to_string();
        // strip the trailing T00:00:00 for pure dates to read as a plain date
        let iso = iso.strip_suffix("T00:00:00").unwrap_or(&iso);
        xml_escape_simd(iso.as_bytes(), buf);
        buf.extend_from_slice(b"</t></is></c>");
        return;
    }
    buf.extend_from_slice(b"<c r=\"");
    buf.extend_from_slice(cell_ref);
    buf.extend_from_slice(b"\" s=\"");
    buf.extend_from_slice(itoa::Buffer::new().format(style_id.unwrap_or(1)).as_bytes());
    buf.extend_from_slice(b"\"><v>");
    buf.extend_from_slice(ryu_buf.format(serial).as_bytes());
    buf.extend_from_slice(b"</v></c>");
}

/// Dict API - Original path (kept for backward compatibility)
pub fn generate_sheet_xml_from_dict(
    sheet: &SheetData,
    config: &StyleConfig,
) -> Result<Vec<u8>, WriteError> {
    let num_rows = sheet.num_rows();
    let num_cols = sheet.num_cols();

    if num_cols == 0 {
        return Ok(b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
<dimension ref=\"A1\"/><sheetData/></worksheet>".to_vec());
    }

    let avg_cell_size = estimate_avg_cell_size(sheet);
    let estimated_size = 1000 + (num_rows + 1) * num_cols * avg_cell_size;
    let mut buf = Vec::with_capacity(estimated_size);

    buf.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

    buf.extend_from_slice(b"<dimension ref=\"");
    if num_rows > 0 {
        buf.extend_from_slice(b"A1:");
        let mut col_buf = [0u8; 4];
        let col_len = write_col_letter(num_cols - 1, &mut col_buf);
        buf.extend_from_slice(&col_buf[..col_len]);
        
        let mut row_buf = itoa::Buffer::new();
        buf.extend_from_slice(row_buf.format(num_rows + 1).as_bytes());
    } else {
        buf.extend_from_slice(b"A1");
    }
    buf.extend_from_slice(b"\"/>");

    if config.freeze_rows > 0 || config.freeze_cols > 0 {
        buf.extend_from_slice(b"<sheetViews><sheetView workbookViewId=\"0\">");
        buf.extend_from_slice(b"<pane ");
        
        if config.freeze_cols > 0 {
            buf.extend_from_slice(b"xSplit=\"");
            buf.extend_from_slice(itoa::Buffer::new().format(config.freeze_cols).as_bytes());
            buf.extend_from_slice(b"\" ");
        }
        
        if config.freeze_rows > 0 {
            buf.extend_from_slice(b"ySplit=\"");
            buf.extend_from_slice(itoa::Buffer::new().format(config.freeze_rows).as_bytes());
            buf.extend_from_slice(b"\" ");
        }
        
        buf.extend_from_slice(b"topLeftCell=\"");
        write_cell_ref(config.freeze_cols, config.freeze_rows + 1, &mut buf);
        buf.extend_from_slice(b"\" activePane=\"bottomRight\" state=\"frozen\"/>");
        buf.extend_from_slice(b"</sheetView></sheetViews>");
    }

    buf.extend_from_slice(b"<sheetData>");

    let col_letters: Vec<([u8; 4], usize)> = (0..num_cols)
        .map(|i| {
            let mut col_buf = [0u8; 4];
            let len = write_col_letter(i, &mut col_buf);
            (col_buf, len)
        })
        .collect();

    let mut ryu_buf = zmij::Buffer::new();
    let mut int_buf = itoa::Buffer::new();
    let mut cell_int_buf = itoa::Buffer::new();
    let mut cell_ref = [0u8; 16];

    buf.extend_from_slice(b"<row r=\"1\">");
    for (col_idx, (header, _)) in sheet.columns.iter().enumerate() {
        let (col_letter, col_len) = &col_letters[col_idx];
        
        buf.extend_from_slice(b"<c r=\"");
        buf.extend_from_slice(&col_letter[..*col_len]);
        buf.extend_from_slice(b"1\" t=\"inlineStr\">");
        write_inline_string(header.as_bytes(), &mut buf);
        buf.extend_from_slice(b"</c>");
    }
    buf.extend_from_slice(b"</row>");

    for row_idx in 0..num_rows {
        let row_num = row_idx + 2;
        let row_str = int_buf.format(row_num);
        let row_bytes = row_str.as_bytes();

        buf.extend_from_slice(b"<row r=\"");
        buf.extend_from_slice(row_bytes);
        buf.extend_from_slice(b"\">");

        for col_idx in 0..num_cols {
            let cell_val = &sheet.columns[col_idx].1[row_idx];
            let (col_letter, col_len) = &col_letters[col_idx];

            let cell_ref_len = {
                cell_ref[..*col_len].copy_from_slice(&col_letter[..*col_len]);
                cell_ref[*col_len..*col_len + row_bytes.len()].copy_from_slice(row_bytes);
                *col_len + row_bytes.len()
            };
            let cell_ref_slice = &cell_ref[..cell_ref_len];

            match cell_val {
                CellValue::Empty => {
                    buf.extend_from_slice(b"<c r=\"");
                    buf.extend_from_slice(cell_ref_slice);
                    buf.extend_from_slice(b"\"/>");
                }
                CellValue::String(s) => {
                    buf.extend_from_slice(b"<c r=\"");
                    buf.extend_from_slice(cell_ref_slice);
                    buf.extend_from_slice(b"\" t=\"inlineStr\">");
                    write_inline_string(s.as_bytes(), &mut buf);
                    buf.extend_from_slice(b"</c>");
                }
                CellValue::Number(n) => {
                    buf.extend_from_slice(b"<c r=\"");
                    buf.extend_from_slice(cell_ref_slice);
                    buf.extend_from_slice(b"\"><v>");
                    
                    let abs = n.abs();
                    if n.fract() == 0.0 && abs < 9007199254740992.0 && abs > 0.0 {
                        buf.extend_from_slice(cell_int_buf.format(*n as i64).as_bytes());
                    } else {
                        buf.extend_from_slice(ryu_buf.format(*n).as_bytes());
                    }
                    buf.extend_from_slice(b"</v></c>");
                }
                CellValue::Bool(b) => {
                    buf.extend_from_slice(b"<c r=\"");
                    buf.extend_from_slice(cell_ref_slice);
                    buf.extend_from_slice(b"\" t=\"b\"><v>");
                    buf.push(if *b { b'1' } else { b'0' });
                    buf.extend_from_slice(b"</v></c>");
                }
                CellValue::Date(dt) => {
                    let serial = datetime_to_excel_serial(dt);
                    if serial < 1.0 {
                        // Pre-1900 date: Excel can't display it as a date serial,
                        // so preserve the value as an ISO string (see write_date_cell).
                        buf.extend_from_slice(b"<c r=\"");
                        buf.extend_from_slice(cell_ref_slice);
                        buf.extend_from_slice(b"\" t=\"inlineStr\"><is><t>");
                        let iso = dt.format("%Y-%m-%dT%H:%M:%S").to_string();
                        let iso = iso.strip_suffix("T00:00:00").unwrap_or(&iso);
                        xml_escape_simd(iso.as_bytes(), &mut buf);
                        buf.extend_from_slice(b"</t></is></c>");
                    } else {
                        buf.extend_from_slice(b"<c r=\"");
                        buf.extend_from_slice(cell_ref_slice);
                        buf.extend_from_slice(b"\" s=\"1\"><v>");
                        buf.extend_from_slice(ryu_buf.format(serial).as_bytes());
                        buf.extend_from_slice(b"</v></c>");
                    }
                }
            }
        }
        buf.extend_from_slice(b"</row>");
    }

    buf.extend_from_slice(b"</sheetData>");
    
    if config.auto_filter && num_rows > 0 {
        buf.extend_from_slice(b"<autoFilter ref=\"A1:");
        let mut col_buf = [0u8; 4];
        let col_len = write_col_letter(num_cols - 1, &mut col_buf);
        buf.extend_from_slice(&col_buf[..col_len]);
        buf.extend_from_slice(itoa::Buffer::new().format(num_rows + 1).as_bytes());
        buf.extend_from_slice(b"\"/>");
    }
    

    if !config.charts.is_empty() {
    buf.extend_from_slice(b"<drawing r:id=\"rIdDraw1\"/>");
    }
    
    buf.extend_from_slice(b"</worksheet>");
    Ok(buf)
}

#[inline]
fn estimate_avg_cell_size(sheet: &SheetData) -> usize {
    if sheet.columns.is_empty() {
        return 30;
    }
    
    let sample_size = sheet.num_rows().min(100);
    if sample_size == 0 {
        return 30;
    }
    
    let mut total = 0;
    for (_, col_data) in &sheet.columns {
        for cell in col_data.iter().take(sample_size) {
            total += match cell {
                CellValue::Empty => 15,
                CellValue::String(s) => 40 + s.len(),
                CellValue::Number(_) => 25,
                CellValue::Bool(_) => 20,
                CellValue::Date(_) => 30,
            };
        }
    }
    
    (total / (sample_size * sheet.num_cols())).max(25)
}


/// Generate drawing XML with both charts and images
pub fn generate_drawing_xml_combined(charts: &[ExcelChart], images: &[ExcelImage]) -> String {
    let total_elements = charts.len() + images.len();
    let mut xml = String::with_capacity(2000 + total_elements * 1000);
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    xml.push_str("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" ");
    xml.push_str("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n");
    
    let mut element_id = 1;
    
    // Add charts
    for (idx, chart) in charts.iter().enumerate() {
        let chart_id = idx + 1;
        xml.push_str("<xdr:twoCellAnchor>\n");
        
        xml.push_str("<xdr:from>\n");
        xml.push_str(&format!("<xdr:col>{}</xdr:col>\n", chart.position.from_col));
        xml.push_str("<xdr:colOff>0</xdr:colOff>\n");
        xml.push_str(&format!("<xdr:row>{}</xdr:row>\n", chart.position.from_row));
        xml.push_str("<xdr:rowOff>0</xdr:rowOff>\n");
        xml.push_str("</xdr:from>\n");
        
        xml.push_str("<xdr:to>\n");
        xml.push_str(&format!("<xdr:col>{}</xdr:col>\n", chart.position.to_col));
        xml.push_str("<xdr:colOff>0</xdr:colOff>\n");
        xml.push_str(&format!("<xdr:row>{}</xdr:row>\n", chart.position.to_row));
        xml.push_str("<xdr:rowOff>0</xdr:rowOff>\n");
        xml.push_str("</xdr:to>\n");
        
        xml.push_str("<xdr:graphicFrame macro=\"\">\n");
        xml.push_str("<xdr:nvGraphicFramePr>\n");
        xml.push_str(&format!("<xdr:cNvPr id=\"{}\" name=\"Chart {}\"/>\n", element_id, chart_id));
        element_id += 1;
        xml.push_str("<xdr:cNvGraphicFramePr/>\n");
        xml.push_str("</xdr:nvGraphicFramePr>\n");
        xml.push_str("<xdr:xfrm>\n");
        xml.push_str("<a:off x=\"0\" y=\"0\"/>\n");
        xml.push_str("<a:ext cx=\"0\" cy=\"0\"/>\n");
        xml.push_str("</xdr:xfrm>\n");
        xml.push_str("<a:graphic>\n");
        xml.push_str("<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\n");
        xml.push_str(&format!("<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdChart{}\"/>\n", chart_id));
        xml.push_str("</a:graphicData>\n");
        xml.push_str("</a:graphic>\n");
        xml.push_str("</xdr:graphicFrame>\n");
        xml.push_str("<xdr:clientData/>\n");
        xml.push_str("</xdr:twoCellAnchor>\n");
    }
    
    // Add images
    for (idx, image) in images.iter().enumerate() {
        let image_id = idx + 1;
        xml.push_str("<xdr:twoCellAnchor>\n");
        
        xml.push_str("<xdr:from>\n");
        xml.push_str(&format!("<xdr:col>{}</xdr:col>\n", image.position.from_col));
        xml.push_str("<xdr:colOff>0</xdr:colOff>\n");
        xml.push_str(&format!("<xdr:row>{}</xdr:row>\n", image.position.from_row));
        xml.push_str("<xdr:rowOff>0</xdr:rowOff>\n");
        xml.push_str("</xdr:from>\n");
        
        xml.push_str("<xdr:to>\n");
        xml.push_str(&format!("<xdr:col>{}</xdr:col>\n", image.position.to_col));
        xml.push_str("<xdr:colOff>0</xdr:colOff>\n");
        xml.push_str(&format!("<xdr:row>{}</xdr:row>\n", image.position.to_row));
        xml.push_str("<xdr:rowOff>0</xdr:rowOff>\n");
        xml.push_str("</xdr:to>\n");
        
        xml.push_str("<xdr:pic>\n");
        xml.push_str("<xdr:nvPicPr>\n");
        xml.push_str(&format!("<xdr:cNvPr id=\"{}\" name=\"Image {}\"/>\n", element_id, image_id));
        element_id += 1;
        xml.push_str("<xdr:cNvPicPr>\n");
        xml.push_str("<a:picLocks noChangeAspect=\"1\"/>\n");
        xml.push_str("</xdr:cNvPicPr>\n");
        xml.push_str("</xdr:nvPicPr>\n");
        
        xml.push_str("<xdr:blipFill>\n");
        xml.push_str(&format!("<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rIdImage{}\"/>\n", image_id));
        xml.push_str("<a:stretch>\n");
        xml.push_str("<a:fillRect/>\n");
        xml.push_str("</a:stretch>\n");
        xml.push_str("</xdr:blipFill>\n");
        
        xml.push_str("<xdr:spPr>\n");
        xml.push_str("<a:xfrm>\n");
        xml.push_str("<a:off x=\"0\" y=\"0\"/>\n");
        xml.push_str("<a:ext cx=\"0\" cy=\"0\"/>\n");
        xml.push_str("</a:xfrm>\n");
        xml.push_str("<a:prstGeom prst=\"rect\">\n");
        xml.push_str("<a:avLst/>\n");
        xml.push_str("</a:prstGeom>\n");
        xml.push_str("</xdr:spPr>\n");
        
        xml.push_str("</xdr:pic>\n");
        xml.push_str("<xdr:clientData/>\n");
        xml.push_str("</xdr:twoCellAnchor>\n");
    }
    
    xml.push_str("</xdr:wsDr>");
    xml
}

/// Generate drawing relationships for both charts and images
pub fn generate_drawing_rels_combined(num_charts: usize, images: &[ExcelImage], start_chart_id: usize, start_image_id: usize) -> String {
    let mut xml = String::with_capacity(300 + (num_charts + images.len()) * 150);
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    xml.push_str("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
    
    for i in 0..num_charts {
        let local_id = i + 1;
        let global_chart_id = start_chart_id + i;
        xml.push_str(&format!("<Relationship Id=\"rIdChart{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart{}.xml\"/>\n", local_id, global_chart_id));
    }
    
    for (idx, image) in images.iter().enumerate() {
        // rIdImage{local} is the drawing-local relationship id, but the media
        // file name must be workbook-GLOBAL: multiple sheets each numbering their
        // images from 1 collided on xl/media/image1.png (a duplicate ZIP entry),
        // so a second sheet's image silently overwrote/aliased the first.
        let local_id = idx + 1;
        let global_image_id = start_image_id + idx;
        xml.push_str(&format!("<Relationship Id=\"rIdImage{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image{}.{}\"/>\n", local_id, global_image_id, image.extension));
    }
    
    xml.push_str("</Relationships>");
    xml
}