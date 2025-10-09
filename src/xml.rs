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
        sheet_names.iter().map(|n| format!("<vt:lpstr>{}</vt:lpstr>", n)).collect::<Vec<_>>().join("")
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

#[inline(always)]
fn datetime_to_excel_serial(dt: &chrono::NaiveDateTime) -> f64 {
    let excel_epoch = chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
    let days = (dt.date() - excel_epoch).num_days() as f64;
    let time_fraction = (dt.hour() * 3600 + dt.minute() * 60 + dt.second()) as f64 / 86400.0;
    days + time_fraction
}

/// SIMD-accelerated XML escaping
#[inline(always)]
pub fn xml_escape_simd(input: &[u8], output: &mut Vec<u8>) {
    let needs_escape = memchr::memchr3(b'&', b'<', b'>', input).is_some()
        || memchr::memchr2(b'"', b'\'', input).is_some();
    
    if !needs_escape {
        output.extend_from_slice(input);
        return;
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

pub fn generate_content_types(sheet_names: &[&str]) -> String {
    let mut xml = String::with_capacity(800 + sheet_names.len() * 150);
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
        xml.push_str(name);
        xml.push_str("\" sheetId=\"");
        xml.push_str(&id.to_string());
        xml.push_str("\" r:id=\"rId");
        xml.push_str(&id.to_string());
        xml.push_str("\"/>");
    }

    xml.push_str("</sheets><calcPr calcId=\"191029\"/></workbook>");
    xml
}

pub fn generate_workbook_rels(num_sheets: usize) -> String {
    let mut xml = String::with_capacity(300 + num_sheets * 150);
    xml.push_str(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
<Relationship Id=\"rId100\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>",
    );

    for i in 1..=num_sheets {
        xml.push_str("<Relationship Id=\"rId");
        xml.push_str(&i.to_string());
        xml.push_str("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet");
        xml.push_str(&i.to_string());
        xml.push_str(".xml\"/>");
    }

    xml.push_str("</Relationships>");
    xml
}

/// Generate worksheet relationships (for hyperlinks)
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
        xml.push_str(url);
        xml.push_str("\" TargetMode=\"External\"/>");
    }

    xml.push_str("</Relationships>");
    Some(xml)
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
    
    let num_rows = arr.len();
    if num_rows == 0 {
        return 0;
    }
    
    let mut total = 0;
    for i in 0..num_rows {
        if !arr.is_null(i) {
            total += arr.value(i).len();
        }
    }
    total
}

fn get_large_string_array_total_bytes(arr: &arrow_array::LargeStringArray) -> usize {
    use arrow_array::Array;
    
    let num_rows = arr.len();
    if num_rows == 0 {
        return 0;
    }
    
    let mut total = 0;
    for i in 0..num_rows {
        if !arr.is_null(i) {
            total += arr.value(i).len();
        }
    }
    total
}

/// Generate complete sheet XML with all enhanced features
/// Element order: dimension → sheetViews → sheetFormatPr → cols → sheetData → 
///                autoFilter → mergeCells → dataValidations → hyperlinks → conditionalFormatting
pub fn generate_sheet_xml_from_arrow(
    batches: &[RecordBatch],
    config: &StyleConfig,
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

    // Build style registry for custom cell styles
    let mut style_registry = StyleRegistry::new();
    let mut cell_style_map: HashMap<(usize, usize), u32> = HashMap::new();
    
    for cell_style in &config.cell_styles {
        let style_id = style_registry.register_cell_style(&cell_style.style);
        cell_style_map.insert((cell_style.row, cell_style.col), style_id);
    }

    let exact_size = calculate_exact_xml_size(batches)?;
    let mut buf = Vec::with_capacity(exact_size);

    buf.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\
<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

    // 1. DIMENSION (must come before sheetViews)
    buf.extend_from_slice(b"<dimension ref=\"");
    if total_rows > 0 {
        buf.extend_from_slice(b"A1:");
        let mut col_buf = [0u8; 4];
        let col_len = write_col_letter(num_cols - 1, &mut col_buf);
        buf.extend_from_slice(&col_buf[..col_len]);
        
        let mut row_buf = itoa::Buffer::new();
        buf.extend_from_slice(row_buf.format(total_rows + 1).as_bytes());
    } else {
        buf.extend_from_slice(b"A1");
    }
    buf.extend_from_slice(b"\"/>");

    // 2. SHEETVIEWS (always include)
    buf.extend_from_slice(b"<sheetViews><sheetView workbookViewId=\"0\"");
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

    // 3. SHEETFORMATPR (default row height if specified)
    if config.row_heights.is_some() {
        // Default row height: 15 points (standard)
        buf.extend_from_slice(b"<sheetFormatPr defaultRowHeight=\"15\"/>");
    }

    // 4. COLS (column widths)
    if config.auto_width || config.column_widths.is_some() {
        buf.extend_from_slice(b"<cols>");
        
        for (col_idx, field) in schema.fields().iter().enumerate() {
            let width = if let Some(widths) = &config.column_widths {
                widths.get(field.name()).copied().unwrap_or_else(|| {
                    if config.auto_width && !batches.is_empty() {
                        calculate_column_width(batches[0].column(col_idx).as_ref(), 
                                             field.name(), 100)
                    } else { 8.43 }
                })
            } else if config.auto_width && !batches.is_empty() {
                calculate_column_width(batches[0].column(col_idx).as_ref(), 
                                     field.name(), 100)
            } else {
                8.43
            };
            
            buf.extend_from_slice(b"<col min=\"");
            buf.extend_from_slice(itoa::Buffer::new().format(col_idx + 1).as_bytes());
            buf.extend_from_slice(b"\" max=\"");
            buf.extend_from_slice(itoa::Buffer::new().format(col_idx + 1).as_bytes());
            buf.extend_from_slice(b"\" width=\"");
            buf.extend_from_slice(ryu::Buffer::new().format(width).as_bytes());
            buf.extend_from_slice(b"\" customWidth=\"1\"/>");
        }
        
        buf.extend_from_slice(b"</cols>");
    }

    // 5. SHEETDATA (the actual data)
    buf.extend_from_slice(b"<sheetData>");

    let col_letters: Vec<([u8; 4], usize)> = (0..num_cols)
        .map(|i| {
            let mut col_buf = [0u8; 4];
            let len = write_col_letter(i, &mut col_buf);
            (col_buf, len)
        })
        .collect();

    let mut ryu_buf = ryu::Buffer::new();
    let mut int_buf = itoa::Buffer::new();
    let mut cell_int_buf = itoa::Buffer::new();
    let mut cell_ref = [0u8; 16];

    // Map column formats to style IDs
    let col_format_map: HashMap<usize, u32> = if let Some(formats) = &config.column_formats {
        schema.fields().iter().enumerate()
            .filter_map(|(idx, field)| {
                formats.get(field.name()).map(|fmt| {
                    let style_id = match fmt {
                        NumberFormat::General => 0,
                        NumberFormat::Date | NumberFormat::DateTime | NumberFormat::Time => 1,
                        NumberFormat::Currency | NumberFormat::CurrencyRounded => 4,
                        NumberFormat::Percentage => 5,
                        NumberFormat::PercentageDecimal => 6,
                        NumberFormat::Integer => 7,
                        NumberFormat::Decimal2 | NumberFormat::Decimal4 => 8,
                    };
                    (idx, style_id)
                })
            })
            .collect()
    } else {
        HashMap::new()
    };

    // Build hyperlink and formula lookup maps
    let hyperlink_map: HashMap<(usize, usize), &Hyperlink> = config.hyperlinks
        .iter()
        .map(|h| ((h.row, h.col), h))
        .collect();
    
    let formula_map: HashMap<(usize, usize), &Formula> = config.formulas
        .iter()
        .map(|f| ((f.row, f.col), f))
        .collect();

    // Header row
    let header_row_height = config.row_heights.as_ref().and_then(|h| h.get(&1));
    buf.extend_from_slice(b"<row r=\"1\"");
    if let Some(height) = header_row_height {
        buf.extend_from_slice(b" ht=\"");
        buf.extend_from_slice(ryu::Buffer::new().format(*height).as_bytes());
        buf.extend_from_slice(b"\" customHeight=\"1\"");
    }
    buf.extend_from_slice(b">");
    
    for (col_idx, field) in schema.fields().iter().enumerate() {
        let (col_letter, col_len) = &col_letters[col_idx];
        
        let style_id = if config.styled_headers { 2 } else { 0 };
        
        buf.extend_from_slice(b"<c r=\"");
        buf.extend_from_slice(&col_letter[..*col_len]);
        buf.extend_from_slice(b"1\"");
        if style_id > 0 {
            buf.extend_from_slice(b" s=\"");
            buf.extend_from_slice(int_buf.format(style_id).as_bytes());
            buf.extend_from_slice(b"\"");
        }
        buf.extend_from_slice(b" t=\"inlineStr\"><is><t>");
        xml_escape_simd(field.name().as_bytes(), &mut buf);
        buf.extend_from_slice(b"</t></is></c>");
    }
    buf.extend_from_slice(b"</row>");

    // Data rows
    let mut current_row = 2;
    
    for batch in batches {
        let batch_rows = batch.num_rows();
        
        for row_idx in 0..batch_rows {
            let row_num = current_row;
            let row_str = int_buf.format(row_num);
            let row_bytes = row_str.as_bytes();

            buf.extend_from_slice(b"<row r=\"");
            buf.extend_from_slice(row_bytes);
            buf.extend_from_slice(b"\"");
            
            // Check for custom row height
            if let Some(heights) = &config.row_heights {
                if let Some(height) = heights.get(&row_num) {
                    buf.extend_from_slice(b" ht=\"");
                    buf.extend_from_slice(ryu::Buffer::new().format(*height).as_bytes());
                    buf.extend_from_slice(b"\" customHeight=\"1\"");
                }
            }
            
            buf.extend_from_slice(b">");

            for col_idx in 0..num_cols {
                let array = batch.column(col_idx);
                let (col_letter, col_len) = &col_letters[col_idx];

                let cell_ref_len = {
                    cell_ref[..*col_len].copy_from_slice(&col_letter[..*col_len]);
                    cell_ref[*col_len..*col_len + row_bytes.len()].copy_from_slice(row_bytes);
                    *col_len + row_bytes.len()
                };
                let cell_ref_slice = &cell_ref[..cell_ref_len];

                // Check for custom cell style, formula, or hyperlink
                let custom_style_id = cell_style_map.get(&(row_num, col_idx)).copied();
                let default_style_id = col_format_map.get(&col_idx).copied();
                let style_id = custom_style_id.or(default_style_id);
                
                let hyperlink = hyperlink_map.get(&(row_num, col_idx));
                let formula = formula_map.get(&(row_num, col_idx));

                write_arrow_cell_to_xml_optimized(
                    array.as_ref(),
                    row_idx,
                    cell_ref_slice,
                    style_id,
                    hyperlink,
                    formula,
                    &mut buf,
                    &mut ryu_buf,
                    &mut cell_int_buf,
                )?;
            }
            
            buf.extend_from_slice(b"</row>");
            current_row += 1;
        }
    }

    buf.extend_from_slice(b"</sheetData>");

    // 6. AUTOFILTER
    if config.auto_filter && total_rows > 0 {
        buf.extend_from_slice(b"<autoFilter ref=\"A1:");
        let mut col_buf = [0u8; 4];
        let col_len = write_col_letter(num_cols - 1, &mut col_buf);
        buf.extend_from_slice(&col_buf[..col_len]);
        buf.extend_from_slice(int_buf.format(total_rows + 1).as_bytes());
        buf.extend_from_slice(b"\"/>");
    }

    // 7. MERGED CELLS
    if !config.merge_cells.is_empty() {
        buf.extend_from_slice(b"<mergeCells count=\"");
        buf.extend_from_slice(itoa::Buffer::new().format(config.merge_cells.len()).as_bytes());
        buf.extend_from_slice(b"\">");
        
        for merge in &config.merge_cells {
            buf.extend_from_slice(b"<mergeCell ref=\"");
            write_cell_ref(merge.start_col, merge.start_row, &mut buf);
            buf.push(b':');
            write_cell_ref(merge.end_col, merge.end_row, &mut buf);
            buf.extend_from_slice(b"\"/>");
        }
        
        buf.extend_from_slice(b"</mergeCells>");
    }
     // 8. CONDITIONAL FORMATTING
    if !config.conditional_formats.is_empty() {
        write_conditional_formatting(&mut buf, &config.conditional_formats, config);
    }

    // 9. DATA VALIDATIONS
    if !config.data_validations.is_empty() {
        buf.extend_from_slice(b"<dataValidations count=\"");
        buf.extend_from_slice(itoa::Buffer::new().format(config.data_validations.len()).as_bytes());
        buf.extend_from_slice(b"\">");
        
        for validation in &config.data_validations {
            buf.extend_from_slice(b"<dataValidation sqref=\"");
            write_cell_ref(validation.start_col, validation.start_row, &mut buf);
            buf.push(b':');
            write_cell_ref(validation.end_col, validation.end_row, &mut buf);
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
                    buf.extend_from_slice(ryu::Buffer::new().format(*min).as_bytes());
                    buf.extend_from_slice(b"</formula1><formula2>");
                    buf.extend_from_slice(ryu::Buffer::new().format(*max).as_bytes());
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

    // 10. HYPERLINKS
    if !config.hyperlinks.is_empty() {
        buf.extend_from_slice(b"<hyperlinks>");
        
        for (idx, hyperlink) in config.hyperlinks.iter().enumerate() {
            buf.extend_from_slice(b"<hyperlink ref=\"");
            write_cell_ref(hyperlink.col, hyperlink.row, &mut buf);
            buf.extend_from_slice(b"\" r:id=\"rId");
            buf.extend_from_slice(itoa::Buffer::new().format(idx + 1).as_bytes());
            buf.extend_from_slice(b"\"/>");
        }
        
        buf.extend_from_slice(b"</hyperlinks>");
    }

    buf.extend_from_slice(b"</worksheet>");
    
    Ok(buf)
}


/// Write conditional formatting section
fn write_conditional_formatting(buf: &mut Vec<u8>, formats: &[ConditionalFormat], config: &StyleConfig) {
    for (idx, format) in formats.iter().enumerate() {
        buf.extend_from_slice(b"<conditionalFormatting sqref=\"");
        write_cell_ref(format.start_col, format.start_row, buf);
        buf.push(b':');
        write_cell_ref(format.end_col, format.end_row, buf);
        buf.extend_from_slice(b"\">");
        
        buf.extend_from_slice(b"<cfRule type=\"");
        
        match &format.rule {
            ConditionalRule::CellValue { operator, value } => {
                let dxf_id = config.cond_format_dxf_ids.get(&idx).copied().unwrap_or(0);
                buf.extend_from_slice(b"cellIs\" dxfId=\"");
                buf.extend_from_slice(itoa::Buffer::new().format(dxf_id).as_bytes());
                buf.extend_from_slice(b"\" operator=\"");
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
                buf.extend_from_slice(b"<color rgb=\"");
                buf.extend_from_slice(min_color.as_bytes());
                buf.extend_from_slice(b"\"/>");
                if let Some(mid) = mid_color {
                    buf.extend_from_slice(b"<color rgb=\"");
                    buf.extend_from_slice(mid.as_bytes());
                    buf.extend_from_slice(b"\"/>");
                }
                buf.extend_from_slice(b"<color rgb=\"");
                buf.extend_from_slice(max_color.as_bytes());
                buf.extend_from_slice(b"\"/></colorScale></cfRule>");
            }
            ConditionalRule::DataBar { color, show_value } => {
                buf.extend_from_slice(b"dataBar\" priority=\"");
                buf.extend_from_slice(itoa::Buffer::new().format(format.priority).as_bytes());
                buf.extend_from_slice(b"\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"");
                buf.extend_from_slice(color.as_bytes());
                buf.extend_from_slice(b"\"/>");
                if !show_value {
                    buf.extend_from_slice(b"<showValue val=\"0\"/>");
                }
                buf.extend_from_slice(b"</dataBar></cfRule>");
            }
            ConditionalRule::Top10 { rank, bottom } => {
                let dxf_id = config.cond_format_dxf_ids.get(&idx).copied().unwrap_or(0);
                buf.extend_from_slice(b"top10\" dxfId=\"");
                buf.extend_from_slice(itoa::Buffer::new().format(dxf_id).as_bytes());
                buf.extend_from_slice(b"\" priority=\"");
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
    ryu_buf: &mut ryu::Buffer,
    int_buf: &mut itoa::Buffer,
) -> Result<(), WriteError> {
    use arrow_array::*;
    
    // Handle formulas - formula takes precedence
    if let Some(f) = formula {
        buf.extend_from_slice(b"<c r=\"");
        buf.extend_from_slice(cell_ref);
        if let Some(sid) = style_id {
            buf.extend_from_slice(b"\" s=\"");
            buf.extend_from_slice(itoa::Buffer::new().format(sid).as_bytes());
        }
        buf.extend_from_slice(b"\"><f>");
        xml_escape_simd(f.formula.as_bytes(), buf);
        buf.extend_from_slice(b"</f>");
        
        // Add cached value if provided
        if let Some(ref cached) = f.cached_value {
            buf.extend_from_slice(b"<v>");
            xml_escape_simd(cached.as_bytes(), buf);
            buf.extend_from_slice(b"</v>");
        }
        
        buf.extend_from_slice(b"</c>");
        return Ok(());
    }
    
    // Handle hyperlinks - display text instead of cell value
    if let Some(hl) = hyperlink {
        let display_text = hl.display.as_ref().map(|s| s.as_str()).unwrap_or(&hl.url);
        
        buf.extend_from_slice(b"<c r=\"");
        buf.extend_from_slice(cell_ref);
        buf.extend_from_slice(b"\" s=\"9\" t=\"inlineStr\"><is><t>");
        xml_escape_simd(display_text.as_bytes(), buf);
        buf.extend_from_slice(b"</t></is></c>");
        return Ok(());
    }
    
    // Handle null cells
    if array.is_null(row_idx) {
        buf.extend_from_slice(b"<c r=\"");
        buf.extend_from_slice(cell_ref);
        if let Some(sid) = style_id {
            buf.extend_from_slice(b"\" s=\"");
            buf.extend_from_slice(itoa::Buffer::new().format(sid).as_bytes());
        }
        buf.extend_from_slice(b"\"/>");
        return Ok(());
    }

    // Handle regular cell values by type
    match array.data_type() {
        DataType::Utf8 => {
            let arr = array.as_any().downcast_ref::<StringArray>().unwrap();
            
            let offsets = arr.offsets();
            let values = arr.values();
            let start = offsets[row_idx] as usize;
            let end = offsets[row_idx + 1] as usize;
            let str_bytes = &values.as_ref()[start..end];
            
            buf.extend_from_slice(b"<c r=\"");
            buf.extend_from_slice(cell_ref);
            if let Some(sid) = style_id {
                buf.extend_from_slice(b"\" s=\"");
                buf.extend_from_slice(itoa::Buffer::new().format(sid).as_bytes());
            }
            buf.extend_from_slice(b"\" t=\"inlineStr\"><is><t>");
            xml_escape_simd(str_bytes, buf);
            buf.extend_from_slice(b"</t></is></c>");
        }
        DataType::LargeUtf8 => {
            let arr = array.as_any().downcast_ref::<LargeStringArray>().unwrap();
            
            let offsets = arr.offsets();
            let values = arr.values();
            let start = offsets[row_idx] as usize;
            let end = offsets[row_idx + 1] as usize;
            let str_bytes = &values.as_ref()[start..end];
            
            buf.extend_from_slice(b"<c r=\"");
            buf.extend_from_slice(cell_ref);
            if let Some(sid) = style_id {
                buf.extend_from_slice(b"\" s=\"");
                buf.extend_from_slice(itoa::Buffer::new().format(sid).as_bytes());
            }
            buf.extend_from_slice(b"\" t=\"inlineStr\"><is><t>");
            xml_escape_simd(str_bytes, buf);
            buf.extend_from_slice(b"</t></is></c>");
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
            write_number_cell_int(arr.value(row_idx) as i64, cell_ref, style_id, buf, int_buf);
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
                buf.extend_from_slice(itoa::Buffer::new().format(sid).as_bytes());
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
            write_date_cell(&dt, cell_ref, style_id, buf, ryu_buf);
        }
        DataType::Date64 => {
            let arr = array.as_any().downcast_ref::<Date64Array>().unwrap();
            let millis = arr.value(row_idx);
            let datetime = chrono::DateTime::from_timestamp_millis(millis)
                .ok_or_else(|| WriteError::Validation("Invalid timestamp".to_string()))?;
            write_date_cell(&datetime.naive_utc(), cell_ref, style_id, buf, ryu_buf);
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
            write_date_cell(&dt, cell_ref, style_id, buf, ryu_buf);
        }
        _ => {
            buf.extend_from_slice(b"<c r=\"");
            buf.extend_from_slice(cell_ref);
            if let Some(sid) = style_id {
                buf.extend_from_slice(b"\" s=\"");
                buf.extend_from_slice(itoa::Buffer::new().format(sid).as_bytes());
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
        buf.extend_from_slice(itoa::Buffer::new().format(sid).as_bytes());
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
    ryu_buf: &mut ryu::Buffer,
    int_buf: &mut itoa::Buffer,
) {
    buf.extend_from_slice(b"<c r=\"");
    buf.extend_from_slice(cell_ref);
    if let Some(sid) = style_id {
        buf.extend_from_slice(b"\" s=\"");
        buf.extend_from_slice(itoa::Buffer::new().format(sid).as_bytes());
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
    ryu_buf: &mut ryu::Buffer,
) {
    buf.extend_from_slice(b"<c r=\"");
    buf.extend_from_slice(cell_ref);
    buf.extend_from_slice(b"\" s=\"");
    buf.extend_from_slice(itoa::Buffer::new().format(style_id.unwrap_or(1)).as_bytes());
    buf.extend_from_slice(b"\"><v>");
    buf.extend_from_slice(ryu_buf.format(datetime_to_excel_serial(dt)).as_bytes());
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

    let mut ryu_buf = ryu::Buffer::new();
    let mut int_buf = itoa::Buffer::new();
    let mut cell_int_buf = itoa::Buffer::new();
    let mut cell_ref = [0u8; 16];

    buf.extend_from_slice(b"<row r=\"1\">");
    for (col_idx, (header, _)) in sheet.columns.iter().enumerate() {
        let (col_letter, col_len) = &col_letters[col_idx];
        
        buf.extend_from_slice(b"<c r=\"");
        buf.extend_from_slice(&col_letter[..*col_len]);
        buf.extend_from_slice(b"1\" t=\"inlineStr\"><is><t>");
        xml_escape_simd(header.as_bytes(), &mut buf);
        buf.extend_from_slice(b"</t></is></c>");
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
                    buf.extend_from_slice(b"\" t=\"inlineStr\"><is><t>");
                    xml_escape_simd(s.as_bytes(), &mut buf);
                    buf.extend_from_slice(b"</t></is></c>");
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
                    buf.extend_from_slice(b"<c r=\"");
                    buf.extend_from_slice(cell_ref_slice);
                    buf.extend_from_slice(b"\" s=\"1\"><v>");
                    buf.extend_from_slice(ryu_buf.format(datetime_to_excel_serial(dt)).as_bytes());
                    buf.extend_from_slice(b"</v></c>");
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