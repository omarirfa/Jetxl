use crate::types::{CellValue, SheetData, WriteError};
use crate::styles::*;
use arrow_array::{Array, RecordBatch,Time32SecondArray, Time32MillisecondArray, Time64MicrosecondArray, Time64NanosecondArray};
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

pub fn generate_content_types_with_charts(
    sheet_names: &[&str], 
    tables_per_sheet: &[usize], 
    charts_per_sheet: &[usize],
    images_per_sheet: &[(&[ExcelImage], usize)]
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

/// Generate worksheet relationships with table support
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
        xml.push_str(url);
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
    xml.push_str(&table.name);
    xml.push_str("\" displayName=\"");
    xml.push_str(&table.display_name);
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
        xml.push_str(style);
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
        
        let font_size = chart.title_font_size.unwrap_or(1400);
        xml.push_str(&format!("<a:defRPr sz=\"{}\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" spc=\"0\" baseline=\"0\">\n", font_size));
        
        if let Some(ref color) = chart.title_color {
            xml.push_str(&format!("<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>\n", color));
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
        xml.push_str(&format!("<a:t>{}</a:t>\n", title));
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
            xml.push_str(&format!("<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>\n", color));
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
        
        let legend_size = chart.legend_font_size.unwrap_or(900);
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
    xml.push_str("<c:dispBlanksAs val=\"gap\"/>\n");
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
    
    let font_size = chart.axis_title_font_size.unwrap_or(1000);
    xml.push_str(&format!("<a:defRPr sz=\"{}\"", font_size));
    if chart.axis_title_bold {
        xml.push_str(" b=\"1\"");
    } else {
        xml.push_str(" b=\"0\"");
    }
    xml.push_str(" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n");
    
    if let Some(ref color) = chart.axis_title_color {
        xml.push_str(&format!("<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>\n", color));
    } else {
        xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill>\n");
    }
    
    xml.push_str("<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\n");
    xml.push_str("</a:defRPr>\n");
    xml.push_str("</a:pPr>\n");
    xml.push_str("<a:r>\n");
    xml.push_str("<a:rPr lang=\"en-US\"/>\n");
    xml.push_str(&format!("<a:t>{}</a:t>\n", title));
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
        xml.push_str(&format!("<c:v>{}</c:v>\n", series_name));
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
        
        let font_size = chart.axis_title_font_size.unwrap_or(1000);
        xml.push_str(&format!("<a:defRPr sz=\"{}\"", font_size));
        if chart.axis_title_bold {
            xml.push_str(" b=\"1\"");
        } else {
            xml.push_str(" b=\"0\"");
        }
        xml.push_str(" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">\n");
        
        if let Some(ref color) = chart.axis_title_color {
            xml.push_str(&format!("<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>\n", color));
        } else {
            xml.push_str("<a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill>\n");
        }
        
        xml.push_str("<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\n");
        xml.push_str("</a:defRPr>\n");
        xml.push_str("</a:pPr>\n");
        xml.push_str("<a:r>\n");
        xml.push_str("<a:rPr lang=\"en-US\"/>\n");
        xml.push_str(&format!("<a:t>{}</a:t>\n", y_title));
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
        xml.push_str(&format!("<c:v>{}</c:v>\n", series_name));
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
        xml.push_str(&format!("<c:v>{}</c:v>\n", series_name));
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
    if let Some(ref y_title) = chart.y_axis_title {
        write_axis_title(xml, y_title, chart);
    }
    xml.push_str("<c:majorGridlines/>\n");
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
        xml.push_str(&format!("<c:v>{}</c:v>\n", series_name));
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


/// Generate complete sheet XML with all enhanced features
/// Element order: dimension → sheetViews → sheetFormatPr → cols → sheetData → 
///                autoFilter → mergeCells → conditionalFormatting → dataValidations → 
///                hyperlinks → drawing → tableParts
pub fn generate_sheet_xml_from_arrow(
    batches: &[RecordBatch],
    config: &StyleConfig,
    col_format_map: &HashMap<usize, u32>,
    cell_style_map: &HashMap<(usize, usize), u32>,
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

    // Build map of table header rows that need to be inserted
    let mut table_header_rows: HashMap<usize, (usize, usize)> = HashMap::new();
    let mut num_inserted_headers = 0;
    for table in &config.tables {
        let (start_row, start_col, _, end_col) = table.range;
        if start_row > 1 {
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
        buf.extend_from_slice(color.as_bytes());
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
    buf.extend_from_slice(ryu::Buffer::new().format(default_height).as_bytes());
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
            buf.extend_from_slice(ryu::Buffer::new().format(width).as_bytes());
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

    let mut ryu_buf = ryu::Buffer::new();
    let mut int_buf = itoa::Buffer::new();
    let mut cell_int_buf = itoa::Buffer::new();
    let mut cell_ref = [0u8; 16];

    let hyperlink_map: HashMap<(usize, usize), &Hyperlink> = config.hyperlinks
        .iter()
        .map(|h| ((h.row, h.col), h))
        .collect();
    
    let formula_map: HashMap<(usize, usize), &Formula> = config.formulas
        .iter()
        .map(|f| ((f.row, f.col), f))
        .collect();

    // Determine where DataFrame data actually starts
    let data_start = if config.write_header_row { 
        config.data_start_row.max(1) 
    } else { 
        config.data_start_row 
    };

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
                    buf.extend_from_slice(ryu::Buffer::new().format(*height).as_bytes());
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
                    buf.extend_from_slice(b" t=\"inlineStr\"><is><t>");
                    xml_escape_simd(text.as_bytes(), &mut buf);
                    buf.extend_from_slice(b"</t></is></c>");
                }
            }
            
            buf.extend_from_slice(b"</row>");
        }
    }

    // Write DataFrame header row at data_start (only if enabled)
    if config.write_header_row {
        let header_row_height = config.row_heights.as_ref().and_then(|h| h.get(&data_start));
        buf.extend_from_slice(b"<row r=\"");
        buf.extend_from_slice(itoa::Buffer::new().format(data_start).as_bytes());
        buf.push(b'\"');
        if let Some(height) = header_row_height {
            buf.extend_from_slice(b" ht=\"");
            buf.extend_from_slice(ryu::Buffer::new().format(*height).as_bytes());
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
            buf.extend_from_slice(b"\" t=\"inlineStr\"><is><t>");
            xml_escape_simd(field.name().as_bytes(), &mut buf);
            buf.extend_from_slice(b"</t></is></c>");
        }
        buf.extend_from_slice(b"</row>");
    }

    let mut current_row = if config.write_header_row { data_start + 1 } else { data_start };
    
    // Build map of table header rows that need to be inserted
    let mut table_header_rows: HashMap<usize, (usize, usize)> = HashMap::new();
    let mut num_inserted_headers = 0;
    for table in &config.tables {
        let (start_row, start_col, _, end_col) = table.range;
        // Only insert header if table starts after data_start and doesn't already have a header
        if start_row > data_start {
            table_header_rows.insert(start_row, (start_col, end_col));
            num_inserted_headers += 1;
        }
    }
    
    
    // Cache feature flags to avoid repeated checks
    let has_table_headers = !table_header_rows.is_empty();
    let has_row_heights = config.row_heights.is_some();
    let has_hidden_rows = !config.hidden_rows.is_empty();
    
    // Write data rows (with optional table header insertion)
    for batch in batches {
        let batch_rows = batch.num_rows();
        
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
                            buf.extend_from_slice(ryu::Buffer::new().format(*height).as_bytes());
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
                        buf.extend_from_slice(b"\" t=\"inlineStr\"><is><t>");
                        xml_escape_simd(field_name.as_bytes(), &mut buf);
                        buf.extend_from_slice(b"</t></is></c>");
                    }
                    
                    buf.extend_from_slice(b"</row>");
                    current_row += 1;
                }
            }
            
            // Write actual data row
            let row_num = current_row;
            let row_str = int_buf.format(row_num);
            let row_bytes = row_str.as_bytes();

            buf.extend_from_slice(b"<row r=\"");
            buf.extend_from_slice(row_bytes);
            buf.push(b'\"');
            
            if has_row_heights {
                if let Some(height) = config.row_heights.as_ref().unwrap().get(&row_num) {
                    buf.extend_from_slice(b" ht=\"");
                    buf.extend_from_slice(ryu::Buffer::new().format(*height).as_bytes());
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

    // Hyperlinks
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
        write_cell_ref(format.start_col, format.start_row, buf);
        buf.push(b':');
        write_cell_ref(format.end_col, format.end_row, buf);
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
    ryu_buf: &mut ryu::Buffer,
    int_buf: &mut itoa::Buffer,
) -> Result<(), WriteError> {
    use arrow_array::*;
    
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
        buf.extend_from_slice(b"\" s=\"9\" t=\"inlineStr\"><is><t>");
        xml_escape_simd(display_text.as_bytes(), buf);
        buf.extend_from_slice(b"</t></is></c>");
        return Ok(());
    }
    
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

            // Skip empty strings entirely to allow text overflow
            if str_bytes.is_empty() && style_id.is_none() && hyperlink.is_none() && formula.is_none() {
                return Ok(());
            }
            
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

    // Excel can't handle NaN or inf - write empty cell instead
    if !n.is_finite() {
        buf.extend_from_slice(b"<c r=\"");
        buf.extend_from_slice(cell_ref);
        if let Some(sid) = style_id {
            buf.extend_from_slice(b"\" s=\"");
            buf.extend_from_slice(itoa::Buffer::new().format(sid).as_bytes());
        }
        buf.extend_from_slice(b"\"/>");
        return;
    }

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
pub fn generate_drawing_rels_combined(num_charts: usize, images: &[ExcelImage]) -> String {
    let mut xml = String::with_capacity(300 + (num_charts + images.len()) * 150);
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    xml.push_str("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
    
    for i in 1..=num_charts {
        xml.push_str(&format!("<Relationship Id=\"rIdChart{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart{}.xml\"/>\n", i, i));
    }
    
    for (idx, image) in images.iter().enumerate() {
        let i = idx + 1;
        xml.push_str(&format!("<Relationship Id=\"rIdImage{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image{}.{}\"/>\n", i, i, image.extension));
    }
    
    xml.push_str("</Relationships>");
    xml
}