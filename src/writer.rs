use crate::types::{SheetData, WriteError};
use crate::styles::{StyleConfig, generate_styles_xml, generate_styles_xml_enhanced, StyleRegistry, ConditionalRule, CellStyle, ExcelImage};
// use crate::xml::{self, generate_drawing_xml_combined, generate_drawing_rels_combined};
use crate::xml::{self, generate_drawing_xml_combined, generate_drawing_rels_combined};
// ZIP writer: jetxl's own in-memory writer (see fastzip.rs) replaces mtzip.
// It compresses each already-in-memory part directly with miniz_oxide (no
// Read-trait staging copy) and emits standard DEFLATE ZIP structure, so output
// remains a spec-valid .xlsx. The `CompressionLevel`/`ZipArchive` names and
// builder API match what mtzip exposed, so call sites are unchanged.
use crate::fastzip::{CompressionLevel, ZipArchive};
use std::fs::File;
use std::io::Write;
use std::collections::HashMap;
use arrow_array::RecordBatch;
use rayon::prelude::*;
use std::collections::HashSet;

// ============================================================================
// Thread-pool cache
// ============================================================================
//
// Each multi-sheet write is given a `num_threads` cap. Building a fresh
// `rayon::ThreadPool` for every call spawns that many OS threads and tears them
// down again when the pool drops -- a fixed cost paid on every write regardless
// of how much data is involved. For small-to-medium sheets that spawn/join cost
// is a large fraction of total time, which is exactly why parallel efficiency
// used to climb with row count (bigger sheets amortized the fixed overhead).
//
// We instead build at most one pool per distinct thread count and reuse it for
// the process lifetime. Pools are cheap to keep idle (threads park), and the
// set of thread counts a program uses is tiny (typically just one). This turns
// a per-call thread-spawn storm into a one-time cost.
/// Lightweight, env-gated phase profiler. Active only when `JETXL_PROFILE` is
/// set, so it costs a single `Instant::now()` per phase boundary otherwise and
/// prints nothing. Its purpose is to reveal, on real multi-core hardware, which
/// phase of a multi-sheet write fails to speed up as threads increase -- the
/// information a single-CPU environment cannot provide.
struct PhaseProfiler {
    enabled: bool,
    last: std::time::Instant,
    marks: Vec<(&'static str, f64)>,
}

impl PhaseProfiler {
    fn new() -> Self {
        PhaseProfiler {
            enabled: std::env::var("JETXL_PROFILE").is_ok(),
            last: std::time::Instant::now(),
            marks: Vec::new(),
        }
    }

    #[inline]
    fn mark(&mut self, name: &'static str) {
        if !self.enabled {
            return;
        }
        let now = std::time::Instant::now();
        let ms = now.duration_since(self.last).as_secs_f64() * 1000.0;
        self.marks.push((name, ms));
        self.last = now;
    }

    fn report(&self, num_threads: usize) {
        if !self.enabled {
            return;
        }
        let total: f64 = self.marks.iter().map(|(_, ms)| ms).sum();
        eprintln!("[jetxl profile] threads={} total={:.1}ms", num_threads, total);
        for (name, ms) in &self.marks {
            eprintln!("    {:<18} {:>8.1}ms  ({:>4.1}%)", name, ms, ms / total * 100.0);
        }
    }
}


fn cached_pool(num_threads: usize) -> Result<std::sync::Arc<rayon::ThreadPool>, WriteError> {
    use std::sync::{Arc, Mutex, OnceLock};
    static POOLS: OnceLock<Mutex<HashMap<usize, Arc<rayon::ThreadPool>>>> = OnceLock::new();
    let pools = POOLS.get_or_init(|| Mutex::new(HashMap::new()));

    let mut guard = pools
        .lock()
        .map_err(|_| WriteError::Validation("thread-pool cache poisoned".to_string()))?;
    if let Some(pool) = guard.get(&num_threads) {
        return Ok(Arc::clone(pool));
    }
    let pool = rayon::ThreadPoolBuilder::new()
        .num_threads(num_threads)
        .build()
        .map_err(|e| WriteError::Validation(format!("Thread pool error: {}", e)))?;
    let pool = Arc::new(pool);
    guard.insert(num_threads, Arc::clone(&pool));
    Ok(pool)
}

// ============================================================================
// DICT API - Dict-based (backward compatibility)
// ============================================================================

// Retained dict-based single-sheet writer, superseded by the arrow write path
// but kept for the dict API surface and potential direct callers.
#[allow(dead_code)]
pub fn write_single_sheet(
    sheet: &SheetData,
    filename: &str,
) -> Result<(), WriteError> {
    sheet.validate().map_err(WriteError::Validation)?;

    let mut zipper = ZipArchive::new();
    let sheet_names = vec![sheet.name.as_str()];
    
    add_static_files(&mut zipper, &sheet_names, None, &[0], &[0], &[]);
    
    let config = StyleConfig::default();
    let xml_data = xml::generate_sheet_xml_from_dict(sheet, &config)?;
    zipper
        .add_file_from_memory(xml_data, "xl/worksheets/sheet1.xml".to_string())
        .compression_level(CompressionLevel::fast())
        .done();

    
    write_zip_to_file(zipper, filename)
}

pub fn write_single_sheet_with_config(
    sheet: &SheetData,
    filename: &str,
    config: &StyleConfig,
) -> Result<(), WriteError> {
    sheet.validate().map_err(WriteError::Validation)?;

    let mut zipper = ZipArchive::new();
    let sheet_names = vec![sheet.name.as_str()];
    let charts_count = vec![config.charts.len()];
    let drawing_count = if config.charts.is_empty() && config.images.is_empty() { 0 } else { 1 };
    
    add_static_files(&mut zipper, &sheet_names, None, &[0], &charts_count, &[(vec![], drawing_count)]);
    
    let xml_data = xml::generate_sheet_xml_from_dict(sheet, config)?;
    zipper
        .add_file_from_memory(xml_data, "xl/worksheets/sheet1.xml".to_string())
        .compression_level(CompressionLevel::fast())
        .done();

    // Add chart files if any
    if !config.charts.is_empty() {
        let drawing_xml = xml::generate_drawing_xml(&config.charts);
        zipper
            .add_file_from_memory(drawing_xml.into_bytes(), "xl/drawings/drawing1.xml".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
        
        let drawing_rels = generate_drawing_rels_combined(config.charts.len(), &config.images, 1, 1);
        zipper
            .add_file_from_memory(drawing_rels.into_bytes(), "xl/drawings/_rels/drawing1.xml.rels".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
        
        for (idx, chart) in config.charts.iter().enumerate() {
            let chart_xml = xml::generate_chart_xml(chart, &sheet.name);
            zipper
                .add_file_from_memory(
                    chart_xml.into_bytes(),
                    format!("xl/charts/chart{}.xml", idx + 1)
                )
                .compression_level(CompressionLevel::fast())
                .done();
        }
        
        // Add worksheet rels for drawing
        let mut rels_xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
        rels_xml.push_str("<Relationship Id=\"rIdDraw1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing1.xml\"/>\n");
        rels_xml.push_str("</Relationships>");
        
        zipper
            .add_file_from_memory(rels_xml.into_bytes(), "xl/worksheets/_rels/sheet1.xml.rels".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
    }
    
    write_zip_to_file(zipper, filename)
}

pub fn write_multiple_sheets(
    sheets: &[SheetData],
    filename: &str,
    num_threads: usize,
) -> Result<(), WriteError> {
    for sheet in sheets {
        sheet.validate().map_err(WriteError::Validation)?;
    }
    let names: Vec<&str> = sheets.iter().map(|s| s.name.as_str()).collect();
    validate_unique_sheet_names(&names)?;

    let config = StyleConfig::default();
    
    // Generate XMLs in parallel if num_threads > 1 and multiple sheets
    let xml_sheets: Vec<Vec<u8>> = if num_threads > 1 && sheets.len() > 1 {
        let pool = cached_pool(num_threads)?;
        
        pool.install(|| {
            sheets
                .par_iter()
                .map(|sheet| xml::generate_sheet_xml_from_dict(sheet, &config))
                .collect::<Result<Vec<_>, _>>()
        })?
    } else {
        // Sequential fallback
        sheets
            .iter()
            .map(|sheet| xml::generate_sheet_xml_from_dict(sheet, &config))
            .collect::<Result<Vec<_>, _>>()?
    };

    // Build ZIP sequentially (not thread-safe)
    let mut zipper = ZipArchive::new();
    let sheet_names: Vec<&str> = sheets.iter().map(|s| s.name.as_str()).collect();

    add_static_files(&mut zipper, &sheet_names, None, &vec![0; sheets.len()], &vec![0; sheets.len()], &vec![(vec![], 0); sheets.len()]);

    for (idx, xml_data) in xml_sheets.into_iter().enumerate() {
        zipper
            .add_file_from_memory(xml_data, format!("xl/worksheets/sheet{}.xml", idx + 1))
            .compression_level(CompressionLevel::fast())
            .done();
    }

    write_zip_to_file(zipper, filename)
}

// ============================================================================
// ARROW API - Direct Arrow → XML (Zero-Copy)
// ============================================================================
#[allow(dead_code)]
pub fn write_single_sheet_arrow(
    batches: &[RecordBatch],
    sheet_name: &str,
    filename: &str,
) -> Result<(), WriteError> {
    write_single_sheet_arrow_with_config(batches, sheet_name, filename, &StyleConfig::default())
}

pub fn write_single_sheet_arrow_with_config(
    batches: &[RecordBatch],
    sheet_name: &str,
    filename: &str,
    config: &StyleConfig,
) -> Result<(), WriteError> {
    validate_sheet_name(sheet_name)?;

    // Guard against an empty batch slice BEFORE indexing batches[0]. Without this
    // an empty input panics; combined with `panic = "abort"` that terminates the
    // host Python process with no catchable exception. Return a clean error.
    if batches.is_empty() {
        return Err(WriteError::Validation(
            "Cannot write sheet from an empty set of record batches".to_string(),
        ));
    }

    // Reject sheets exceeding Excel's row/column grid (O(1), off the hot path).
    {
        let rows: usize = batches.iter().map(|b| b.num_rows()).sum::<usize>()
            + if config.write_header_row { 1 } else { 0 };
        validate_sheet_dimensions(rows, batches[0].schema().fields().len())?;
    }

    let mut registry = StyleRegistry::new();
    let mut updated_config = config.clone();

    let schema = batches[0].schema();
    let col_format_map: HashMap<usize, u32> = if let Some(formats) = &config.column_formats {
        let mut map = HashMap::new();
        for (idx, field) in schema.fields().iter().enumerate() {
            if let Some(fmt) = formats.get(field.name()) {
                let cell_style = CellStyle {
                    font: None,
                    fill: None,
                    border: None,
                    alignment: None,
                    number_format: Some(fmt.clone()),
                };
                let style_id = registry.register_cell_style(&cell_style)
                    .map_err(|e| WriteError::Validation(e))?;
                map.insert(idx, style_id);
            }
        }
        map
    } else {
        HashMap::new()
    };

    // Build cell style map - register and map user's custom cell styles
    let mut cell_style_map: HashMap<(usize, usize), u32> = HashMap::new();
    for cell_style in &config.cell_styles {
        let style_id = registry.register_cell_style(&cell_style.style)
            .map_err(|e| WriteError::Validation(e))?;
        cell_style_map.insert((cell_style.row, cell_style.col), style_id);
    }

    if !config.conditional_formats.is_empty() {
        let mut dxf_ids = HashMap::new();
        for (idx, cond_format) in config.conditional_formats.iter().enumerate() {
            match &cond_format.rule {
                ConditionalRule::CellValue { .. } | ConditionalRule::Top10 { .. } => {
                    registry.register_cell_style(&cond_format.style)
                        .map_err(|e| WriteError::Validation(e))?;
                    let dxf_id = registry.register_dxf(&cond_format.style);
                    dxf_ids.insert(idx, dxf_id);
                }
                _ => {}
            }
        }
        updated_config.cond_format_dxf_ids = dxf_ids;
    }

    let mut zipper = ZipArchive::new();
    let sheet_names = vec![sheet_name];
    let charts_count = vec![config.charts.len()];
    // let images_data = vec![(config.images.clone(), if config.images.is_empty() { 0 } else { 1 })];
    let drawing_count = if config.charts.is_empty() && config.images.is_empty() { 0 } else { 1 };
    let images_data = vec![(config.images.clone(), drawing_count)];

    // Build the workbook's shared-string table (once) before emitting the static
    // parts, since [Content_Types].xml and workbook rels must reference the
    // sharedStrings part when it exists.
    let (sst, per_sheet_shared) = xml::build_shared_strings(&[batches]);
    let empty_shared: Vec<bool> = Vec::new();
    let shared_cols = per_sheet_shared.first().unwrap_or(&empty_shared);
    let has_sst = !sst.is_empty();

    add_static_files_ext(&mut zipper, &sheet_names, Some(&registry), &vec![config.tables.len()], &charts_count, &images_data, has_sst);

    if has_sst {
        zipper
            .add_file_from_memory(xml::generate_shared_strings_xml(&sst), "xl/sharedStrings.xml".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
    }

    let xml_data = xml::generate_sheet_xml_from_arrow(batches, &updated_config, &col_format_map, &cell_style_map, &sst, shared_cols)?;
    
    // DEBUG: Check for leading garbage
    // if xml_data.len() > 0 {
    //     eprintln!("First 100 bytes: {:?}", &xml_data[..xml_data.len().min(100)]);
    //     eprintln!("Starts with '<?xml': {}", xml_data.starts_with(b"<?xml"));
    // }

    
    zipper
        .add_file_from_memory(xml_data, "xl/worksheets/sheet1.xml".to_string())
        .compression_level(CompressionLevel::fast())
        .done();

    let hyperlinks_with_idx: Vec<(String, usize)> = config.hyperlinks
        .iter()
        .enumerate()
        .map(|(idx, h)| (h.url.clone(), idx + 1))
        .collect();
    
    let has_any_rels = !config.hyperlinks.is_empty() || !config.tables.is_empty() || !config.charts.is_empty() || !config.images.is_empty();
    
    if has_any_rels {
        let mut rels_xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
        
        for (url, idx) in &hyperlinks_with_idx {
            rels_xml.push_str(&format!("<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{}\" TargetMode=\"External\"/>\n", idx, xml::xml_escape_str(url)));
        }
        
        for idx in 0..config.tables.len() {
            rels_xml.push_str(&format!("<Relationship Id=\"rIdTable{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table{}.xml\"/>\n", idx + 1, idx + 1));
        }
        
        if !config.charts.is_empty() || !config.images.is_empty() {
            rels_xml.push_str("<Relationship Id=\"rIdDraw1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing1.xml\"/>\n");
        }
        
        rels_xml.push_str("</Relationships>");
        
        zipper
            .add_file_from_memory(rels_xml.into_bytes(), "xl/worksheets/_rels/sheet1.xml.rels".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
    }
    
    if !config.tables.is_empty() {
        // Calculate total rows once for all tables
        let total_data_rows: usize = batches.iter().map(|b| b.num_rows()).sum();
        let num_cols = if !batches.is_empty() { batches[0].schema().fields().len() } else { 0 };
        
        for (idx, table) in config.tables.iter().enumerate() {
            let table_id = (idx + 1) as u32;
            
            let mut adjusted_table = table.clone();
            
            // Auto-calculate end_row if not specified (0 means auto)
            if adjusted_table.range.2 == 0 {
                // end_row = start_row + num_data_rows - 1 (inclusive)
                adjusted_table.range.2 = adjusted_table.range.0 + total_data_rows;
            }
            
            // Auto-calculate end_col if not specified (0 means auto)
            if adjusted_table.range.3 == 0 {
                if num_cols > 0 {
                    adjusted_table.range.3 = adjusted_table.range.1 + num_cols - 1;
                }
            }
            
            // If table starts after row 1, we inserted a header row, so adjust end_row
            // Only adjust if user manually specified end_row (not auto-calculated)
            if adjusted_table.range.0 > 1 && table.range.2 != 0 {
                adjusted_table.range.2 += 1; // end_row++
            }
            
            let col_names = if table.column_names.is_empty() && !batches.is_empty() {
                let schema = batches[0].schema();
                let (_, start_col, _, end_col) = adjusted_table.range;
                // Clamp the column span to the actual schema width. A table whose
                // end_col exceeds the number of columns previously panicked on the
                // slice index -- and with panic="abort" that crashes the whole
                // Python process instead of raising. Clamp so any range is safe.
                let nfields = schema.fields().len();
                if nfields == 0 || start_col >= nfields {
                    Vec::new()
                } else {
                    let end = end_col.min(nfields - 1).max(start_col);
                    schema.fields()[start_col..=end]
                        .iter()
                        .map(|f| f.name().clone())
                        .collect()
                }
            } else {
                table.column_names.clone()
            };
            
            let table_xml = xml::generate_table_xml(&adjusted_table, table_id, &col_names);
            zipper
                .add_file_from_memory(
                    table_xml.into_bytes(),
                    format!("xl/tables/table{}.xml", table_id)
                )
                .compression_level(CompressionLevel::fast())
                .done();
        }
    }
    
    let has_drawing = !config.charts.is_empty() || !config.images.is_empty();
    
    if has_drawing {
        let drawing_xml = generate_drawing_xml_combined(&config.charts, &config.images);
        zipper
            .add_file_from_memory(drawing_xml.into_bytes(), "xl/drawings/drawing1.xml".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
        
        let drawing_rels = generate_drawing_rels_combined(config.charts.len(), &config.images, 1, 1);
        zipper
            .add_file_from_memory(drawing_rels.into_bytes(), "xl/drawings/_rels/drawing1.xml.rels".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
        
        for (idx, chart) in config.charts.iter().enumerate() {
            let chart_xml = xml::generate_chart_xml(chart, sheet_name);
            zipper
                .add_file_from_memory(
                    chart_xml.into_bytes(),
                    format!("xl/charts/chart{}.xml", idx + 1)
                )
                .compression_level(CompressionLevel::fast())
                .done();
        }
        
        // Add image files
        for (idx, image) in config.images.iter().enumerate() {
            zipper
                .add_file_from_memory(
                    image.image_data.clone(),
                    format!("xl/media/image{}.{}", idx + 1, image.extension)
                )
                .compression_level(CompressionLevel::fast())
                .done();
        }
    }

    write_zip_to_file(zipper, filename)
}

pub fn write_single_sheet_arrow_to_bytes(
    batches: &[RecordBatch],
    sheet_name: &str,
    config: &StyleConfig,
) -> Result<Vec<u8>, WriteError> {
    validate_sheet_name(sheet_name)?;

    // Guard against an empty batch slice BEFORE indexing batches[0] (see the
    // matching note in write_single_sheet_arrow_with_config).
    if batches.is_empty() {
        return Err(WriteError::Validation(
            "Cannot write sheet from an empty set of record batches".to_string(),
        ));
    }

    // Reject sheets exceeding Excel's row/column grid (O(1), off the hot path).
    {
        let rows: usize = batches.iter().map(|b| b.num_rows()).sum::<usize>()
            + if config.write_header_row { 1 } else { 0 };
        validate_sheet_dimensions(rows, batches[0].schema().fields().len())?;
    }

    let mut registry = StyleRegistry::new();
    let mut updated_config = config.clone();

    let schema = batches[0].schema();
    let col_format_map: HashMap<usize, u32> = if let Some(formats) = &config.column_formats {
        let mut map = HashMap::new();
        for (idx, field) in schema.fields().iter().enumerate() {
            if let Some(fmt) = formats.get(field.name()) {
                let cell_style = CellStyle {
                    font: None,
                    fill: None,
                    border: None,
                    alignment: None,
                    number_format: Some(fmt.clone()),
                };
                let style_id = registry.register_cell_style(&cell_style)
                    .map_err(|e| WriteError::Validation(e))?;
                map.insert(idx, style_id);
            }
        }
        map
    } else {
        HashMap::new()
    };

    let mut cell_style_map: HashMap<(usize, usize), u32> = HashMap::new();
    for cell_style in &config.cell_styles {
        let style_id = registry.register_cell_style(&cell_style.style)
            .map_err(|e| WriteError::Validation(e))?;
        cell_style_map.insert((cell_style.row, cell_style.col), style_id);
    }

    if !config.conditional_formats.is_empty() {
        let mut dxf_ids = HashMap::new();
        for (idx, cond_format) in config.conditional_formats.iter().enumerate() {
            match &cond_format.rule {
                ConditionalRule::CellValue { .. } | ConditionalRule::Top10 { .. } => {
                    registry.register_cell_style(&cond_format.style)
                        .map_err(|e| WriteError::Validation(e))?;
                    dxf_ids.insert(idx, idx);
                }
                _ => {}
            }
        }
        updated_config.conditional_formats = config.conditional_formats.clone();
    }

    let (sst, per_sheet_shared) = xml::build_shared_strings(&[batches]);
    let empty_shared: Vec<bool> = Vec::new();
    let shared_cols = per_sheet_shared.first().unwrap_or(&empty_shared);
    let has_sst = !sst.is_empty();

    let xml_data = xml::generate_sheet_xml_from_arrow(
        batches,
        &updated_config,
        &col_format_map,
        &cell_style_map,
        &sst,
        shared_cols,
    )?;

    let mut zipper = ZipArchive::new();
    let sheet_names = vec![sheet_name];
    let charts_count = vec![config.charts.len()];
    let drawing_count = if config.charts.is_empty() && config.images.is_empty() { 0 } else { 1 };
    
    add_static_files_ext(
        &mut zipper, 
        &sheet_names, 
        Some(&registry), 
        &[config.tables.len()], 
        &charts_count, 
        &[(config.images.clone(), drawing_count)],
        has_sst,
    );

    if has_sst {
        zipper
            .add_file_from_memory(xml::generate_shared_strings_xml(&sst), "xl/sharedStrings.xml".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
    }

    zipper
        .add_file_from_memory(xml_data, "xl/worksheets/sheet1.xml".to_string())
        .compression_level(CompressionLevel::fast())
        .done();

    // Emit the drawing part whenever there are charts OR images. Previously this
    // was gated on `!config.charts.is_empty()` and used the charts-only
    // `generate_drawing_xml`, so an images-only sheet wrote the media bytes and a
    // dangling drawing relationship but NO drawing1.xml -- the image never showed
    // and the reference dangled. Use the combined generator (charts + images),
    // matching the file-based `_with_config` path.
    let has_drawing = !config.charts.is_empty() || !config.images.is_empty();
    if has_drawing {
        let drawing_xml = generate_drawing_xml_combined(&config.charts, &config.images);
        zipper
            .add_file_from_memory(drawing_xml.into_bytes(), "xl/drawings/drawing1.xml".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
        
        let drawing_rels = generate_drawing_rels_combined(config.charts.len(), &config.images, 1, 1);
        zipper
            .add_file_from_memory(drawing_rels.into_bytes(), "xl/drawings/_rels/drawing1.xml.rels".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
        
        for (idx, chart) in config.charts.iter().enumerate() {
            let chart_xml = xml::generate_chart_xml(chart, sheet_name);
            zipper
                .add_file_from_memory(
                    chart_xml.into_bytes(),
                    format!("xl/charts/chart{}.xml", idx + 1)
                )
                .compression_level(CompressionLevel::fast())
                .done();
        }
    }

    if !config.tables.is_empty() {
        for (idx, table) in config.tables.iter().enumerate() {
            let col_names = if table.column_names.is_empty() {
                let (_, start_col, _, end_col) = table.range;
                let nfields = schema.fields().len();
                if nfields == 0 || start_col >= nfields {
                    Vec::new()
                } else {
                    let end = end_col.min(nfields - 1).max(start_col);
                    schema.fields()[start_col..=end]
                        .iter()
                        .map(|f| f.name().clone())
                        .collect()
                }
            } else {
                table.column_names.clone()
            };
            
            let table_xml = xml::generate_table_xml(table, (idx + 1) as u32, &col_names);
            zipper
                .add_file_from_memory(
                    table_xml.into_bytes(),
                    format!("xl/tables/table{}.xml", idx + 1)
                )
                .compression_level(CompressionLevel::fast())
                .done();
        }
    }

    if !config.images.is_empty() {
        for (idx, image) in config.images.iter().enumerate() {
            zipper
                .add_file_from_memory(
                    image.image_data.clone(),
                    format!("xl/media/image{}.{}", idx + 1, image.extension)
                )
                .compression_level(CompressionLevel::fast())
                .done();
        }
    }

    // One worksheet-rels file covering every relationship the sheet XML
    // references: hyperlinks (rId{n}), tables (rIdTable{n}), and the drawing
    // (rIdDraw1). Writing this once at the end -- rather than in each per-feature
    // block -- avoids both the missing-hyperlink-rels bug and the
    // double-write collision when a sheet had both tables and charts.
    let has_any_rels = !config.hyperlinks.is_empty()
        || !config.tables.is_empty()
        || !config.charts.is_empty()
        || !config.images.is_empty();

    if has_any_rels {
        let mut rels_xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");

        for (idx, h) in config.hyperlinks.iter().enumerate() {
            rels_xml.push_str(&format!("<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{}\" TargetMode=\"External\"/>\n", idx + 1, xml::xml_escape_str(&h.url)));
        }

        for idx in 0..config.tables.len() {
            rels_xml.push_str(&format!("<Relationship Id=\"rIdTable{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table{}.xml\"/>\n", idx + 1, idx + 1));
        }

        if !config.charts.is_empty() || !config.images.is_empty() {
            rels_xml.push_str("<Relationship Id=\"rIdDraw1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing1.xml\"/>\n");
        }

        rels_xml.push_str("</Relationships>");

        zipper
            .add_file_from_memory(rels_xml.into_bytes(), "xl/worksheets/_rels/sheet1.xml.rels".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
    }

    write_zip_to_buffer(zipper)
}

pub fn write_multiple_sheets_arrow_to_bytes(
    sheets: &[(Vec<RecordBatch>, &str, StyleConfig)],
    num_threads: usize,
) -> Result<Vec<u8>, WriteError> {
    for (batches, sheet_name, _) in sheets {
        validate_sheet_name(sheet_name)?;
        if batches.is_empty() {
            return Err(WriteError::Validation(format!(
                "Sheet '{}' has an empty set of record batches",
                sheet_name
            )));
        }
    }
    let names: Vec<&str> = sheets.iter().map(|(_, n, _)| *n).collect();
    validate_unique_sheet_names(&names)?;

    // Optional phase profiler: set JETXL_PROFILE=1 to print per-phase timings to
    // stderr. Lets scaling be diagnosed on real multi-core hardware (which phase
    // fails to speed up as threads increase) rather than guessed at.
    let mut prof = PhaseProfiler::new();

    // Build one workbook-global style registry plus the per-sheet
    // column-format / cell-style / dxf maps BEFORE the parallel worksheet pass.
    // This mirrors the file-based `write_multiple_sheets_arrow_with_configs`
    // path. Previously this bytes path used a helper that passed an EMPTY
    // cell_style_map and never registered conditional-format dxfs, so custom
    // cell styles and conditional formatting silently vanished on multi-sheet
    // in-memory writes while working on the single-sheet and file paths. Doing
    // the registration up front (serially) keeps the parallel pass doing only
    // read-only lookups -- no shared mutable registry, so no added contention.
    let mut style_registry = StyleRegistry::new();
    let mut sheet_col_format_maps: Vec<HashMap<usize, u32>> = Vec::with_capacity(sheets.len());
    let mut sheet_cell_style_maps: Vec<HashMap<(usize, usize), u32>> = Vec::with_capacity(sheets.len());
    let mut sheet_dxf_mappings: Vec<HashMap<usize, u32>> = Vec::with_capacity(sheets.len());

    for (batches, _, config) in sheets {
        let schema = batches[0].schema();
        // Reject any sheet exceeding Excel's row/column grid (O(1) per sheet).
        {
            let rows: usize = batches.iter().map(|b| b.num_rows()).sum::<usize>()
                + if config.write_header_row { 1 } else { 0 };
            validate_sheet_dimensions(rows, schema.fields().len())?;
        }
        let mut col_format_map = HashMap::new();
        if let Some(formats) = &config.column_formats {
            for (idx, field) in schema.fields().iter().enumerate() {
                if let Some(fmt) = formats.get(field.name()) {
                    let cell_style = CellStyle {
                        font: None,
                        fill: None,
                        border: None,
                        alignment: None,
                        number_format: Some(fmt.clone()),
                    };
                    let style_id = style_registry.register_cell_style(&cell_style)
                        .map_err(WriteError::Validation)?;
                    col_format_map.insert(idx, style_id);
                }
            }
        }
        sheet_col_format_maps.push(col_format_map);

        let mut cell_style_map: HashMap<(usize, usize), u32> = HashMap::new();
        for cell_style in &config.cell_styles {
            let style_id = style_registry.register_cell_style(&cell_style.style)
                .map_err(WriteError::Validation)?;
            cell_style_map.insert((cell_style.row, cell_style.col), style_id);
        }
        sheet_cell_style_maps.push(cell_style_map);

        let mut dxf_ids = HashMap::new();
        for (idx, cond_format) in config.conditional_formats.iter().enumerate() {
            match &cond_format.rule {
                ConditionalRule::CellValue { .. } | ConditionalRule::Top10 { .. } => {
                    style_registry.register_cell_style(&cond_format.style)
                        .map_err(WriteError::Validation)?;
                    let dxf_id = style_registry.register_dxf(&cond_format.style);
                    dxf_ids.insert(idx, dxf_id);
                }
                _ => {}
            }
        }
        sheet_dxf_mappings.push(dxf_ids);
    }
    prof.mark("prepass_styles");

    // Build the ONE workbook-global shared-string table before the parallel
    // worksheet pass, so each thread only does read-only lookups (no locking, no
    // index remap). per_sheet_shared[i] marks which columns of sheet i are shared.
    let batch_refs: Vec<&[RecordBatch]> = sheets.iter().map(|(b, _, _)| b.as_slice()).collect();
    let (sst, per_sheet_shared) = xml::build_shared_strings(&batch_refs);
    let has_sst = !sst.is_empty();
    prof.mark("prepass_sst");

    // Pipelined worksheet pass: each parallel task generates a sheet's XML AND
    // compresses it in the same step, so compression overlaps generation instead
    // of running as a separate phase after all XML is done. Each task returns a
    // ready-to-place compressed part. This removes the barrier between the
    // "generate all XML" and "compress all XML" phases that previously serialized
    // the two and capped multi-core scaling.
    let ws_parts: Vec<crate::fastzip::PreparedPart> = if num_threads > 1 && sheets.len() > 1 {
        let pool = cached_pool(num_threads)?;

        pool.install(|| {
            sheets
                .par_iter()
                .enumerate()
                .map(|(sheet_idx, (batches, _, config))| {
                    let mut modified_config = (*config).clone();
                    modified_config.cond_format_dxf_ids = sheet_dxf_mappings[sheet_idx].clone();
                    let col_format_map = &sheet_col_format_maps[sheet_idx];
                    let cell_style_map = &sheet_cell_style_maps[sheet_idx];
                    let shared_cols = per_sheet_shared.get(sheet_idx).map(|v| v.as_slice()).unwrap_or(&[]);
                    let xml = xml::generate_sheet_xml_from_arrow(batches, &modified_config, col_format_map, cell_style_map, &sst, shared_cols)?;
                    Ok(crate::fastzip::compress_part(
                        format!("xl/worksheets/sheet{}.xml", sheet_idx + 1),
                        xml,
                        CompressionLevel::fast(),
                    ))
                })
                .collect::<Result<Vec<_>, WriteError>>()
        })?
    } else {
        sheets
            .iter()
            .enumerate()
            .map(|(sheet_idx, (batches, _, config))| {
                let mut modified_config = (*config).clone();
                modified_config.cond_format_dxf_ids = sheet_dxf_mappings[sheet_idx].clone();
                let col_format_map = &sheet_col_format_maps[sheet_idx];
                let cell_style_map = &sheet_cell_style_maps[sheet_idx];
                let shared_cols = per_sheet_shared.get(sheet_idx).map(|v| v.as_slice()).unwrap_or(&[]);
                let xml = xml::generate_sheet_xml_from_arrow(batches, &modified_config, col_format_map, cell_style_map, &sst, shared_cols)?;
                Ok(crate::fastzip::compress_part(
                    format!("xl/worksheets/sheet{}.xml", sheet_idx + 1),
                    xml,
                    CompressionLevel::fast(),
                ))
            })
            .collect::<Result<Vec<_>, WriteError>>()?
    };
    prof.mark("gen_compress_ws");

    let mut zipper = ZipArchive::new();
    let sheet_names: Vec<&str> = sheets.iter().map(|(_, name, _)| *name).collect();
    let tables_count: Vec<usize> = sheets.iter().map(|(_, _, config)| config.tables.len()).collect();
    let charts_count: Vec<usize> = sheets.iter().map(|(_, _, config)| config.charts.len()).collect();
    let images_data: Vec<(Vec<ExcelImage>, usize)> = sheets.iter().map(|(_, _, config)| {
        let drawing_count = if config.charts.is_empty() && config.images.is_empty() { 0 } else { 1 };
        (config.images.clone(), drawing_count)
    }).collect();

    // Pass the populated style registry (was previously None) so the workbook's
    // styles.xml carries the custom cell styles and conditional-format dxfs we
    // registered above.
    add_static_files_ext(&mut zipper, &sheet_names, Some(&style_registry), &tables_count, &charts_count, &images_data, has_sst);

    if has_sst {
        zipper
            .add_file_from_memory(xml::generate_shared_strings_xml(&sst), "xl/sharedStrings.xml".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
    }

    for part in ws_parts {
        zipper.add_prepared(part);
    }

    let mut global_chart_id = 1;
    let mut global_table_id = 1;
    let mut global_image_id = 1;
    let mut drawing_id = 1;

    for (idx, (_, _, sheet_config)) in sheets.iter().enumerate() {
        let has_charts = !sheet_config.charts.is_empty();
        let has_tables = !sheet_config.tables.is_empty();
        let has_hyperlinks = !sheet_config.hyperlinks.is_empty();
        let has_images = !sheet_config.images.is_empty();

        // A worksheet rels part is needed if the sheet references ANY external
        // relationship: hyperlinks (rId{n}), tables (rIdTable{n}) or a drawing
        // (rIdDraw1). Previously hyperlinks were omitted here, so a multi-sheet
        // in-memory write with hyperlinks emitted <hyperlink r:id="rId1"> in the
        // sheet but no matching relationship -- openpyxl (and Excel) then error
        // with "Unknown relationship: rId1". The hyperlink r:id namespace
        // (rId{n}) is distinct from rIdTable{n} / rIdDraw1, so there's no clash.
        if has_tables || has_charts || has_images || has_hyperlinks {
            let mut rels_xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");

            for (h_idx, h) in sheet_config.hyperlinks.iter().enumerate() {
                rels_xml.push_str(&format!("<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{}\" TargetMode=\"External\"/>\n", h_idx + 1, xml::xml_escape_str(&h.url)));
            }

            if has_tables {
                for table_idx in 0..sheet_config.tables.len() {
                    rels_xml.push_str(&format!("<Relationship Id=\"rIdTable{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table{}.xml\"/>\n", table_idx + 1, global_table_id + table_idx));
                }
            }
            
            if has_charts || has_images {
                rels_xml.push_str(&format!("<Relationship Id=\"rIdDraw1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing{}.xml\"/>\n", drawing_id));
            }
            
            rels_xml.push_str("</Relationships>");
            zipper
                .add_file_from_memory(rels_xml.into_bytes(), format!("xl/worksheets/_rels/sheet{}.xml.rels", idx + 1))
                .compression_level(CompressionLevel::fast())
                .done();
        }
        
        if has_tables {
            let total_data_rows: usize = sheets[idx].0.iter().map(|b| b.num_rows()).sum();
            let num_cols = if !sheets[idx].0.is_empty() { 
                sheets[idx].0[0].schema().fields().len() 
            } else { 
                0 
            };
            
            for table in &sheet_config.tables {
                let mut adjusted_table = table.clone();
                
                if adjusted_table.range.2 == 0 {
                    adjusted_table.range.2 = adjusted_table.range.0 + total_data_rows;
                }
                
                if adjusted_table.range.3 == 0 {
                    if num_cols > 0 {
                        adjusted_table.range.3 = adjusted_table.range.1 + num_cols - 1;
                    }
                }
                
                if adjusted_table.range.0 > 1 && table.range.2 != 0 {
                    adjusted_table.range.2 += 1;
                }
                
                let col_names = if table.column_names.is_empty() && !sheets[idx].0.is_empty() {
                    let schema = sheets[idx].0[0].schema();
                    let (_, start_col, _, end_col) = adjusted_table.range;
                    let nfields = schema.fields().len();
                    if nfields == 0 || start_col >= nfields {
                        Vec::new()
                    } else {
                        let end = end_col.min(nfields - 1).max(start_col);
                        schema.fields()[start_col..=end]
                            .iter()
                            .map(|f| f.name().clone())
                            .collect()
                    }
                } else {
                    table.column_names.clone()
                };
                
                let table_xml = xml::generate_table_xml(&adjusted_table, global_table_id as u32, &col_names);
                zipper
                    .add_file_from_memory(
                        table_xml.into_bytes(),
                        format!("xl/tables/table{}.xml", global_table_id)
                    )
                    .compression_level(CompressionLevel::fast())
                    .done();
                global_table_id += 1;
            }
        }
        
        let has_images = !sheet_config.images.is_empty();
        if has_charts || has_images {
            let drawing_xml = generate_drawing_xml_combined(&sheet_config.charts, &sheet_config.images);
            zipper
                .add_file_from_memory(drawing_xml.into_bytes(), format!("xl/drawings/drawing{}.xml", drawing_id))
                .compression_level(CompressionLevel::fast())
                .done();
            
            let drawing_rels = generate_drawing_rels_combined(sheet_config.charts.len(), &sheet_config.images, global_chart_id, global_image_id);
            
            zipper
                .add_file_from_memory(drawing_rels.into_bytes(), format!("xl/drawings/_rels/drawing{}.xml.rels", drawing_id))
                .compression_level(CompressionLevel::fast())
                .done();
            
            for chart in &sheet_config.charts {
                let chart_xml = xml::generate_chart_xml(chart, sheets[idx].1);
                zipper
                    .add_file_from_memory(
                        chart_xml.into_bytes(),
                        format!("xl/charts/chart{}.xml", global_chart_id)
                    )
                    .compression_level(CompressionLevel::fast())
                    .done();
                global_chart_id += 1;
            }
            
            // Media file names must be workbook-global so images on different
            // sheets don't collide on xl/media/image1.png. global_image_id must
            // advance in lockstep with the Target written by the drawing rels.
            for image in &sheet_config.images {
                zipper
                    .add_file_from_memory(
                        image.image_data.clone(),
                        format!("xl/media/image{}.{}", global_image_id, image.extension)
                    )
                    .compression_level(CompressionLevel::fast())
                    .done();
                global_image_id += 1;
            }
            
            drawing_id += 1;
        }
    }
    prof.mark("zip_metadata");

    let result = write_zip_to_buffer(zipper);
    prof.mark("assemble");
    prof.report(num_threads);
    result
}


// Superseded by write_multiple_sheets_arrow_with_configs / _to_bytes; retained
// for API completeness.
#[allow(dead_code)]
pub fn write_multiple_sheets_arrow(
    sheets: &[(Vec<RecordBatch>, String)],
    filename: &str,
    num_threads: usize,
) -> Result<(), WriteError> {
    write_multiple_sheets_arrow_with_configs(
        &sheets.iter().map(|(b, n)| (b.as_slice(), n.as_str(), StyleConfig::default())).collect::<Vec<_>>(),
        filename,
        num_threads,
    )
}

pub fn write_multiple_sheets_arrow_with_configs(
    sheets: &[(&[RecordBatch], &str, StyleConfig)],
    filename: &str,
    num_threads: usize,
) -> Result<(), WriteError> {
    for (batches, name, _) in sheets {
        validate_sheet_name(name)?;
        // Guard before any batches[0] indexing below (avoids a process-aborting
        // panic on an empty sheet under panic="abort").
        if batches.is_empty() {
            return Err(WriteError::Validation(format!(
                "Sheet '{}' has an empty set of record batches",
                name
            )));
        }
    }
    let names: Vec<&str> = sheets.iter().map(|(_, n, _)| *n).collect();
    validate_unique_sheet_names(&names)?;

    let mut style_registry = StyleRegistry::new();
    let mut sheet_col_format_maps = Vec::new();
    let mut sheet_cell_style_maps = Vec::new();
    let mut sheet_dxf_mappings = Vec::new();

    for (batches, _, config) in sheets {
        let schema = batches[0].schema();
        // Reject any sheet exceeding Excel's row/column grid (O(1) per sheet).
        {
            let rows: usize = batches.iter().map(|b| b.num_rows()).sum::<usize>()
                + if config.write_header_row { 1 } else { 0 };
            validate_sheet_dimensions(rows, schema.fields().len())?;
        }
        let mut col_format_map = HashMap::new();
        if let Some(formats) = &config.column_formats {
            for (idx, field) in schema.fields().iter().enumerate() {
                if let Some(fmt) = formats.get(field.name()) {
                    let cell_style = CellStyle {
                        font: None,
                        fill: None,
                        border: None,
                        alignment: None,
                        number_format: Some(fmt.clone()),
                    };
                    let style_id = style_registry.register_cell_style(&cell_style)
                        .map_err(|e| WriteError::Validation(e))?;
                    col_format_map.insert(idx, style_id);
                }
            }
        }
        sheet_col_format_maps.push(col_format_map);

        // Build cell style map for this sheet
        let mut cell_style_map: HashMap<(usize, usize), u32> = HashMap::new();
        for cell_style in &config.cell_styles {
            let style_id = style_registry.register_cell_style(&cell_style.style)
                .map_err(|e| WriteError::Validation(e))?;
            cell_style_map.insert((cell_style.row, cell_style.col), style_id);
        }
        sheet_cell_style_maps.push(cell_style_map);

        let mut dxf_ids = HashMap::new();
        for (idx, cond_format) in config.conditional_formats.iter().enumerate() {
            match &cond_format.rule {
                ConditionalRule::CellValue { .. } | ConditionalRule::Top10 { .. } => {
                    style_registry.register_cell_style(&cond_format.style)
                        .map_err(|e| WriteError::Validation(e))?;
                    let dxf_id = style_registry.register_dxf(&cond_format.style);
                    dxf_ids.insert(idx, dxf_id);
                }
                _ => {}
            }
        }
        sheet_dxf_mappings.push(dxf_ids);
    }

    // Build the workbook-global shared-string table once, before the parallel
    // worksheet pass (read-only lookups per thread; no locking, no remap).
    let batch_refs: Vec<&[RecordBatch]> = sheets.iter().map(|(b, _, _)| *b).collect();
    let (sst, per_sheet_shared) = xml::build_shared_strings(&batch_refs);
    let has_sst = !sst.is_empty();

    // Pipelined worksheet pass (file path): fuse XML generation with compression
    // so the two overlap instead of running as separate phases. Each task returns
    // the compressed worksheet part plus that sheet's hyperlink list (needed for
    // rels later). See the to_bytes path for the rationale.
    let ws_parts_and_links: Vec<(crate::fastzip::PreparedPart, Vec<(String, usize)>)> =
        if num_threads > 1 && sheets.len() > 1 {
            let pool = cached_pool(num_threads)?;

            pool.install(|| {
                sheets
                    .par_iter()
                    .enumerate()
                    .map(|(sheet_idx, (batches, _, config))| {
                        let mut modified_config = (*config).clone();
                        if sheet_idx < sheet_dxf_mappings.len() {
                            modified_config.cond_format_dxf_ids = sheet_dxf_mappings[sheet_idx].clone();
                        }

                        let col_format_map = &sheet_col_format_maps[sheet_idx];
                        let cell_style_map = &sheet_cell_style_maps[sheet_idx];
                        let shared_cols = per_sheet_shared.get(sheet_idx).map(|v| v.as_slice()).unwrap_or(&[]);
                        let xml_data = xml::generate_sheet_xml_from_arrow(batches, &modified_config, col_format_map, cell_style_map, &sst, shared_cols)?;
                        let hyperlinks: Vec<(String, usize)> = modified_config.hyperlinks
                            .iter()
                            .enumerate()
                            .map(|(i, h)| (h.url.clone(), i + 1))
                            .collect();
                        let part = crate::fastzip::compress_part(
                            format!("xl/worksheets/sheet{}.xml", sheet_idx + 1),
                            xml_data,
                            CompressionLevel::fast(),
                        );
                        Ok((part, hyperlinks))
                    })
                    .collect::<Result<Vec<_>, WriteError>>()
            })?
        } else {
            sheets
                .iter()
                .enumerate()
                .map(|(sheet_idx, (batches, _, config))| {
                    let mut modified_config = (*config).clone();
                    if sheet_idx < sheet_dxf_mappings.len() {
                        modified_config.cond_format_dxf_ids = sheet_dxf_mappings[sheet_idx].clone();
                    }

                    let col_format_map = &sheet_col_format_maps[sheet_idx];
                    let cell_style_map = &sheet_cell_style_maps[sheet_idx];
                    let shared_cols = per_sheet_shared.get(sheet_idx).map(|v| v.as_slice()).unwrap_or(&[]);
                    let xml_data = xml::generate_sheet_xml_from_arrow(batches, &modified_config, col_format_map, cell_style_map, &sst, shared_cols)?;
                    let hyperlinks: Vec<(String, usize)> = modified_config.hyperlinks
                        .iter()
                        .enumerate()
                        .map(|(i, h)| (h.url.clone(), i + 1))
                        .collect();
                    let part = crate::fastzip::compress_part(
                        format!("xl/worksheets/sheet{}.xml", sheet_idx + 1),
                        xml_data,
                        CompressionLevel::fast(),
                    );
                    Ok((part, hyperlinks))
                })
                .collect::<Result<Vec<_>, WriteError>>()?
        };

    let mut zipper = ZipArchive::new();
    let sheet_names: Vec<&str> = sheets.iter().map(|(_, name, _)| *name).collect();
    let tables_per_sheet: Vec<usize> = sheets.iter().map(|(_, _, cfg)| cfg.tables.len()).collect();
    let charts_per_sheet: Vec<usize> = sheets.iter().map(|(_, _, cfg)| cfg.charts.len()).collect();

    let images_per_sheet: Vec<(Vec<ExcelImage>, usize)> = sheets.iter()
            .map(|(_, _, cfg)| {
                // count drawing if charts OR images exist
                let count = if cfg.charts.is_empty() && cfg.images.is_empty() { 0 } else { 1 };
                (cfg.images.clone(), count)
            })
            .collect();
    add_static_files_ext(&mut zipper, &sheet_names, Some(&style_registry), &tables_per_sheet, &charts_per_sheet, &images_per_sheet, has_sst);

    if has_sst {
        zipper
            .add_file_from_memory(xml::generate_shared_strings_xml(&sst), "xl/sharedStrings.xml".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
    }

    let mut global_chart_id = 1;
    let mut global_table_id = 1;
    let mut global_image_id = 1;
    let mut drawing_id = 1;

    for (idx, (ws_part, hyperlinks)) in ws_parts_and_links.into_iter().enumerate() {
        let sheet_config = &sheets[idx].2;

        zipper.add_prepared(ws_part);

        let has_hyperlinks = !hyperlinks.is_empty();
        let has_tables = !sheet_config.tables.is_empty();
        let has_charts = !sheet_config.charts.is_empty();
        let has_images = !sheet_config.images.is_empty();

        if has_hyperlinks || has_tables || has_charts || has_images {
            let mut rels_xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
            
            for (url, rid) in &hyperlinks {
                rels_xml.push_str(&format!("<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{}\" TargetMode=\"External\"/>\n", rid, xml::xml_escape_str(url)));
            }
            
            let sheet_start_table_id = global_table_id;
            for i in 0..sheet_config.tables.len() {
                rels_xml.push_str(&format!("<Relationship Id=\"rIdTable{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table{}.xml\"/>\n", 
                    i + 1, 
                    sheet_start_table_id + i));
            }
            
            if has_charts || has_images {
                rels_xml.push_str(&format!("<Relationship Id=\"rIdDraw1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing{}.xml\"/>\n", drawing_id));
            }
            
            rels_xml.push_str("</Relationships>");
            
            zipper
                .add_file_from_memory(
                    rels_xml.into_bytes(),
                    format!("xl/worksheets/_rels/sheet{}.xml.rels", idx + 1)
                )
                .compression_level(CompressionLevel::fast())
                .done();
        }
        
        if has_tables {
            // Calculate total rows and cols for this sheet
            let total_data_rows: usize = sheets[idx].0.iter().map(|b| b.num_rows()).sum();
            let num_cols = if !sheets[idx].0.is_empty() { 
                sheets[idx].0[0].schema().fields().len() 
            } else { 
                0 
            };
            
            for table in &sheet_config.tables {
                let mut adjusted_table = table.clone();
                
                // Auto-calculate end_row if not specified (0 means auto)
                if adjusted_table.range.2 == 0 {
                    // end_row = start_row + num_data_rows - 1 (inclusive)
                    adjusted_table.range.2 = adjusted_table.range.0 + total_data_rows;
                }
                
                // Auto-calculate end_col if not specified (0 means auto)
                if adjusted_table.range.3 == 0 {
                    if num_cols > 0 {
                        adjusted_table.range.3 = adjusted_table.range.1 + num_cols - 1;
                    }
                }
                
                // If table starts after row 1, we inserted a header row, so adjust end_row
                // Only adjust if user manually specified end_row (not auto-calculated)
                if adjusted_table.range.0 > 1 && table.range.2 != 0 {
                    adjusted_table.range.2 += 1; // end_row++
                }
                
                let col_names = if table.column_names.is_empty() && !sheets[idx].0.is_empty() {
                    let schema = sheets[idx].0[0].schema();
                    let (_, start_col, _, end_col) = adjusted_table.range;
                    let nfields = schema.fields().len();
                    if nfields == 0 || start_col >= nfields {
                        Vec::new()
                    } else {
                        let end = end_col.min(nfields - 1).max(start_col);
                        schema.fields()[start_col..=end]
                            .iter()
                            .map(|f| f.name().clone())
                            .collect()
                    }
                } else {
                    table.column_names.clone()
                };
                
                let table_xml = xml::generate_table_xml(&adjusted_table, global_table_id as u32, &col_names);
                zipper
                    .add_file_from_memory(
                        table_xml.into_bytes(),
                        format!("xl/tables/table{}.xml", global_table_id)
                    )
                    .compression_level(CompressionLevel::fast())
                    .done();
                global_table_id += 1;
            }
        }
        
        let has_images = !sheet_config.images.is_empty();
        if has_charts || has_images {
            let drawing_xml = generate_drawing_xml_combined(&sheet_config.charts, &sheet_config.images);
            zipper
                .add_file_from_memory(drawing_xml.into_bytes(), format!("xl/drawings/drawing{}.xml", drawing_id))
                .compression_level(CompressionLevel::fast())
                .done();
            
            let drawing_rels = generate_drawing_rels_combined(sheet_config.charts.len(), &sheet_config.images, global_chart_id, global_image_id);
            
            zipper
                .add_file_from_memory(drawing_rels.into_bytes(), format!("xl/drawings/_rels/drawing{}.xml.rels", drawing_id))
                .compression_level(CompressionLevel::fast())
                .done();
            
            for chart in &sheet_config.charts {
                let chart_xml = xml::generate_chart_xml(chart, sheets[idx].1);
                zipper
                    .add_file_from_memory(
                        chart_xml.into_bytes(),
                        format!("xl/charts/chart{}.xml", global_chart_id)
                    )
                    .compression_level(CompressionLevel::fast())
                    .done();
                global_chart_id += 1;
            }
            // Add image files with workbook-GLOBAL names so images on different
            // sheets don't collide on xl/media/image1.png. Must stay in lockstep
            // with the Target emitted by generate_drawing_rels_combined.
            for image in &sheet_config.images {
                zipper
                    .add_file_from_memory(
                        image.image_data.clone(),
                        format!("xl/media/image{}.{}", global_image_id, image.extension)
                    )
                    .compression_level(CompressionLevel::fast())
                    .done();
                global_image_id += 1;
            }
            
            drawing_id += 1;
        }
    }

    write_zip_to_file(zipper, filename)
}

// ============================================================================
// Helper functions
// ============================================================================

fn add_static_files(
    zipper: &mut ZipArchive,
    sheet_names: &[&str],
    style_registry: Option<&StyleRegistry>,
    tables_count: &[usize], // Number of tables per sheet
    charts_count: &[usize],
    images_data: &[(Vec<ExcelImage>, usize)],
) {
    // Back-compat wrapper: paths that never emit shared strings (dict API,
    // legacy) call this. It forwards to the extended version with has_sst=false.
    add_static_files_ext(
        zipper, sheet_names, style_registry, tables_count, charts_count, images_data, false,
    );
}

fn add_static_files_ext(
    zipper: &mut ZipArchive,
    sheet_names: &[&str],
    style_registry: Option<&StyleRegistry>,
    tables_count: &[usize], // Number of tables per sheet
    charts_count: &[usize],
    images_data: &[(Vec<ExcelImage>, usize)],
    has_shared_strings: bool,
) {
    let images_per_sheet: Vec<(&[ExcelImage], usize)> = images_data.iter()
            .map(|(imgs, count)| (imgs.as_slice(), *count))
            .collect();
        
        zipper
            .add_file_from_memory(
                xml::generate_content_types_with_charts_ext(sheet_names, tables_count, charts_count, &images_per_sheet, has_shared_strings).into_bytes(),
                "[Content_Types].xml".to_string(),
            )
            .compression_level(CompressionLevel::fast())
            .done();
    
    zipper
        .add_file_from_memory(
            xml::generate_rels().as_bytes().to_vec(),
            "_rels/.rels".to_string(),
        )
        .compression_level(CompressionLevel::fast())
        .done();
    
    // Add document properties
    zipper
        .add_file_from_memory(
            xml::generate_core_xml().as_bytes().to_vec(),
            "docProps/core.xml".to_string(),
        )
        .compression_level(CompressionLevel::fast())
        .done();
    
    zipper
        .add_file_from_memory(
            xml::generate_app_xml(sheet_names).into_bytes(),
            "docProps/app.xml".to_string(),
        )
        .compression_level(CompressionLevel::fast())
        .done();
    
    zipper
        .add_file_from_memory(
            xml::generate_workbook(sheet_names).into_bytes(),
            "xl/workbook.xml".to_string(),
        )
        .compression_level(CompressionLevel::fast())
        .done();
    
    zipper
        .add_file_from_memory(
            xml::generate_workbook_rels_ext(sheet_names.len(), has_shared_strings).into_bytes(),
            "xl/_rels/workbook.xml.rels".to_string(),
        )
        .compression_level(CompressionLevel::fast())
        .done();
    
    let styles_xml = if let Some(registry) = style_registry {
        generate_styles_xml_enhanced(registry)
    } else {
        generate_styles_xml()
    };
    
    zipper
        .add_file_from_memory(
            styles_xml.into_bytes(),
            "xl/styles.xml".to_string(),
        )
        .compression_level(CompressionLevel::fast())
        .done();
}

fn write_zip_to_file(mut zipper: ZipArchive, filename: &str) -> Result<(), WriteError> {
    // Assemble the archive with the same pre-sized, single-allocation path the
    // in-memory writer uses, then write it to disk in one bulk `write_all`.
    //
    // The earlier approach streamed each part through a BufWriter, which for a
    // multi-MB worksheet meant many write calls and, combined with a trailing
    // fsync, added a per-file tax that grew with archive size (~a few percent at
    // 1M rows). One buffer + one write is both simpler and faster. We flush to
    // hand the bytes to the OS but do NOT `sync_all()`: forcing a physical disk
    // commit is unnecessary for producing a file and is pure added latency.
    let buffer = zipper.write_to_vec();
    let mut file = File::create(filename)?;
    file.write_all(&buffer)?;
    file.flush()?;
    Ok(())
}

fn write_zip_to_buffer(mut zipper: ZipArchive) -> Result<Vec<u8>, WriteError> {
    // fastzip assembles directly into a correctly pre-sized Vec, avoiding both
    // the Cursor-driven regrowth and a redundant full-archive copy.
    Ok(zipper.write_to_vec())
}

/// Excel's .xlsx grid is hard-capped at 1,048,576 rows and 16,384 columns.
/// Writing beyond either produces a file Excel refuses to open (or silently
/// truncates), so we reject up front with a clear error rather than emit an
/// invalid workbook. This is O(1) — two integer comparisons per sheet, well off
/// the per-cell hot path — so it costs nothing measurable even at 1M rows.
pub const EXCEL_MAX_ROWS: usize = 1_048_576;
pub const EXCEL_MAX_COLS: usize = 16_384;

fn validate_sheet_dimensions(total_rows_incl_header: usize, num_cols: usize) -> Result<(), WriteError> {
    if total_rows_incl_header > EXCEL_MAX_ROWS {
        return Err(WriteError::Validation(format!(
            "Sheet has {} rows (incl. header) which exceeds Excel's maximum of {} rows",
            total_rows_incl_header, EXCEL_MAX_ROWS
        )));
    }
    if num_cols > EXCEL_MAX_COLS {
        return Err(WriteError::Validation(format!(
            "Sheet has {} columns which exceeds Excel's maximum of {} columns",
            num_cols, EXCEL_MAX_COLS
        )));
    }
    Ok(())
}

fn validate_sheet_name(name: &str) -> Result<(), WriteError> {
    // Excel counts the 31-character limit in UNICODE CHARACTERS, not bytes.
    // Using name.len() (bytes) wrongly rejected legal names such as a 20-glyph
    // Hebrew or CJK title, and wrongly accepted nothing it should have.
    if name.chars().count() > 31 {
        return Err(WriteError::Validation(format!("Sheet name '{}' exceeds 31 characters", name)));
    }
    if name.is_empty() {
        return Err(WriteError::Validation("Sheet name cannot be empty".to_string()));
    }
    // Excel also forbids a leading or trailing apostrophe and these characters.
    if name.chars().any(|c| "[]':*?/\\".contains(c)) {
        return Err(WriteError::Validation(format!("Sheet name '{}' contains invalid chars", name)));
    }
    // The history/quote name is reserved by Excel.
    if name.eq_ignore_ascii_case("History") {
        return Err(WriteError::Validation("Sheet name 'History' is reserved by Excel".to_string()));
    }
    Ok(())
}

/// Excel requires every worksheet name to be unique, compared
/// case-insensitively ("Data" and "data" collide). Duplicate names silently
/// corrupted the workbook before; catch it up front with a clear error.
fn validate_unique_sheet_names(names: &[&str]) -> Result<(), WriteError> {
    let mut seen: HashSet<String> = HashSet::with_capacity(names.len());
    for name in names {
        let key = name.to_lowercase();
        if !seen.insert(key) {
            return Err(WriteError::Validation(format!(
                "Duplicate sheet name '{}' (names must be unique, case-insensitively)",
                name
            )));
        }
    }
    Ok(())
}