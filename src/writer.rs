use crate::types::{SheetData, WriteError};
use crate::styles::{StyleConfig, generate_styles_xml, generate_styles_xml_enhanced, StyleRegistry, ConditionalRule, CellStyle, ExcelImage};
// use crate::xml::{self, generate_drawing_xml_combined, generate_drawing_rels_combined};
use crate::xml::{self, generate_drawing_xml_combined, generate_drawing_rels_combined};
use mtzip::{level::CompressionLevel, ZipArchive};
use std::fs::File;
use std::io::Write;
use std::collections::HashMap;
use arrow_array::RecordBatch;
use rayon::prelude::*;
// ============================================================================
// DICT API - Dict-based (backward compatibility)
// ============================================================================

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
        
        let drawing_rels = generate_drawing_rels_combined(config.charts.len(), &config.images);
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

    let config = StyleConfig::default();
    
    // Generate XMLs in parallel if num_threads > 1 and multiple sheets
    let xml_sheets: Vec<Vec<u8>> = if num_threads > 1 && sheets.len() > 1 {
        let pool = rayon::ThreadPoolBuilder::new()
            .num_threads(num_threads)
            .build()
            .map_err(|e| WriteError::Validation(format!("Thread pool error: {}", e)))?;
        
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
    

    add_static_files(&mut zipper, &sheet_names, Some(&registry), &vec![config.tables.len()], &charts_count, &images_data);
    
    let xml_data = xml::generate_sheet_xml_from_arrow(batches, &updated_config, &col_format_map, &cell_style_map)?;
    
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
            rels_xml.push_str(&format!("<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{}\" TargetMode=\"External\"/>\n", idx, url));
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
        for (idx, table) in config.tables.iter().enumerate() {
            let table_id = (idx + 1) as u32;
            
            // If table starts after row 1, we inserted a header row, so adjust end_row
            let mut adjusted_table = table.clone();
            if adjusted_table.range.0 > 1 {
                adjusted_table.range.2 += 1; // end_row++
            }
            
            let col_names = if table.column_names.is_empty() && !batches.is_empty() {
                let schema = batches[0].schema();
                let (_, start_col, _, end_col) = table.range;
                schema.fields()[start_col..=end_col]
                    .iter()
                    .map(|f| f.name().clone())
                    .collect()
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
        
        let drawing_rels = generate_drawing_rels_combined(config.charts.len(), &config.images);
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
    for (_, name, _) in sheets {
        validate_sheet_name(name)?;
    }

    let mut style_registry = StyleRegistry::new();
    let mut sheet_col_format_maps = Vec::new();
    let mut sheet_cell_style_maps = Vec::new();
    let mut sheet_dxf_mappings = Vec::new();

    for (batches, _, config) in sheets {
        let schema = batches[0].schema();
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

    let xml_and_hyperlinks: Vec<(Vec<u8>, Vec<(String, usize)>)> = 
        if num_threads > 1 && sheets.len() > 1 {
            let pool = rayon::ThreadPoolBuilder::new()
                .num_threads(num_threads)
                .build()
                .map_err(|e| WriteError::Validation(format!("Thread pool error: {}", e)))?;
            
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
                        let xml_data = xml::generate_sheet_xml_from_arrow(batches, &modified_config, col_format_map, cell_style_map)?;
                        let hyperlinks: Vec<(String, usize)> = modified_config.hyperlinks
                            .iter()
                            .enumerate()
                            .map(|(i, h)| (h.url.clone(), i + 1))
                            .collect();
                        Ok((xml_data, hyperlinks))
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
                    let xml_data = xml::generate_sheet_xml_from_arrow(batches, &modified_config, col_format_map, cell_style_map)?;
                    let hyperlinks: Vec<(String, usize)> = modified_config.hyperlinks
                        .iter()
                        .enumerate()
                        .map(|(i, h)| (h.url.clone(), i + 1))
                        .collect();
                    Ok((xml_data, hyperlinks))
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
    add_static_files(&mut zipper, &sheet_names, Some(&style_registry), &tables_per_sheet, &charts_per_sheet, &images_per_sheet);

    let mut global_chart_id = 1;
    let mut global_table_id = 1;
    let mut drawing_id = 1;

    for (idx, (xml_data, hyperlinks)) in xml_and_hyperlinks.into_iter().enumerate() {
        let sheet_config = &sheets[idx].2;
        
        zipper
            .add_file_from_memory(xml_data, format!("xl/worksheets/sheet{}.xml", idx + 1))
            .compression_level(CompressionLevel::fast())
            .done();

        let has_hyperlinks = !hyperlinks.is_empty();
        let has_tables = !sheet_config.tables.is_empty();
        let has_charts = !sheet_config.charts.is_empty();
        let has_images = !sheet_config.images.is_empty();

        if has_hyperlinks || has_tables || has_charts || has_images {
            let mut rels_xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
            
            for (url, rid) in &hyperlinks {
                rels_xml.push_str(&format!("<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{}\" TargetMode=\"External\"/>\n", rid, url));
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
            for table in &sheet_config.tables {
                let col_names = if table.column_names.is_empty() && !sheets[idx].0.is_empty() {
                    let schema = sheets[idx].0[0].schema();
                    let (_, start_col, _, end_col) = table.range;
                    schema.fields()[start_col..=end_col]
                        .iter()
                        .map(|f| f.name().clone())
                        .collect()
                } else {
                    table.column_names.clone()
                };
                
                let table_xml = xml::generate_table_xml(table, global_table_id as u32, &col_names);
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
            let sheet_start_chart_id = global_chart_id;
            
            let drawing_xml = generate_drawing_xml_combined(&sheet_config.charts, &sheet_config.images);
            zipper
                .add_file_from_memory(drawing_xml.into_bytes(), format!("xl/drawings/drawing{}.xml", drawing_id))
                .compression_level(CompressionLevel::fast())
                .done();
            
            let drawing_rels = generate_drawing_rels_combined(sheet_config.charts.len(), &sheet_config.images);
            
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
            // Add image files
            for (idx, image) in sheet_config.images.iter().enumerate() {
                zipper
                    .add_file_from_memory(
                        image.image_data.clone(),
                        format!("xl/media/image{}.{}", idx + 1, image.extension)
                    )
                    .compression_level(CompressionLevel::fast())
                    .done();
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
    let images_per_sheet: Vec<(&[ExcelImage], usize)> = images_data.iter()
            .map(|(imgs, count)| (imgs.as_slice(), *count))
            .collect();
        
        zipper
            .add_file_from_memory(
                xml::generate_content_types_with_charts(sheet_names, tables_count, charts_count, &images_per_sheet).into_bytes(),
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
            xml::generate_workbook_rels(sheet_names.len()).into_bytes(),
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
    let mut file = File::create(filename)?;
    zipper.write(&mut file)
        .map_err(|e| WriteError::Validation(e.to_string()))?;
    file.flush()?;
    file.sync_all()?;
    Ok(())
}

fn validate_sheet_name(name: &str) -> Result<(), WriteError> {
    if name.len() > 31 {
        return Err(WriteError::Validation(format!("Sheet name '{}' exceeds 31 chars", name)));
    }
    if name.chars().any(|c| "[]':*?/\\".contains(c)) {
        return Err(WriteError::Validation(format!("Sheet name '{}' contains invalid chars", name)));
    }
    Ok(())
}