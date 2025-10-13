use crate::types::{SheetData, WriteError};
use crate::styles::{StyleConfig, generate_styles_xml, generate_styles_xml_enhanced, StyleRegistry, ConditionalRule};
use crate::xml;
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
    
    add_static_files(&mut zipper, &sheet_names, None, &[0], &[0]);
    
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
    
    add_static_files(&mut zipper, &sheet_names, None, &[0], &charts_count);
    
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
        
        let drawing_rels = xml::generate_drawing_rels(config.charts.len());
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

    add_static_files(&mut zipper, &sheet_names, None, &vec![0; sheets.len()], &vec![0; sheets.len()]);

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
    // VALIDATE SHEET NAME
    crate::validation::validate_sheet_name(sheet_name)
        .map_err(WriteError::Validation)?;
    
    // CALCULATE DATA BOUNDS
    let schema = batches.first()
        .ok_or_else(|| WriteError::Validation("No data to write".to_string()))?
        .schema();
    let num_cols = schema.fields().len();
    let total_rows: usize = batches.iter().map(|b| b.num_rows()).sum();
    
    // PRE-VALIDATE ALL CONFIG
    let validation_result = crate::validation::validate_style_config(
        config, 
        total_rows + 1, // +1 for header
        num_cols
    );
    validation_result.to_error()?;
    
    // Build style registry with FIXED DXF IDs
    let (final_registry, updated_config) = if !config.cell_styles.is_empty() 
        || !config.conditional_formats.is_empty() {
        
        let configs = vec![config];
        let (mut registry, dxf_mappings) = crate::validation::build_global_dxf_mappings(&configs);
        
        let mut modified_config = config.clone();
        if let Some(mapping) = dxf_mappings.first() {
            modified_config.cond_format_dxf_ids = mapping.clone();
        }
        
        (Some(registry), modified_config)
    } else {
        (None, config.clone())
    };

    let mut zipper = ZipArchive::new();
    let sheet_names = vec![sheet_name];
    let charts_count = vec![config.charts.len()];
    
    add_static_files(&mut zipper, &sheet_names, final_registry.as_ref(), &vec![config.tables.len()], &charts_count);
    
    let xml_data = xml::generate_sheet_xml_from_arrow(batches, &updated_config)?;
    zipper
        .add_file_from_memory(xml_data, "xl/worksheets/sheet1.xml".to_string())
        .compression_level(CompressionLevel::fast())
        .done();

    // Add worksheet relationships if hyperlinks OR tables OR charts exist
    let hyperlinks_with_idx: Vec<(String, usize)> = config.hyperlinks
        .iter()
        .enumerate()
        .map(|(idx, h)| (h.url.clone(), idx + 1))
        .collect();
    
    let has_any_rels = !config.hyperlinks.is_empty() || !config.tables.is_empty() || !config.charts.is_empty();
    
    if has_any_rels {
        let mut rels_xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
        
        // Add hyperlinks
        for (url, idx) in &hyperlinks_with_idx {
            rels_xml.push_str(&format!("<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{}\" TargetMode=\"External\"/>\n", idx, url));
        }
        
        // Add tables
        for idx in 0..config.tables.len() {
            rels_xml.push_str(&format!("<Relationship Id=\"rIdTable{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table{}.xml\"/>\n", idx + 1, idx + 1));
        }
        
        // Add drawing (for charts)
        if !config.charts.is_empty() {
            rels_xml.push_str("<Relationship Id=\"rIdDraw1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing1.xml\"/>\n");
        }
        
        rels_xml.push_str("</Relationships>");
        
        zipper
            .add_file_from_memory(rels_xml.into_bytes(), "xl/worksheets/_rels/sheet1.xml.rels".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
    }
    
    // ADD TABLE FILES
    if !config.tables.is_empty() {
        for (idx, table) in config.tables.iter().enumerate() {
            let table_id = (idx + 1) as u32;
            
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
            
            let table_xml = xml::generate_table_xml(table, table_id, &col_names);
            zipper
                .add_file_from_memory(
                    table_xml.into_bytes(),
                    format!("xl/tables/table{}.xml", table_id)
                )
                .compression_level(CompressionLevel::fast())
                .done();
        }
    }
    
    // ADD CHART FILES
    if !config.charts.is_empty() {
        // Add drawing XML
        let drawing_xml = xml::generate_drawing_xml(&config.charts);
        zipper
            .add_file_from_memory(drawing_xml.into_bytes(), "xl/drawings/drawing1.xml".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
        
        // Add drawing relationships
        let drawing_rels = xml::generate_drawing_rels(config.charts.len());
        zipper
            .add_file_from_memory(drawing_rels.into_bytes(), "xl/drawings/_rels/drawing1.xml.rels".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
        
        // Add chart files
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
    // VALIDATE ALL SHEET NAMES
    let sheet_names: Vec<&str> = sheets.iter().map(|(_, name, _)| *name).collect();
    let name_validation = crate::validation::validate_sheet_names(&sheet_names);
    name_validation.to_error()?;
    
    // VALIDATE EACH SHEET'S CONFIG
    for (idx, (batches, name, config)) in sheets.iter().enumerate() {
        let num_cols = batches.first()
            .ok_or_else(|| WriteError::Validation(format!("Sheet '{}': No data", name)))?
            .schema()
            .fields()
            .len();
        let total_rows: usize = batches.iter().map(|b| b.num_rows()).sum();
        
        let validation_result = crate::validation::validate_style_config(
            config,
            total_rows + 1,
            num_cols
        );
        
        if !validation_result.is_valid() {
            return Err(WriteError::Validation(format!(
                "Sheet '{}' validation failed:\n{}",
                name,
                validation_result.errors.join("\n")
            )));
        }
    }
    
    // BUILD GLOBAL DXF MAPPINGS - FIXES #2 (DXF ID COLLISIONS)
    let configs: Vec<&StyleConfig> = sheets.iter().map(|(_, _, c)| c).collect();
    let (style_registry, sheet_dxf_mappings) = 
        crate::validation::build_global_dxf_mappings(&configs);
    // Build combined style registry for all sheets
    let (style_registry, sheet_dxf_mappings) = {
        let mut registry = StyleRegistry::new();
        let mut has_custom_styles = false;
        let mut all_dxf_mappings = Vec::new();
        
        for (_, _, config) in sheets {
            let mut dxf_ids = HashMap::new();
            
            if !config.cell_styles.is_empty() || !config.conditional_formats.is_empty() {
                has_custom_styles = true;
                
                for cell_style in &config.cell_styles {
                    registry.register_cell_style(&cell_style.style);
                }
                
                for (idx, cond_format) in config.conditional_formats.iter().enumerate() {
                    match &cond_format.rule {
                        ConditionalRule::CellValue { .. } | ConditionalRule::Top10 { .. } => {
                            registry.register_cell_style(&cond_format.style);
                            let dxf_id = registry.register_dxf(&cond_format.style);
                            dxf_ids.insert(idx, dxf_id);
                        }
                        _ => {}
                    }
                }
            }
            
            all_dxf_mappings.push(dxf_ids);
        }
        
        if has_custom_styles {
            (Some(registry), all_dxf_mappings)
        } else {
            (None, Vec::new())
        }
    };

    // Generate XMLs in parallel if num_threads > 1 and multiple sheets
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
                        
                        let xml_data = xml::generate_sheet_xml_from_arrow(batches, &modified_config)?;
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
            // Sequential fallback
            sheets
                .iter()
                .enumerate()
                .map(|(sheet_idx, (batches, _, config))| {
                    let mut modified_config = (*config).clone();
                    if sheet_idx < sheet_dxf_mappings.len() {
                        modified_config.cond_format_dxf_ids = sheet_dxf_mappings[sheet_idx].clone();
                    }
                    
                    let xml_data = xml::generate_sheet_xml_from_arrow(batches, &modified_config)?;
                    let hyperlinks: Vec<(String, usize)> = modified_config.hyperlinks
                        .iter()
                        .enumerate()
                        .map(|(i, h)| (h.url.clone(), i + 1))
                        .collect();
                    Ok((xml_data, hyperlinks))
                })
                .collect::<Result<Vec<_>, WriteError>>()?
        };

    // Build ZIP sequentially (not thread-safe)
    let mut zipper = ZipArchive::new();
    let sheet_names: Vec<&str> = sheets.iter().map(|(_, name, _)| *name).collect();
    let tables_per_sheet: Vec<usize> = sheets.iter().map(|(_, _, cfg)| cfg.tables.len()).collect();
    let charts_per_sheet: Vec<usize> = sheets.iter().map(|(_, _, cfg)| cfg.charts.len()).collect();

    add_static_files(&mut zipper, &sheet_names, style_registry.as_ref(), &tables_per_sheet, &charts_per_sheet);

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
        
        if has_hyperlinks || has_tables || has_charts {
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
            
            if has_charts {
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
        
        if has_charts {
            let sheet_start_chart_id = global_chart_id;
            
            let drawing_xml = xml::generate_drawing_xml(&sheet_config.charts);
            zipper
                .add_file_from_memory(drawing_xml.into_bytes(), format!("xl/drawings/drawing{}.xml", drawing_id))
                .compression_level(CompressionLevel::fast())
                .done();
            
            let mut drawing_rels = String::with_capacity(300 + sheet_config.charts.len() * 150);
            drawing_rels.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
            for i in 0..sheet_config.charts.len() {
                drawing_rels.push_str(&format!("<Relationship Id=\"rIdChart{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart{}.xml\"/>\n", 
                    i + 1, 
                    sheet_start_chart_id + i));
            }
            drawing_rels.push_str("</Relationships>");
            
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
) {
    zipper
        .add_file_from_memory(
            xml::generate_content_types_with_charts(sheet_names, tables_count, charts_count).into_bytes(),
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
    use crate::validation::write_file_atomic;
    
    write_file_atomic(filename, |file| {
        zipper.write(file)
            .map_err(|e| WriteError::Validation(e.to_string()))?;
        Ok(())
    })
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