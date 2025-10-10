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
    
    add_static_files(&mut zipper, &sheet_names, None, &[0]);
    
    let config = StyleConfig::default();
    let xml_data = xml::generate_sheet_xml_from_dict(sheet, &config)?;
    zipper
        .add_file_from_memory(xml_data, "xl/worksheets/sheet1.xml".to_string())
        .compression_level(CompressionLevel::fast())
        .done();

    
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

    add_static_files(&mut zipper, &sheet_names, None, &vec![0; sheets.len()]);

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

    // Build style registry and update config with proper DXF IDs
    let (final_registry, updated_config) = if !config.cell_styles.is_empty() || !config.conditional_formats.is_empty() {
        let mut registry = StyleRegistry::new();
        
        for cell_style in &config.cell_styles {
            registry.register_cell_style(&cell_style.style);
        }
        
        // Build proper DXF IDs for conditional formats
        let mut dxf_ids = HashMap::new();
        for (idx, cond_format) in config.conditional_formats.iter().enumerate() {
            // Only register dxf for rules that need it (cellIs, top10)
            match &cond_format.rule {
                ConditionalRule::CellValue { .. } | ConditionalRule::Top10 { .. } => {
                    registry.register_cell_style(&cond_format.style);
                    let dxf_id = registry.register_dxf(&cond_format.style);
                    dxf_ids.insert(idx, dxf_id);
                }
                // colorScale, dataBar don't use dxfId
                _ => {}
            }
        }
        
        let mut modified_config = config.clone();
        modified_config.cond_format_dxf_ids = dxf_ids;
        
        (Some(registry), modified_config)
    } else {
        (None, config.clone())
    };

let mut zipper = ZipArchive::new();
    let sheet_names = vec![sheet_name];
    
    add_static_files(&mut zipper, &sheet_names, final_registry.as_ref(), &vec![config.tables.len()]);

    // // DEBUG styles
    // if let Some(ref reg) = final_registry {
    //     let styles_xml = generate_styles_xml_enhanced(reg);
    //     std::fs::write("debug_styles.xml", styles_xml.as_bytes()).ok();
    // }
    
    let xml_data = xml::generate_sheet_xml_from_arrow(batches, &updated_config)?;
    // // DEBUG
    // std::fs::write("debug_sheet.xml", &xml_data).ok();
    zipper
        .add_file_from_memory(xml_data, "xl/worksheets/sheet1.xml".to_string())
        .compression_level(CompressionLevel::fast())
        .done();

    // Add worksheet relationships if hyperlinks OR tables exist
    let hyperlinks_with_idx: Vec<(String, usize)> = config.hyperlinks
        .iter()
        .enumerate()
        .map(|(idx, h)| (h.url.clone(), idx + 1))
        .collect();
    
    if !config.hyperlinks.is_empty() && config.tables.is_empty() {
        if let Some(rels_xml) = xml::generate_worksheet_rels(&hyperlinks_with_idx) {
            zipper
                .add_file_from_memory(rels_xml.into_bytes(), "xl/worksheets/_rels/sheet1.xml.rels".to_string())
                .compression_level(CompressionLevel::fast())
                .done();
        }
    }

        // ADD TABLE FILES
    if !config.tables.is_empty() {
        // Generate table relationship entries
        let mut table_rels = Vec::new();
        for idx in 0..config.tables.len() {
            table_rels.push((
                format!("rIdTable{}", idx + 1),
                format!("../tables/table{}.xml", idx + 1),
            ));
        }
        
        // Create worksheet rels with both hyperlinks and tables
        let rels_xml = xml::generate_worksheet_rels_with_tables(&hyperlinks_with_idx, &table_rels);
        zipper
            .add_file_from_memory(rels_xml.into_bytes(), "xl/worksheets/_rels/sheet1.xml.rels".to_string())
            .compression_level(CompressionLevel::fast())
            .done();
        
        // Add table XML files
        for (idx, table) in config.tables.iter().enumerate() {
            let table_id = (idx + 1) as u32;
            
            // Extract column names from batches
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
    } else if !config.hyperlinks.is_empty() {
        // Only hyperlinks, no tables
        if let Some(rels_xml) = xml::generate_worksheet_rels(&hyperlinks_with_idx) {
            zipper
                .add_file_from_memory(rels_xml.into_bytes(), "xl/worksheets/_rels/sheet1.xml.rels".to_string())
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
    // Validate sheet names
    for (_, name, _) in sheets {
        validate_sheet_name(name)?;
    }

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

    add_static_files(&mut zipper, &sheet_names, style_registry.as_ref(), &tables_per_sheet);

    for (idx, (xml_data, hyperlinks)) in xml_and_hyperlinks.into_iter().enumerate() {
        zipper
            .add_file_from_memory(xml_data, format!("xl/worksheets/sheet{}.xml", idx + 1))
            .compression_level(CompressionLevel::fast())
            .done();

        // Add worksheet relationships if hyperlinks exist
        if !hyperlinks.is_empty() {
            if let Some(rels_xml) = xml::generate_worksheet_rels(&hyperlinks) {
                zipper
                    .add_file_from_memory(
                        rels_xml.into_bytes(),
                        format!("xl/worksheets/_rels/sheet{}.xml.rels", idx + 1)
                    )
                    .compression_level(CompressionLevel::fast())
                    .done();
            }
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
) {
    zipper
        .add_file_from_memory(
            xml::generate_content_types(sheet_names, tables_count).into_bytes(),
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