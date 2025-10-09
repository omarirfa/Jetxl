use crate::types::{SheetData, WriteError};
use crate::styles::{StyleConfig, generate_styles_xml, generate_styles_xml_enhanced, StyleRegistry};
use crate::xml;
use mtzip::{level::CompressionLevel, ZipArchive};
use std::fs::File;
use std::io::Write;
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
    
    add_static_files(&mut zipper, &sheet_names, None);
    
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

    add_static_files(&mut zipper, &sheet_names, None);

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

    // Build style registry if we have custom cell styles
    let style_registry = if !config.cell_styles.is_empty() || !config.conditional_formats.is_empty() {
        let mut registry = StyleRegistry::new();
        
        for cell_style in &config.cell_styles {
            registry.register_cell_style(&cell_style.style);
        }
        
        for cond_format in &config.conditional_formats {
            registry.register_cell_style(&cond_format.style);
            registry.register_dxf(&cond_format.style);  // Register dxf too
        }
        
        Some(registry)
    } else {
        None
    };

    let mut zipper = ZipArchive::new();
    let sheet_names = vec![sheet_name];
    
    add_static_files(&mut zipper, &sheet_names, style_registry.as_ref());
    
    let xml_data = xml::generate_sheet_xml_from_arrow(batches, config)?;
    zipper
        .add_file_from_memory(xml_data, "xl/worksheets/sheet1.xml".to_string())
        .compression_level(CompressionLevel::fast())
        .done();

    // Add worksheet relationships if hyperlinks exist
    if !config.hyperlinks.is_empty() {
        let hyperlinks_with_idx: Vec<(String, usize)> = config.hyperlinks
            .iter()
            .enumerate()
            .map(|(idx, h)| (h.url.clone(), idx + 1))
            .collect();
        
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
    let style_registry = {
        let mut registry = StyleRegistry::new();
        let mut has_custom_styles = false;
        
        for (_, _, config) in sheets {
            if !config.cell_styles.is_empty() || !config.conditional_formats.is_empty() {
                has_custom_styles = true;
                
                for cell_style in &config.cell_styles {
                    registry.register_cell_style(&cell_style.style);
                }
                
                for cond_format in &config.conditional_formats {
                    registry.register_cell_style(&cond_format.style);
                    registry.register_dxf(&cond_format.style);  // Register dxf too
                }
            }
        }
        
        if has_custom_styles {
            Some(registry)
        } else {
            None
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
                    .map(|(batches, _, config)| {
                        let xml_data = xml::generate_sheet_xml_from_arrow(batches, config)?;
                        let hyperlinks: Vec<(String, usize)> = config.hyperlinks
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
                .map(|(batches, _, config)| {
                    let xml_data = xml::generate_sheet_xml_from_arrow(batches, config)?;
                    let hyperlinks: Vec<(String, usize)> = config.hyperlinks
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

    add_static_files(&mut zipper, &sheet_names, style_registry.as_ref());

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
) {
    zipper
        .add_file_from_memory(
            xml::generate_content_types(sheet_names).into_bytes(),
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