mod types;
mod writer;
mod xml;
mod styles;

use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};
use pyo3_arrow::input::AnyRecordBatch;
use arrow_array::RecordBatch;
use types::{CellValue, SheetData};
use styles::*;
use std::collections::HashMap;

// ============================================================================
// LEGACY API - Dict-based (backward compatibility)
// ============================================================================

#[pyfunction]
#[pyo3(signature = (columns, filename, sheet_name = None))]
fn write_sheet(
    py: Python,
    columns: Bound<PyDict>,
    filename: String,
    sheet_name: Option<String>,
) -> PyResult<()> {
    let sheet = extract_sheet_data(py, &columns, sheet_name)?;

    py.detach(|| {
        writer::write_single_sheet(&sheet, &filename)
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyIOError, _>(e.to_string()))
    })
}

#[pyfunction]
#[pyo3(signature = (sheets_data, filename, num_threads))]
fn write_sheets(
    py: Python,
    sheets_data: Vec<Bound<PyDict>>,
    filename: String,
    num_threads: usize,  
) -> PyResult<()> {
    let sheets: Result<Vec<_>, _> = sheets_data
        .into_iter()
        .enumerate()
        .map(|(i, sheet_dict)| {
            let name = sheet_dict
                .get_item("name")?
                .and_then(|n| n.extract::<String>().ok())
                .unwrap_or_else(|| format!("Sheet{}", i + 1));

            let cols_item = sheet_dict
                .get_item("columns")?
                .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyKeyError, _>("Missing 'columns' key"))?;
            let cols = cols_item.downcast::<PyDict>()?;

            extract_sheet_data(py, cols, Some(name))
        })
        .collect();

    let sheets = sheets?;


    py.detach(|| {
        writer::write_multiple_sheets(&sheets, &filename, num_threads)
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyIOError, _>(e.to_string()))
    })
}

// ============================================================================
// ARROW API - Direct Arrow → XML (Zero-Copy) - ENHANCED
// ============================================================================

#[pyfunction]
#[pyo3(signature = (
    arrow_data,
    filename,
    sheet_name = None,
    auto_filter = false,
    freeze_rows = 0,
    freeze_cols = 0,
    auto_width = false,
    styled_headers = false,
    column_widths = None,
    column_formats = None,
    merge_cells = None,
    data_validations = None,
    hyperlinks = None,
    row_heights = None,
    cell_styles = None,
    formulas = None,
    conditional_formats = None,
))]
#[allow(clippy::too_many_arguments)]
fn write_sheet_arrow(
    py: Python,
    arrow_data: &Bound<PyAny>,
    filename: String,
    sheet_name: Option<String>,
    auto_filter: bool,
    freeze_rows: usize,
    freeze_cols: usize,
    auto_width: bool,
    styled_headers: bool,
    column_widths: Option<HashMap<String, f64>>,
    column_formats: Option<HashMap<String, String>>,
    merge_cells: Option<Vec<(usize, usize, usize, usize)>>,
    data_validations: Option<Vec<Bound<PyDict>>>,
    hyperlinks: Option<Vec<(usize, usize, String, Option<String>)>>,
    row_heights: Option<HashMap<usize, f64>>,
    cell_styles: Option<Vec<Bound<PyDict>>>,
    formulas: Option<Vec<(usize, usize, String, Option<String>)>>,
    conditional_formats: Option<Vec<Bound<PyDict>>>,
) -> PyResult<()> {
    let any_batch = AnyRecordBatch::extract_bound(arrow_data)?;
    let reader = any_batch.into_reader()?;
    
    let batches: Vec<RecordBatch> = reader
        .collect::<Result<Vec<_>, _>>()
        .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(
            format!("Failed to read Arrow data: {}", e)
        ))?;
    
    if batches.is_empty() {
        return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
            "Arrow data is empty"
        ));
    }

    let name = sheet_name.unwrap_or_else(|| "Sheet1".to_string());

    // Build config
    let mut config = StyleConfig {
        auto_filter,
        freeze_rows,
        freeze_cols,
        styled_headers,
        column_widths,
        auto_width,
        column_formats: column_formats.map(|cf| {
            cf.into_iter()
                .filter_map(|(k, v)| parse_number_format(&v).map(|fmt| (k, fmt)))
                .collect()
        }),
        merge_cells: merge_cells.unwrap_or_default().into_iter().map(|(sr, sc, er, ec)| {
            MergeRange { start_row: sr, start_col: sc, end_row: er, end_col: ec }
        }).collect(),
        data_validations: Vec::new(),
        hyperlinks: hyperlinks.unwrap_or_default().into_iter().map(|(row, col, url, display)| {
            Hyperlink { row, col, url, display }
        }).collect(),
        row_heights,
        cell_styles: Vec::new(),
        formulas: Vec::new(),
        conditional_formats: Vec::new(),
    };

    // Parse data validations
    if let Some(validations) = data_validations {
        for val_dict in validations {
            if let Ok(validation) = extract_data_validation(&val_dict) {
                config.data_validations.push(validation);
            }
        }
    }

    // Parse cell styles
    if let Some(styles) = cell_styles {
        for style_dict in styles {
            if let Ok(cell_style) = extract_cell_style(&style_dict) {
                config.cell_styles.push(cell_style);
            }
        }
    }

    // Parse formulas
    if let Some(formulas_vec) = formulas {
        for (row, col, formula, cached_value) in formulas_vec {
            config.formulas.push(Formula { row, col, formula, cached_value });
        }
    }

    // Parse conditional formats
    if let Some(cond_formats) = conditional_formats {
        for cf_dict in cond_formats {
            if let Ok(cond_format) = extract_conditional_format(&cf_dict) {
                config.conditional_formats.push(cond_format);
            }
        }
    }

    py.detach(|| {
        writer::write_single_sheet_arrow_with_config(&batches, &name, &filename, &config)
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyIOError, _>(e.to_string()))
    })
}

#[pyfunction]
#[pyo3(signature = (arrow_sheets, filename, num_threads))]
fn write_sheets_arrow(
    py: Python,
    arrow_sheets: Vec<(Bound<PyAny>, String)>, 
    filename: String,
    num_threads: usize,
) -> PyResult<()> {
    let sheets: Result<Vec<_>, _> = arrow_sheets
        .into_iter()
        .map(|(arrow_data, name)| {
            let any_batch = AnyRecordBatch::extract_bound(&arrow_data)?;
            let reader = any_batch.into_reader()?;
            
            let batches: Vec<RecordBatch> = reader
                .collect::<Result<Vec<_>, _>>()
                .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(
                    format!("Failed to read Arrow data: {}", e)
                ))?;
            
            if batches.is_empty() {
                return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                    "Arrow data is empty"
                ));
            }
            
            Ok((batches, name))
        })
        .collect();

    let sheets = sheets?;


    py.detach(|| {
        writer::write_multiple_sheets_arrow(&sheets, &filename, num_threads)
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyIOError, _>(e.to_string()))
    })
}

// ============================================================================
// Helper functions - Extraction from Python
// ============================================================================

fn extract_sheet_data(
    py: Python,
    columns: &Bound<PyDict>,
    sheet_name: Option<String>,
) -> PyResult<SheetData> {
    let mut cols = Vec::with_capacity(columns.len());

    for (key, value) in columns.iter() {
        let col_name = key.extract::<String>()?;
        let col_data = extract_column(py, &value)?;
        cols.push((col_name, col_data));
    }

    Ok(SheetData {
        name: sheet_name.unwrap_or_else(|| "Sheet1".to_string()),
        columns: cols,
    })
}

fn extract_column(py: Python, value: &Bound<PyAny>) -> PyResult<Vec<CellValue>> {
    if let Ok(list) = value.downcast::<PyList>() {
        let len = list.len();
        let mut result = Vec::with_capacity(len);

        for item in list.iter() {
            result.push(CellValue::from_py(py, &item)?);
        }

        Ok(result)
    } else {
        Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(
            "Column must be a list",
        ))
    }
}

fn parse_number_format(s: &str) -> Option<NumberFormat> {
    match s.to_lowercase().as_str() {
        "general" => Some(NumberFormat::General),
        "integer" => Some(NumberFormat::Integer),
        "decimal2" => Some(NumberFormat::Decimal2),
        "decimal4" => Some(NumberFormat::Decimal4),
        "percentage" => Some(NumberFormat::Percentage),
        "percentage_decimal" => Some(NumberFormat::PercentageDecimal),
        "currency" => Some(NumberFormat::Currency),
        "currency_rounded" => Some(NumberFormat::CurrencyRounded),
        "date" => Some(NumberFormat::Date),
        "datetime" => Some(NumberFormat::DateTime),
        "time" => Some(NumberFormat::Time),
        _ => None,
    }
}

fn extract_data_validation(dict: &Bound<PyDict>) -> PyResult<DataValidation> {
    let start_row: usize = dict.get_item("start_row")?.unwrap().extract()?;
    let start_col: usize = dict.get_item("start_col")?.unwrap().extract()?;
    let end_row: usize = dict.get_item("end_row")?.unwrap().extract()?;
    let end_col: usize = dict.get_item("end_col")?.unwrap().extract()?;
    let val_type: String = dict.get_item("type")?.unwrap().extract()?;
    
    let validation_type = match val_type.as_str() {
        "list" => {
            let items: Vec<String> = dict.get_item("items")?.unwrap().extract()?;
            ValidationType::List(items)
        }
        "whole_number" => {
            let min: i64 = dict.get_item("min")?.unwrap().extract()?;
            let max: i64 = dict.get_item("max")?.unwrap().extract()?;
            ValidationType::WholeNumber { min, max }
        }
        "decimal" => {
            let min: f64 = dict.get_item("min")?.unwrap().extract()?;
            let max: f64 = dict.get_item("max")?.unwrap().extract()?;
            ValidationType::Decimal { min, max }
        }
        "text_length" => {
            let min: usize = dict.get_item("min")?.unwrap().extract()?;
            let max: usize = dict.get_item("max")?.unwrap().extract()?;
            ValidationType::TextLength { min, max }
        }
        _ => return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>("Invalid validation type")),
    };
    
    let show_dropdown = dict.get_item("show_dropdown")?.map(|v| v.extract()).unwrap_or(Ok(true))?;
    let error_title = dict.get_item("error_title")?.and_then(|v| v.extract().ok());
    let error_message = dict.get_item("error_message")?.and_then(|v| v.extract().ok());
    
    Ok(DataValidation {
        start_row,
        start_col,
        end_row,
        end_col,
        validation_type,
        error_title,
        error_message,
        show_dropdown,
    })
}

fn extract_cell_style(dict: &Bound<PyDict>) -> PyResult<CellStyleMap> {
    let row: usize = dict.get_item("row")?.unwrap().extract()?;
    let col: usize = dict.get_item("col")?.unwrap().extract()?;
    
    let mut cell_style = CellStyle {
        font: None,
        fill: None,
        border: None,
        alignment: None,
        number_format: None,
    };
    
    // Extract font
    if let Some(font_dict) = dict.get_item("font")? {
        let font_dict = font_dict.downcast::<PyDict>()?;
        cell_style.font = Some(FontStyle {
            bold: font_dict.get_item("bold")?.map(|v| v.extract()).unwrap_or(Ok(false))?,
            italic: font_dict.get_item("italic")?.map(|v| v.extract()).unwrap_or(Ok(false))?,
            underline: font_dict.get_item("underline")?.map(|v| v.extract()).unwrap_or(Ok(false))?,
            size: font_dict.get_item("size")?.and_then(|v| v.extract().ok()),
            color: font_dict.get_item("color")?.and_then(|v| v.extract().ok()),
            name: font_dict.get_item("name")?.and_then(|v| v.extract().ok()),
        });
    }
    
    // Extract fill
    if let Some(fill_dict) = dict.get_item("fill")? {
        let fill_dict = fill_dict.downcast::<PyDict>()?;
        let pattern: String = fill_dict.get_item("pattern")?.map(|v| v.extract()).unwrap_or(Ok("none".to_string()))?;
        cell_style.fill = Some(FillStyle {
            pattern_type: match pattern.as_str() {
                "solid" => PatternType::Solid,
                "gray125" => PatternType::Gray125,
                _ => PatternType::None,
            },
            fg_color: fill_dict.get_item("fg_color")?.and_then(|v| v.extract().ok()),
            bg_color: fill_dict.get_item("bg_color")?.and_then(|v| v.extract().ok()),
        });
    }
    
    // Extract border
    if let Some(border_dict) = dict.get_item("border")? {
        let border_dict = border_dict.downcast::<PyDict>()?;
        
        let parse_side = |side_dict: &Bound<PyDict>| -> PyResult<BorderSide> {
            let style: String = side_dict.get_item("style")?.unwrap().extract()?;
            Ok(BorderSide {
                style: match style.as_str() {
                    "medium" => BorderLineStyle::Medium,
                    "thick" => BorderLineStyle::Thick,
                    "double" => BorderLineStyle::Double,
                    "dotted" => BorderLineStyle::Dotted,
                    "dashed" => BorderLineStyle::Dashed,
                    _ => BorderLineStyle::Thin,
                },
                color: side_dict.get_item("color")?.and_then(|v| v.extract().ok()),
            })
        };
        
        // Parse each side separately to avoid borrow checker issues
        let left = if let Some(side) = border_dict.get_item("left")? {
            if let Ok(side_dict) = side.downcast::<PyDict>() {
                parse_side(side_dict).ok()
            } else {
                None
            }
        } else {
            None
        };
        
        let right = if let Some(side) = border_dict.get_item("right")? {
            if let Ok(side_dict) = side.downcast::<PyDict>() {
                parse_side(side_dict).ok()
            } else {
                None
            }
        } else {
            None
        };
        
        let top = if let Some(side) = border_dict.get_item("top")? {
            if let Ok(side_dict) = side.downcast::<PyDict>() {
                parse_side(side_dict).ok()
            } else {
                None
            }
        } else {
            None
        };
        
        let bottom = if let Some(side) = border_dict.get_item("bottom")? {
            if let Ok(side_dict) = side.downcast::<PyDict>() {
                parse_side(side_dict).ok()
            } else {
                None
            }
        } else {
            None
        };
        
        cell_style.border = Some(BorderStyle {
            left,
            right,
            top,
            bottom,
        });
    }
    
    // Extract alignment
    if let Some(align_dict) = dict.get_item("alignment")? {
        let align_dict = align_dict.downcast::<PyDict>()?;
        
        let horizontal = align_dict.get_item("horizontal")?.and_then(|v| {
            let s: String = v.extract().ok()?;
            match s.as_str() {
                "center" => Some(HorizontalAlignment::Center),
                "right" => Some(HorizontalAlignment::Right),
                "justify" => Some(HorizontalAlignment::Justify),
                "left" => Some(HorizontalAlignment::Left),
                _ => None,
            }
        });
        
        let vertical = align_dict.get_item("vertical")?.and_then(|v| {
            let s: String = v.extract().ok()?;
            match s.as_str() {
                "center" => Some(VerticalAlignment::Center),
                "bottom" => Some(VerticalAlignment::Bottom),
                "top" => Some(VerticalAlignment::Top),
                _ => None,
            }
        });
        
        cell_style.alignment = Some(AlignmentStyle {
            horizontal,
            vertical,
            wrap_text: align_dict.get_item("wrap_text")?.map(|v| v.extract()).unwrap_or(Ok(false))?,
            text_rotation: align_dict.get_item("text_rotation")?.and_then(|v| v.extract().ok()),
        });
    }
    
    // Extract number format
    if let Some(fmt_str) = dict.get_item("number_format")? {
        let fmt_str: String = fmt_str.extract()?;
        cell_style.number_format = parse_number_format(&fmt_str);
    }
    
    Ok(CellStyleMap { row, col, style: cell_style })
}

fn extract_conditional_format(dict: &Bound<PyDict>) -> PyResult<ConditionalFormat> {
    let start_row: usize = dict.get_item("start_row")?.unwrap().extract()?;
    let start_col: usize = dict.get_item("start_col")?.unwrap().extract()?;
    let end_row: usize = dict.get_item("end_row")?.unwrap().extract()?;
    let end_col: usize = dict.get_item("end_col")?.unwrap().extract()?;
    let rule_type: String = dict.get_item("rule_type")?.unwrap().extract()?;
    let priority: u32 = dict.get_item("priority")?.map(|v| v.extract()).unwrap_or(Ok(1))?;
    
    let rule = match rule_type.as_str() {
        "cell_value" => {
            let operator: String = dict.get_item("operator")?.unwrap().extract()?;
            let value: String = dict.get_item("value")?.unwrap().extract()?;
            
            let op = match operator.as_str() {
                "greater_than" => ComparisonOperator::GreaterThan,
                "less_than" => ComparisonOperator::LessThan,
                "equal" => ComparisonOperator::Equal,
                "not_equal" => ComparisonOperator::NotEqual,
                "greater_than_or_equal" => ComparisonOperator::GreaterThanOrEqual,
                "less_than_or_equal" => ComparisonOperator::LessThanOrEqual,
                "between" => ComparisonOperator::Between,
                _ => ComparisonOperator::GreaterThan,
            };
            
            ConditionalRule::CellValue { operator: op, value }
        }
        "color_scale" => {
            let min_color: String = dict.get_item("min_color")?.unwrap().extract()?;
            let max_color: String = dict.get_item("max_color")?.unwrap().extract()?;
            let mid_color: Option<String> = dict.get_item("mid_color")?.and_then(|v| v.extract().ok());
            
            ConditionalRule::ColorScale { min_color, max_color, mid_color }
        }
        "data_bar" => {
            let color: String = dict.get_item("color")?.unwrap().extract()?;
            let show_value: bool = dict.get_item("show_value")?.map(|v| v.extract()).unwrap_or(Ok(true))?;
            
            ConditionalRule::DataBar { color, show_value }
        }
        "top10" => {
            let rank: u32 = dict.get_item("rank")?.unwrap().extract()?;
            let bottom: bool = dict.get_item("bottom")?.map(|v| v.extract()).unwrap_or(Ok(false))?;
            
            ConditionalRule::Top10 { rank, bottom }
        }
        _ => return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>("Invalid rule type")),
    };
    
    // Extract style from dict if provided, otherwise use default
    let style = if let Some(style_dict) = dict.get_item("style")? {
        extract_cell_style(style_dict.downcast::<PyDict>()?)?.style
    } else {
        // Default red bold style
        CellStyle {
            font: Some(FontStyle {
                bold: true,
                italic: false,
                underline: false,
                size: None,
                color: Some("FFFF0000".to_string()),
                name: None,
            }),
            fill: None,
            border: None,
            alignment: None,
            number_format: None,
        }
    };
    
    Ok(ConditionalFormat {
        start_row,
        start_col,
        end_row,
        end_col,
        rule,
        style,
        priority,
    })
}

#[pymodule]
fn jetxl(m: &Bound<'_, PyModule>) -> PyResult<()> {
    // Legacy dict-based API
    m.add_function(wrap_pyfunction!(write_sheet, m)?)?;
    m.add_function(wrap_pyfunction!(write_sheets, m)?)?;
    
    // Arrow fast path API
    m.add_function(wrap_pyfunction!(write_sheet_arrow, m)?)?;
    m.add_function(wrap_pyfunction!(write_sheets_arrow, m)?)?;
    
    Ok(())
}