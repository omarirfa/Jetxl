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
#[pyo3(signature = (columns, filename, sheet_name = None, charts = None))]
/// Write dict-based data to Excel (legacy API).
///
/// Args:
///     columns (dict): Dictionary of column_name -> list of values
///     filename (str): Output file path
///     sheet_name (str, optional): Sheet name
fn write_sheet(
    py: Python,
    columns: Bound<PyDict>,
    filename: String,
    sheet_name: Option<String>,
    charts: Option<Vec<Bound<PyDict>>>,
) -> PyResult<()> {
    let sheet = extract_sheet_data(py, &columns, sheet_name)?;

    let mut config = StyleConfig::default();
    if let Some(charts_vec) = charts {
        for chart_dict in charts_vec {
            if let Ok(chart) = extract_chart(&chart_dict) {
                config.charts.push(chart);
            }
        }
    }

    py.detach(|| {
        writer::write_single_sheet_with_config(&sheet, &filename, &config)
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
    write_header_row = true,
    column_widths = None,
    column_formats = None,
    merge_cells = None,
    data_validations = None,
    hyperlinks = None,
    row_heights = None,
    cell_styles = None,
    formulas = None,
    conditional_formats = None,
    tables = None, 
    charts = None,
    images = None,
    gridlines_visible = true,
    zoom_scale = None,
    tab_color = None,
    default_row_height = None,
    hidden_columns = None,
    hidden_rows = None,
    right_to_left = false,
    data_start_row = 0,
))]
/// Write Arrow data to an Excel file with advanced formatting options.
/// 
/// Args:
///     arrow_data: PyArrow Table or RecordBatch
///     filename (str): Output file path
///     sheet_name (str, optional): Sheet name. Defaults to "Sheet1"
///     auto_filter (bool): Enable autofilter on headers
///     freeze_rows (int): Number of rows to freeze
///     freeze_cols (int): Number of columns to freeze
///     auto_width (bool): Auto-calculate column widths
///     styled_headers (bool): Apply bold+gray style to headers
///     write_header_row (bool): Write header row with column names
///     column_widths (dict[str, str|float], optional): Column widths - accepts:
///         - float/int: Excel character units (e.g., 15.5)
///         - "150px": Pixel width (converted to characters)
///         - "auto": Auto-calculate from data
///     column_formats (dict[str, str], optional): Number formats: "integer", "decimal2", "currency", "date", "percentage", etc.
///     merge_cells (list[tuple], optional): List of (start_row, start_col, end_row, end_col)
///     data_validations (list[dict], optional): Data validation rules
///     hyperlinks (list[tuple], optional): List of (row, col, url, display_text)
///     row_heights (dict[int, float], optional): Custom row heights
///     cell_styles (list[dict], optional): Custom cell styles with font, fill, border, alignment
///     formulas (list[tuple], optional): List of (row, col, formula, cached_value)
///     conditional_formats (list[dict], optional): Conditional formatting rules
///     tables (list[dict], optional): Excel table definitions
///     charts (list[dict], optional): Chart definitions
///     images (list[dict], optional): Image definitions
///     gridlines_visible (bool): Show gridlines (default: True)
///     zoom_scale (int, optional): Zoom level 10-400%
///     tab_color (str, optional): Sheet tab color in RGB format (e.g., "FFFF0000")
///     default_row_height (float, optional): Default row height for all rows
///     hidden_columns (list[int], optional): Column indices to hide
///     hidden_rows (list[int], optional): Row indices to hide
///     right_to_left (bool): Enable right-to-left layout (default: False)
///     data_start_row (int): Skip this many rows when calculating auto_width (for dummy rows)
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
    write_header_row: bool,
    column_widths: Option<HashMap<String, Bound<PyAny>>>,
    column_formats: Option<HashMap<String, String>>,
    merge_cells: Option<Vec<(usize, usize, usize, usize)>>,
    data_validations: Option<Vec<Bound<PyDict>>>,
    hyperlinks: Option<Vec<(usize, usize, String, Option<String>)>>,
    row_heights: Option<HashMap<usize, f64>>,
    cell_styles: Option<Vec<Bound<PyDict>>>,
    formulas: Option<Vec<(usize, usize, String, Option<String>)>>,
    conditional_formats: Option<Vec<Bound<PyDict>>>,
    tables: Option<Vec<Bound<PyDict>>>,
    charts: Option<Vec<Bound<PyDict>>>,
    images: Option<Vec<Bound<PyDict>>>,
    gridlines_visible: bool,
    zoom_scale: Option<u16>,
    tab_color: Option<String>,
    default_row_height: Option<f64>,
    hidden_columns: Option<Vec<usize>>,
    hidden_rows: Option<Vec<usize>>,
    right_to_left: bool,
    data_start_row: usize,
) -> PyResult<()> {
    // Convert PyArrow data to RecordBatch
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

    // Parse column_widths - supports float, "auto", or "150px"
    let parsed_column_widths = column_widths.map(|cw| {
        cw.into_iter()
            .filter_map(|(k, v)| {
                let width = if let Ok(s) = v.extract::<String>() {
                    if s.to_lowercase() == "auto" {
                        ColumnWidth::Auto
                    } else if s.ends_with("px") {
                        let px: f64 = s.trim_end_matches("px").parse().unwrap_or(50.0);
                        ColumnWidth::Pixels(px)
                    } else {
                        // Try parsing as number string
                        ColumnWidth::Characters(s.parse().unwrap_or(8.43))
                    }
                } else if let Ok(f) = v.extract::<f64>() {
                    ColumnWidth::Characters(f)
                } else if let Ok(i) = v.extract::<i64>() {
                    ColumnWidth::Characters(i as f64)
                } else {
                    return None;
                };
                Some((k, width))
            })
            .collect()
    });

    // Build config
    let mut config = StyleConfig {
        auto_filter,
        freeze_rows,
        freeze_cols,
        styled_headers,
        write_header_row,
        column_widths: parsed_column_widths,
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
        cond_format_dxf_ids: HashMap::new(), 
        tables: Vec::new(), 
        charts: Vec::new(),
        images: Vec::new(),
        gridlines_visible,
        zoom_scale,
        tab_color,
        default_row_height,
        hidden_columns: hidden_columns.unwrap_or_default(),
        hidden_rows: hidden_rows.unwrap_or_default(),
        right_to_left,
        data_start_row,
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
        for cond_dict in cond_formats {
            if let Ok(cond_format) = extract_conditional_format(&cond_dict) {
                config.conditional_formats.push(cond_format);
            }
        }
    }

    // Parse tables
    if let Some(tables_vec) = tables {
        for table_dict in tables_vec {
            if let Ok(table) = extract_table(&table_dict) {
                config.tables.push(table);
            }
        }
    }

    // Parse charts
    if let Some(charts_vec) = charts {
        for chart_dict in charts_vec {
            if let Ok(chart) = extract_chart(&chart_dict) {
                config.charts.push(chart);
            }
        }
    }

    // Parse images
    if let Some(images_vec) = images {
        for image_dict in images_vec {
            if let Ok(image) = extract_image(&image_dict) {
                config.images.push(image);
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
/// Write multiple Arrow tables to Excel with parallel processing.
///
/// Args:
///     arrow_sheets (list[dict]): List of dicts with keys: data, name, and optional formatting params
///     filename (str): Output file path
///     num_threads (int): Number of parallel threads for XML generation
fn write_sheets_arrow(
    py: Python,
    arrow_sheets: Vec<Bound<PyDict>>,
    filename: String,
    num_threads: usize,
) -> PyResult<()> {
    // Collect sheets with owned data first
    let mut sheets_data: Vec<(Vec<RecordBatch>, String, StyleConfig)> = Vec::new();
    
    for sheet_dict in arrow_sheets {
        let arrow_data = sheet_dict.get_item("data")?.ok_or_else(|| 
            PyErr::new::<pyo3::exceptions::PyKeyError, _>("Missing 'data' key"))?;
        let name: String = sheet_dict.get_item("name")?.ok_or_else(|| 
            PyErr::new::<pyo3::exceptions::PyKeyError, _>("Missing 'name' key"))?.extract()?;
        
        let any_batch = AnyRecordBatch::extract_bound(&arrow_data)?;
        let reader = any_batch.into_reader()?;
        let batches: Vec<RecordBatch> = reader
            .collect::<Result<Vec<_>, _>>()
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(
                format!("Failed to read Arrow data: {}", e)
            ))?;
        
        // Build config from optional parameters
        let mut config = StyleConfig::default();
        
        if let Some(auto_filter) = sheet_dict.get_item("auto_filter")?.and_then(|v| v.extract().ok()) {
            config.auto_filter = auto_filter;
        }
        if let Some(freeze_rows) = sheet_dict.get_item("freeze_rows")?.and_then(|v| v.extract().ok()) {
            config.freeze_rows = freeze_rows;
        }
        if let Some(freeze_cols) = sheet_dict.get_item("freeze_cols")?.and_then(|v| v.extract().ok()) {
            config.freeze_cols = freeze_cols;
        }

        // Extract column_formats
        if let Some(formats) = sheet_dict.get_item("column_formats")? {
            let formats_dict = formats.downcast::<PyDict>()?;
            let mut col_fmts = HashMap::new();
            for (key, value) in formats_dict.iter() {
                let col_name: String = key.extract()?;
                let fmt_str: String = value.extract()?;
                if let Some(fmt) = parse_number_format(&fmt_str) {
                    col_fmts.insert(col_name, fmt);
                }
            }
            config.column_formats = Some(col_fmts);
        }

        // Charts
        if let Some(charts_vec) = sheet_dict.get_item("charts")? {
            let charts_list = charts_vec.downcast::<pyo3::types::PyList>()?;
            for chart_dict in charts_list.iter() {
                if let Ok(chart_dict) = chart_dict.downcast::<PyDict>() {
                    if let Ok(chart) = extract_chart(&chart_dict) {
                        config.charts.push(chart);
                    }
                }
            }
        }

        // Images
        if let Some(images_vec) = sheet_dict.get_item("images")? {
            let images_list = images_vec.downcast::<pyo3::types::PyList>()?;
            for image_dict in images_list.iter() {
                if let Ok(image_dict) = image_dict.downcast::<PyDict>() {
                    if let Ok(image) = extract_image(&image_dict) {
                        config.images.push(image);
                    }
                }
            }
        }
        
        // additions
        if let Some(val) = sheet_dict.get_item("gridlines_visible")?.and_then(|v| v.extract().ok()) {
            config.gridlines_visible = val;
        }
        if let Some(val) = sheet_dict.get_item("zoom_scale")?.and_then(|v| v.extract().ok()) {
            config.zoom_scale = Some(val);
        }
        if let Some(val) = sheet_dict.get_item("tab_color")?.and_then(|v| v.extract().ok()) {
            config.tab_color = Some(val);
        }
        if let Some(val) = sheet_dict.get_item("default_row_height")?.and_then(|v| v.extract().ok()) {
            config.default_row_height = Some(val);
        }
        if let Some(val) = sheet_dict.get_item("hidden_columns")?.and_then(|v| v.extract().ok()) {
            config.hidden_columns = val;
        }
        if let Some(val) = sheet_dict.get_item("hidden_rows")?.and_then(|v| v.extract().ok()) {
            config.hidden_rows = val;
        }
        if let Some(val) = sheet_dict.get_item("right_to_left")?.and_then(|v| v.extract().ok()) {
            config.right_to_left = val;
        }
        
        sheets_data.push((batches, name, config));
    }
    
    // Create references for the writer
    let sheets_refs: Vec<(&[RecordBatch], &str, StyleConfig)> = sheets_data.iter()
        .map(|(b, n, c)| (b.as_slice(), n.as_str(), c.clone()))
        .collect();

    py.detach(|| {
        writer::write_multiple_sheets_arrow_with_configs(&sheets_refs, &filename, num_threads)
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
        "percentage_integer" => Some(NumberFormat::PercentageInteger),
        "currency" => Some(NumberFormat::Currency),
        "currency_rounded" => Some(NumberFormat::CurrencyRounded),
        "date" => Some(NumberFormat::Date),
        "datetime" => Some(NumberFormat::DateTime),
        "time" => Some(NumberFormat::Time),
        "scientific" => Some(NumberFormat::Scientific),
        "fraction" => Some(NumberFormat::Fraction),
        "fraction_two_digits" => Some(NumberFormat::FractionTwoDigits),
        "thousands" => Some(NumberFormat::ThousandsSeparator),
        _ => {
            // Treat unknown strings as custom format codes
            if s.is_empty() {
                None
            } else {
                Some(NumberFormat::Custom(s.to_string()))
            }
        }
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

fn extract_cell_style_inner(dict: &Bound<PyDict>) -> PyResult<CellStyle> {
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
    
    Ok(cell_style)
}

fn extract_cell_style(dict: &Bound<PyDict>) -> PyResult<CellStyleMap> {
    let row: usize = dict.get_item("row")?.unwrap().extract()?;
    let col: usize = dict.get_item("col")?.unwrap().extract()?;
    let style = extract_cell_style_inner(dict)?;
    
    Ok(CellStyleMap { row, col, style })
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
    
    // Extract style or use default
    let style = if let Some(style_dict) = dict.get_item("style")? {
        let style_dict = style_dict.downcast::<PyDict>()?;
        extract_cell_style_inner(style_dict)?
    } else {
        // Default: red bold text
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

fn extract_table(dict: &Bound<PyDict>) -> PyResult<ExcelTable> {
    let name: String = dict.get_item("name")?.unwrap().extract()?;
    let start_row: usize = dict.get_item("start_row")?.unwrap().extract()?;
    let start_col: usize = dict.get_item("start_col")?.unwrap().extract()?;
    let end_row: usize = dict.get_item("end_row")?.unwrap().extract()?;
    let end_col: usize = dict.get_item("end_col")?.unwrap().extract()?;
    
    let mut table = ExcelTable::new(name, (start_row, start_col, end_row, end_col));
    
    if let Some(display_name) = dict.get_item("display_name")?.and_then(|v| v.extract().ok()) {
        table.display_name = display_name;
    }
    
    if let Some(style) = dict.get_item("style")?.and_then(|v| v.extract().ok()) {
        table.style_name = Some(style);
    }
    
    table.show_first_column = dict.get_item("show_first_column")?.map(|v| v.extract()).unwrap_or(Ok(false))?;
    table.show_last_column = dict.get_item("show_last_column")?.map(|v| v.extract()).unwrap_or(Ok(false))?;
    table.show_row_stripes = dict.get_item("show_row_stripes")?.map(|v| v.extract()).unwrap_or(Ok(true))?;
    table.show_column_stripes = dict.get_item("show_column_stripes")?.map(|v| v.extract()).unwrap_or(Ok(false))?;
    
    Ok(table)
}

fn extract_chart(dict: &Bound<PyDict>) -> PyResult<ExcelChart> {
    let chart_type_str: String = dict.get_item("chart_type")?.unwrap().extract()?;
    let chart_type = match chart_type_str.as_str() {
        "column" => ChartType::Column,
        "bar" => ChartType::Bar,
        "line" => ChartType::Line,
        "pie" => ChartType::Pie,
        "scatter" => ChartType::Scatter,
        "area" => ChartType::Area,
        _ => return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>("Invalid chart type")),
    };
    
    let start_row: usize = dict.get_item("start_row")?.unwrap().extract()?;
    let start_col: usize = dict.get_item("start_col")?.unwrap().extract()?;
    let end_row: usize = dict.get_item("end_row")?.unwrap().extract()?;
    let end_col: usize = dict.get_item("end_col")?.unwrap().extract()?;
    
    let from_col: usize = dict.get_item("from_col")?.unwrap().extract()?;
    let from_row: usize = dict.get_item("from_row")?.unwrap().extract()?;
    let to_col: usize = dict.get_item("to_col")?.unwrap().extract()?;
    let to_row: usize = dict.get_item("to_row")?.unwrap().extract()?;
    
    let mut chart = ExcelChart::new(
        chart_type,
        (start_row, start_col, end_row, end_col),
        ChartPosition { from_col, from_row, to_col, to_row },
    );
    
    chart.title = dict.get_item("title")?.and_then(|v| v.extract().ok());
    chart.category_col = dict.get_item("category_col")?.and_then(|v| v.extract().ok());
    chart.show_legend = dict.get_item("show_legend")?.map(|v| v.extract()).unwrap_or(Ok(true))?;
    chart.x_axis_title = dict.get_item("x_axis_title")?.and_then(|v| v.extract().ok());
    chart.y_axis_title = dict.get_item("y_axis_title")?.and_then(|v| v.extract().ok()); 
    if let Some(names) = dict.get_item("series_names")?.and_then(|v| v.extract::<Vec<String>>().ok()) {
        chart.series_names = names;
    }
    
    Ok(chart)
}

fn extract_image(dict: &Bound<PyDict>) -> PyResult<ExcelImage> {
    let from_col: usize = dict.get_item("from_col")?.unwrap().extract()?;
    let from_row: usize = dict.get_item("from_row")?.unwrap().extract()?;
    let to_col: usize = dict.get_item("to_col")?.unwrap().extract()?;
    let to_row: usize = dict.get_item("to_row")?.unwrap().extract()?;
    
    let position = ImagePosition { from_col, from_row, to_col, to_row };
    
    let image = if let Some(path) = dict.get_item("path")? {
        let path_str: String = path.extract()?;
        ExcelImage::from_path(&path_str, position)
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to read image: {}", e)))?
    } else if let Some(data) = dict.get_item("data")? {
        let bytes: Vec<u8> = data.extract()?;
        let ext: String = dict.get_item("extension")?.unwrap().extract()?;
        ExcelImage::from_bytes(bytes, ext, position)
    } else {
        return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>("Image must have 'path' or 'data'"));
    };
    
    Ok(image)
}