/// Critical validation and safety module for jetxl
use crate::types::WriteError;
use crate::styles::*;
use std::collections::{HashMap, HashSet};
use std::path::Path;
use std::fs;
use std::io::Write;

// Excel hard limits
pub const MAX_ROWS: usize = 1_048_576;
pub const MAX_COLS: usize = 16_384;
const MAX_SHEET_NAME_LEN: usize = 31;
const MAX_ROW_HEIGHT: f64 = 409.5;
const MAX_COL_WIDTH: f64 = 255.0;
const INVALID_SHEET_CHARS: &str = "[]:*?/\\";

/// Comprehensive pre-write validation result
#[derive(Debug)]
pub struct ValidationResult {
    pub errors: Vec<String>,
    pub warnings: Vec<String>,
}

impl ValidationResult {
    pub fn new() -> Self {
        Self {
            errors: Vec::new(),
            warnings: Vec::new(),
        }
    }
    
    pub fn is_valid(&self) -> bool {
        self.errors.is_empty()
    }
    
    pub fn add_error(&mut self, msg: String) {
        self.errors.push(msg);
    }
    
    pub fn add_warning(&mut self, msg: String) {
        self.warnings.push(msg);
    }
    
    pub fn to_error(self) -> Result<(), WriteError> {
        if self.is_valid() {
            Ok(())
        } else {
            Err(WriteError::Validation(format!(
                "Validation failed with {} errors:\n{}",
                self.errors.len(),
                self.errors.join("\n")
            )))
        }
    }
}

/// Escape sheet names for use in Excel formulas
/// Example: "My Sheet" -> "'My Sheet'"
///          "Quote's Sheet" -> "'Quote''s Sheet'"
pub fn escape_sheet_name_for_formula(name: &str) -> String {
    let needs_quoting = name.contains(' ') 
        || name.contains('!') 
        || name.contains('\'')
        || name.chars().any(|c| !c.is_alphanumeric() && c != '_');
    
    if needs_quoting {
        format!("'{}'", name.replace('\'', "''"))
    } else {
        name.to_string()
    }
}

/// Validate sheet name meets Excel requirements
pub fn validate_sheet_name(name: &str) -> Result<(), String> {
    if name.is_empty() {
        return Err("Sheet name cannot be empty".to_string());
    }
    
    if name.len() > MAX_SHEET_NAME_LEN {
        return Err(format!(
            "Sheet name '{}' exceeds {} characters (has {})",
            name, MAX_SHEET_NAME_LEN, name.len()
        ));
    }
    
    for c in INVALID_SHEET_CHARS.chars() {
        if name.contains(c) {
            return Err(format!(
                "Sheet name '{}' contains invalid character '{}'",
                name, c
            ));
        }
    }
    
    // Check for control characters
    if name.chars().any(|c| c.is_control()) {
        return Err(format!("Sheet name '{}' contains control characters", name));
    }
    
    Ok(())
}

/// Validate all sheet names in a workbook
pub fn validate_sheet_names(names: &[&str]) -> ValidationResult {
    let mut result = ValidationResult::new();
    let mut seen = HashSet::new();
    
    for (idx, name) in names.iter().enumerate() {
        // Check individual name validity
        if let Err(e) = validate_sheet_name(name) {
            result.add_error(format!("Sheet {}: {}", idx + 1, e));
            continue;
        }
        
        // Check for duplicates (case-insensitive per Excel)
        let lower = name.to_lowercase();
        if seen.contains(&lower) {
            result.add_error(format!(
                "Duplicate sheet name '{}' (sheet names are case-insensitive)",
                name
            ));
        }
        seen.insert(lower);
    }
    
    result
}

/// Validate cell coordinates are within Excel limits
pub fn validate_cell_coords(row: usize, col: usize, context: &str) -> Result<(), String> {
    if row > MAX_ROWS {
        return Err(format!(
            "{}: Row {} is out of range (must be 1-{})",
            context, row, MAX_ROWS
        ));
    }
    
    if col >= MAX_COLS {
        return Err(format!(
            "{}: Column {} is out of range (must be 0-{})",
            context, col, MAX_COLS - 1
        ));
    }
    
    Ok(())
}

/// Validate merge cell range
pub fn validate_merge_range(merge: &MergeRange, max_row: usize, max_col: usize) -> Result<(), String> {
    if merge.start_row > merge.end_row {
        return Err(format!(
            "Merge cell: start_row {} > end_row {}",
            merge.start_row, merge.end_row
        ));
    }
    
    if merge.start_col > merge.end_col {
        return Err(format!(
            "Merge cell: start_col {} > end_col {}",
            merge.start_col, merge.end_col
        ));
    }
    
    validate_cell_coords(merge.start_row, merge.start_col, "Merge cell start")?;
    validate_cell_coords(merge.end_row, merge.end_col, "Merge cell end")?;
    
    if merge.end_row > max_row {
        return Err(format!(
            "Merge cell end_row {} exceeds data rows {}",
            merge.end_row, max_row
        ));
    }
    
    if merge.end_col >= max_col {
        return Err(format!(
            "Merge cell end_col {} exceeds data columns {}",
            merge.end_col, max_col
        ));
    }
    
    Ok(())
}

/// Check if merge ranges overlap
pub fn validate_merge_overlaps(merges: &[MergeRange]) -> Result<(), String> {
    for (i, m1) in merges.iter().enumerate() {
        for (_j, m2) in merges.iter().enumerate().skip(i + 1) {
            if ranges_overlap(m1, m2) {
                return Err(format!(
                    "Merge ranges overlap: ({},{} to {},{}) and ({},{} to {},{})",
                    m1.start_row, m1.start_col, m1.end_row, m1.end_col,
                    m2.start_row, m2.start_col, m2.end_row, m2.end_col
                ));
            }
        }
    }
    Ok(())
}

fn ranges_overlap(m1: &MergeRange, m2: &MergeRange) -> bool {
    !(m1.end_row < m2.start_row 
      || m1.start_row > m2.end_row
      || m1.end_col < m2.start_col
      || m1.start_col > m2.end_col)
}

/// Validate table configuration
pub fn validate_table(table: &ExcelTable, max_row: usize, max_col: usize) -> Result<(), String> {
    let (start_row, start_col, end_row, end_col) = table.range;
    
    if table.name.is_empty() {
        return Err("Table name cannot be empty".to_string());
    }
    
    if table.display_name.is_empty() {
        return Err("Table display_name cannot be empty".to_string());
    }
    
    // Table names must be valid Excel identifiers
    if !table.name.chars().next().unwrap().is_alphabetic() && table.name.chars().next().unwrap() != '_' {
        return Err(format!("Table name '{}' must start with letter or underscore", table.name));
    }
    
    if !table.name.chars().all(|c| c.is_alphanumeric() || c == '_') {
        return Err(format!("Table name '{}' contains invalid characters", table.name));
    }
    
    validate_cell_coords(start_row, start_col, "Table start")?;
    validate_cell_coords(end_row, end_col, "Table end")?;
    
    if start_row > end_row || start_col > end_col {
        return Err(format!(
            "Table '{}': invalid range ({},{}) to ({},{})",
            table.name, start_row, start_col, end_row, end_col
        ));
    }
    
    if end_row > max_row || end_col >= max_col {
        return Err(format!(
            "Table '{}': range exceeds data bounds",
            table.name
        ));
    }
    
    // Table must have at least header row + 1 data row
    if end_row <= start_row {
        return Err(format!("Table '{}': must have at least 2 rows", table.name));
    }
    
    Ok(())
}

/// Validate all tables don't have name collisions
pub fn validate_table_names(tables: &[ExcelTable]) -> Result<(), String> {
    let mut seen = HashSet::new();
    
    for table in tables {
        let lower = table.name.to_lowercase();
        if seen.contains(&lower) {
            return Err(format!(
                "Duplicate table name '{}' (table names are case-insensitive)",
                table.name
            ));
        }
        
        // Display names must also be unique
        let display_lower = table.display_name.to_lowercase();
        if display_lower != lower && seen.contains(&display_lower) {
            return Err(format!(
                "Table display name '{}' conflicts with existing table name",
                table.display_name
            ));
        }
        
        seen.insert(lower);
    }
    
    Ok(())
}

/// Validate chart configuration
pub fn validate_chart(chart: &ExcelChart, max_row: usize, max_col: usize) -> Result<(), String> {
    let (start_row, start_col, end_row, end_col) = chart.data_range;
    
    validate_cell_coords(start_row, start_col, "Chart data start")?;
    validate_cell_coords(end_row, end_col, "Chart data end")?;
    
    if start_row > end_row || start_col > end_col {
        return Err("Chart: invalid data range".to_string());
    }
    
    if end_row > max_row || end_col >= max_col {
        return Err("Chart: data range exceeds sheet bounds".to_string());
    }
    
    // Validate category column is within range
    if let Some(cat_col) = chart.category_col {
        if cat_col < start_col || cat_col > end_col {
            return Err(format!(
                "Chart: category_col {} is outside data range {}-{}",
                cat_col, start_col, end_col
            ));
        }
    }
    
    Ok(())
}

/// Validate row heights
pub fn validate_row_heights(heights: &HashMap<usize, f64>) -> Result<(), String> {
    for (row, height) in heights {
        if *row == 0 || *row > MAX_ROWS {
            return Err(format!("Row height: row {} out of range", row));
        }
        
        if *height < 0.0 || *height > MAX_ROW_HEIGHT {
            return Err(format!(
                "Row {}: height {} out of range (0-{})",
                row, height, MAX_ROW_HEIGHT
            ));
        }
    }
    Ok(())
}

/// Validate column widths
pub fn validate_column_widths(widths: &HashMap<String, f64>) -> Result<(), String> {
    for (col_name, width) in widths {
        if *width < 0.0 || *width > MAX_COL_WIDTH {
            return Err(format!(
                "Column '{}': width {} out of range (0-{})",
                col_name, width, MAX_COL_WIDTH
            ));
        }
    }
    Ok(())
}

/// Validate freeze panes don't exceed data bounds
pub fn validate_freeze_panes(freeze_rows: usize, freeze_cols: usize, max_row: usize, max_col: usize) -> Result<(), String> {
    if freeze_rows >= max_row {
        return Err(format!(
            "freeze_rows {} exceeds data rows {}",
            freeze_rows, max_row
        ));
    }
    
    if freeze_cols >= max_col {
        return Err(format!(
            "freeze_cols {} exceeds data columns {}",
            freeze_cols, max_col
        ));
    }
    
    Ok(())
}

/// Validate conditional format priorities are unique
pub fn validate_conditional_formats(formats: &[ConditionalFormat]) -> Result<(), String> {
    let mut priorities = HashMap::new();
    
    for (idx, format) in formats.iter().enumerate() {
        if let Some(other_idx) = priorities.get(&format.priority) {
            return Err(format!(
                "Conditional formats {} and {} have duplicate priority {}",
                other_idx, idx, format.priority
            ));
        }
        priorities.insert(format.priority, idx);
        
        // Validate range
        let (start_row, start_col, end_row, end_col) = 
            (format.start_row, format.start_col, format.end_row, format.end_col);
        
        validate_cell_coords(start_row, start_col, "Conditional format start")?;
        validate_cell_coords(end_row, end_col, "Conditional format end")?;
        
        if start_row > end_row || start_col > end_col {
            return Err(format!("Conditional format {}: invalid range", idx));
        }
    }
    
    Ok(())
}

/// Comprehensive validation of StyleConfig
pub fn validate_style_config(config: &StyleConfig, max_row: usize, max_col: usize) -> ValidationResult {
    let mut result = ValidationResult::new();
    
    // Validate merge cells
    for (idx, merge) in config.merge_cells.iter().enumerate() {
        if let Err(e) = validate_merge_range(merge, max_row, max_col) {
            result.add_error(format!("Merge cell {}: {}", idx, e));
        }
    }
    
    if let Err(e) = validate_merge_overlaps(&config.merge_cells) {
        result.add_error(e);
    }
    
    // Validate tables
    for (idx, table) in config.tables.iter().enumerate() {
        if let Err(e) = validate_table(table, max_row, max_col) {
            result.add_error(format!("Table {}: {}", idx, e));
        }
    }
    
    if let Err(e) = validate_table_names(&config.tables) {
        result.add_error(e);
    }
    
    // Validate charts
    for (idx, chart) in config.charts.iter().enumerate() {
        if let Err(e) = validate_chart(chart, max_row, max_col) {
            result.add_error(format!("Chart {}: {}", idx, e));
        }
    }
    
    // Validate row heights
    if let Some(ref heights) = config.row_heights {
        if let Err(e) = validate_row_heights(heights) {
            result.add_error(e);
        }
    }
    
    // Validate column widths
    if let Some(ref widths) = config.column_widths {
        if let Err(e) = validate_column_widths(widths) {
            result.add_error(e);
        }
    }
    
    // Validate freeze panes
    if config.freeze_rows > 0 || config.freeze_cols > 0 {
        if let Err(e) = validate_freeze_panes(config.freeze_rows, config.freeze_cols, max_row, max_col) {
            result.add_error(e);
        }
    }
    
    // Validate conditional formats
    if !config.conditional_formats.is_empty() {
        if let Err(e) = validate_conditional_formats(&config.conditional_formats) {
            result.add_error(e);
        }
    }
    
    // Validate cell styles
    for (idx, style) in config.cell_styles.iter().enumerate() {
        if let Err(e) = validate_cell_coords(style.row, style.col, &format!("Cell style {}", idx)) {
            result.add_error(e);
        }
    }
    
    // Validate hyperlinks
    for (idx, link) in config.hyperlinks.iter().enumerate() {
        if let Err(e) = validate_cell_coords(link.row, link.col, &format!("Hyperlink {}", idx)) {
            result.add_error(e);
        }
        
        if link.url.is_empty() {
            result.add_error(format!("Hyperlink {}: URL cannot be empty", idx));
        }
    }
    
    // Validate formulas
    for (idx, formula) in config.formulas.iter().enumerate() {
        if let Err(e) = validate_cell_coords(formula.row, formula.col, &format!("Formula {}", idx)) {
            result.add_error(e);
        }
        
        if formula.formula.is_empty() {
            result.add_error(format!("Formula {}: formula string cannot be empty", idx));
        }
        
        // Check for formula + hyperlink conflict
        if config.hyperlinks.iter().any(|h| h.row == formula.row && h.col == formula.col) {
            result.add_warning(format!(
                "Formula {}: cell ({},{}) also has hyperlink - hyperlink will be ignored",
                idx, formula.row, formula.col
            ));
        }
    }
    
    result
}

/// Atomic file writing with rollback on error
pub fn write_file_atomic<F>(filename: &str, write_fn: F) -> Result<(), WriteError> 
where
    F: FnOnce(&mut fs::File) -> Result<(), WriteError>
{
    // Validate filename
    if filename.is_empty() {
        return Err(WriteError::Validation("Filename cannot be empty".to_string()));
    }
    
    let path = Path::new(filename);
    
    // Check if parent directory exists and is writable
    if let Some(parent) = path.parent() {
        if !parent.as_os_str().is_empty() && !parent.exists() {
            return Err(WriteError::Validation(format!(
                "Directory does not exist: {}",
                parent.display()
            )));
        }
    }
    
    // Create temporary file in same directory for atomic rename
    let temp_filename = format!("{}.tmp.{}", filename, std::process::id());
    let temp_path = Path::new(&temp_filename);
    
    // Write to temporary file
    let write_result = (|| {
        let mut temp_file = fs::File::create(temp_path)
            .map_err(|e| WriteError::Io(e))?;
        
        write_fn(&mut temp_file)?;
        
        temp_file.flush()
            .map_err(|e| WriteError::Io(e))?;
        temp_file.sync_all()
            .map_err(|e| WriteError::Io(e))?;
        
        Ok(())
    })();
    
    // Handle result
    match write_result {
        Ok(_) => {
            // Atomic rename
            fs::rename(temp_path, path)
                .map_err(|e| WriteError::Io(e))?;
            Ok(())
        }
        Err(e) => {
            // Cleanup temp file on error
            let _ = fs::remove_file(temp_path);
            Err(e)
        }
    }
}

/// Build proper DXF IDs for all sheets in a workbook
/// Fixes the DXF ID collision issue in multi-sheet workbooks
pub fn build_global_dxf_mappings(
    sheets_configs: &[&StyleConfig]
) -> (StyleRegistry, Vec<HashMap<usize, u32>>) {
    let mut registry = StyleRegistry::new();
    let mut all_sheet_mappings = Vec::new();
    
    for config in sheets_configs {
        let mut sheet_mapping = HashMap::new();
        
        // Register cell styles first
        for cell_style in &config.cell_styles {
            registry.register_cell_style(&cell_style.style);
        }
        
        // Register DXFs for conditional formats
        for (idx, cond_format) in config.conditional_formats.iter().enumerate() {
            match &cond_format.rule {
                ConditionalRule::CellValue { .. } | ConditionalRule::Top10 { .. } => {
                    registry.register_cell_style(&cond_format.style);
                    let dxf_id = registry.register_dxf(&cond_format.style);
                    sheet_mapping.insert(idx, dxf_id);
                }
                _ => {}
            }
        }
        
        all_sheet_mappings.push(sheet_mapping);
    }
    
    (registry, all_sheet_mappings)
}

#[cfg(test)]
mod tests {
    use super::*;
    
    #[test]
    fn test_escape_sheet_name() {
        assert_eq!(escape_sheet_name_for_formula("Sheet1"), "Sheet1");
        assert_eq!(escape_sheet_name_for_formula("My Sheet"), "'My Sheet'");
        assert_eq!(escape_sheet_name_for_formula("Quote's"), "'Quote''s'");
        assert_eq!(escape_sheet_name_for_formula("Sheet!"), "'Sheet!'");
    }
    
    #[test]
    fn test_validate_sheet_name() {
        assert!(validate_sheet_name("Sheet1").is_ok());
        assert!(validate_sheet_name("").is_err());
        assert!(validate_sheet_name(&"A".repeat(32)).is_err());
        assert!(validate_sheet_name("Invalid:Name").is_err());
    }
    
    #[test]
    fn test_duplicate_sheet_names() {
        let names = vec!["Sheet1", "sheet1"];
        let result = validate_sheet_names(&names);
        assert!(!result.is_valid());
    }
    
    #[test]
    fn test_merge_overlap() {
        let m1 = MergeRange { start_row: 1, start_col: 0, end_row: 3, end_col: 2 };
        let m2 = MergeRange { start_row: 2, start_col: 1, end_row: 4, end_col: 3 };
        assert!(ranges_overlap(&m1, &m2));
        
        let m3 = MergeRange { start_row: 5, start_col: 0, end_row: 7, end_col: 2 };
        assert!(!ranges_overlap(&m1, &m3));
    }
    
    #[test]
    fn test_cell_coords_overflow() {
        assert!(validate_cell_coords(MAX_ROWS + 1, 0, "test").is_err());
        assert!(validate_cell_coords(1, MAX_COLS, "test").is_err());
        assert!(validate_cell_coords(1, 0, "test").is_ok());
    }
}