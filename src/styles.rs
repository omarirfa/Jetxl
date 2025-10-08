use std::collections::HashMap;
use arrow_array::Array;
use arrow_schema::DataType;

#[derive(Debug, Clone, PartialEq)]
pub enum NumberFormat {
    General,
    Integer,
    Decimal2,
    Decimal4,
    Percentage,
    PercentageDecimal,
    Currency,
    CurrencyRounded,
    Date,
    DateTime,
    Time,
}

impl NumberFormat {
    pub fn num_fmt_id(&self) -> u32 {
        match self {
            NumberFormat::General => 0,
            NumberFormat::Integer => 165,
            NumberFormat::Decimal2 => 166,
            NumberFormat::Decimal4 => 167,
            NumberFormat::Percentage => 9,
            NumberFormat::PercentageDecimal => 10,
            NumberFormat::Currency => 168,
            NumberFormat::CurrencyRounded => 169,
            NumberFormat::Date => 14,
            NumberFormat::DateTime => 164,
            NumberFormat::Time => 170,
        }
    }
}

#[derive(Debug, Clone)]
pub struct MergeRange {
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
}

#[derive(Debug, Clone)]
pub enum ValidationType {
    List(Vec<String>),
    WholeNumber { min: i64, max: i64 },
    Decimal { min: f64, max: f64 },
    TextLength { min: usize, max: usize },
}

#[derive(Debug, Clone)]
pub struct DataValidation {
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
    pub validation_type: ValidationType,
    pub error_title: Option<String>,
    pub error_message: Option<String>,
    pub show_dropdown: bool,
}

#[derive(Debug, Clone)]
pub struct Hyperlink {
    pub row: usize,
    pub col: usize,
    pub url: String,
    pub display: Option<String>,
}

// Cell-level styling
#[derive(Debug, Clone, PartialEq)]
pub struct CellStyle {
    pub font: Option<FontStyle>,
    pub fill: Option<FillStyle>,
    pub border: Option<BorderStyle>,
    pub alignment: Option<AlignmentStyle>,
    pub number_format: Option<NumberFormat>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct FontStyle {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub size: Option<f64>,
    pub color: Option<String>,  // RGB format: "FFRRGGBB"
    pub name: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct FillStyle {
    pub pattern_type: PatternType,
    pub fg_color: Option<String>,  // RGB format: "FFRRGGBB"
    pub bg_color: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum PatternType {
    None,
    Solid,
    Gray125,
}

#[derive(Debug, Clone, PartialEq)]
pub struct BorderStyle {
    pub left: Option<BorderSide>,
    pub right: Option<BorderSide>,
    pub top: Option<BorderSide>,
    pub bottom: Option<BorderSide>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct BorderSide {
    pub style: BorderLineStyle,
    pub color: Option<String>,  // RGB format: "FFRRGGBB"
}

#[derive(Debug, Clone, PartialEq)]
pub enum BorderLineStyle {
    Thin,
    Medium,
    Thick,
    Double,
    Dotted,
    Dashed,
}

#[derive(Debug, Clone, PartialEq)]
pub struct AlignmentStyle {
    pub horizontal: Option<HorizontalAlignment>,
    pub vertical: Option<VerticalAlignment>,
    pub wrap_text: bool,
    pub text_rotation: Option<i32>,  // 0-180 degrees, or 255 for vertical
}

#[derive(Debug, Clone, PartialEq)]
pub enum HorizontalAlignment {
    Left,
    Center,
    Right,
    Justify,
}

#[derive(Debug, Clone, PartialEq)]
pub enum VerticalAlignment {
    Top,
    Center,
    Bottom,
}

// Formula support
#[derive(Debug, Clone)]
pub struct Formula {
    pub row: usize,
    pub col: usize,
    pub formula: String,
    pub cached_value: Option<String>,  // Optional pre-calculated value
}

// NEW: Conditional formatting
#[derive(Debug, Clone)]
pub struct ConditionalFormat {
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
    pub rule: ConditionalRule,
    pub style: CellStyle,
    pub priority: u32,
}

#[derive(Debug, Clone)]
pub enum ConditionalRule {
    CellValue { operator: ComparisonOperator, value: String },
    ColorScale { min_color: String, max_color: String, mid_color: Option<String> },
    DataBar { color: String, show_value: bool },
    Top10 { rank: u32, bottom: bool },
}

#[derive(Debug, Clone)]
pub enum ComparisonOperator {
    GreaterThan,
    LessThan,
    Equal,
    NotEqual,
    GreaterThanOrEqual,
    LessThanOrEqual,
    Between,
}

#[derive(Debug, Clone)]
pub struct StyleConfig {
    pub auto_filter: bool,
    pub freeze_rows: usize,
    pub freeze_cols: usize,
    pub styled_headers: bool,
    pub column_widths: Option<HashMap<String, f64>>,
    pub auto_width: bool,
    pub column_formats: Option<HashMap<String, NumberFormat>>,
    pub merge_cells: Vec<MergeRange>,
    pub data_validations: Vec<DataValidation>,
    pub hyperlinks: Vec<Hyperlink>,
    
    
    pub row_heights: Option<HashMap<usize, f64>>,  // row index -> height in points
    pub cell_styles: Vec<CellStyleMap>,  // individual cell styles
    pub formulas: Vec<Formula>,
    pub conditional_formats: Vec<ConditionalFormat>,
}

#[derive(Debug, Clone)]
pub struct CellStyleMap {
    pub row: usize,
    pub col: usize,
    pub style: CellStyle,
}

impl Default for StyleConfig {
    fn default() -> Self {
        Self {
            auto_filter: false,
            freeze_rows: 0,
            freeze_cols: 0,
            styled_headers: false,
            column_widths: None,
            auto_width: false,
            column_formats: None,
            merge_cells: Vec::new(),
            data_validations: Vec::new(),
            hyperlinks: Vec::new(),
            row_heights: None,
            cell_styles: Vec::new(),
            formulas: Vec::new(),
            conditional_formats: Vec::new(),
        }
    }
}

/// Style registry to avoid duplication and optimize file size
pub struct StyleRegistry {
    fonts: Vec<FontStyle>,
    fills: Vec<FillStyle>,
    borders: Vec<BorderStyle>,
    cell_xfs: Vec<CellXfEntry>,
}

#[derive(Debug, Clone)]
struct CellXfEntry {
    num_fmt_id: u32,
    font_id: u32,
    fill_id: u32,
    border_id: u32,
    alignment: Option<AlignmentStyle>,
}

impl StyleRegistry {
    pub fn new() -> Self {
        // Start with default styles
        let mut registry = Self {
            fonts: vec![
                // 0: Normal
                FontStyle { bold: false, italic: false, underline: false, size: Some(11.0), color: None, name: Some("Calibri".to_string()) },
                // 1: Bold (for headers)
                FontStyle { bold: true, italic: false, underline: false, size: Some(11.0), color: None, name: Some("Calibri".to_string()) },
                // 2: Hyperlink (blue, underlined)
                FontStyle { bold: false, italic: false, underline: true, size: Some(11.0), color: Some("FF0000FF".to_string()), name: Some("Calibri".to_string()) },
            ],
            fills: vec![
                // 0: None
                FillStyle { pattern_type: PatternType::None, fg_color: None, bg_color: None },
                // 1: Gray125
                FillStyle { pattern_type: PatternType::Gray125, fg_color: None, bg_color: None },
                // 2: Light gray header
                FillStyle { pattern_type: PatternType::Solid, fg_color: Some("FFD9D9D9".to_string()), bg_color: None },
            ],
            borders: vec![
                // 0: No border
                BorderStyle { left: None, right: None, top: None, bottom: None },
            ],
            cell_xfs: vec![],
        };
        
        // Build default cellXfs entries
        registry.build_default_xfs();
        registry
    }
    
    fn build_default_xfs(&mut self) {
        self.cell_xfs = vec![
            // 0: Normal
            CellXfEntry { num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0, alignment: None },
            // 1: Date/DateTime
            CellXfEntry { num_fmt_id: 164, font_id: 0, fill_id: 0, border_id: 0, alignment: None },
            // 2: Bold header
            CellXfEntry { num_fmt_id: 0, font_id: 1, fill_id: 0, border_id: 0, alignment: None },
            // 3: Bold header with fill
            CellXfEntry { num_fmt_id: 0, font_id: 1, fill_id: 2, border_id: 0, alignment: None },
            // 4: Currency
            CellXfEntry { num_fmt_id: 168, font_id: 0, fill_id: 0, border_id: 0, alignment: None },
            // 5: Percentage
            CellXfEntry { num_fmt_id: 9, font_id: 0, fill_id: 0, border_id: 0, alignment: None },
            // 6: Percentage with decimal
            CellXfEntry { num_fmt_id: 10, font_id: 0, fill_id: 0, border_id: 0, alignment: None },
            // 7: Integer
            CellXfEntry { num_fmt_id: 165, font_id: 0, fill_id: 0, border_id: 0, alignment: None },
            // 8: Decimal
            CellXfEntry { num_fmt_id: 166, font_id: 0, fill_id: 0, border_id: 0, alignment: None },
            // 9: Hyperlink
            CellXfEntry { num_fmt_id: 0, font_id: 2, fill_id: 0, border_id: 0, alignment: None },
        ];
    }
    
    pub fn register_cell_style(&mut self, style: &CellStyle) -> u32 {
        let font_id = if let Some(ref font) = style.font {
            self.get_or_add_font(font)
        } else {
            0
        };
        
        let fill_id = if let Some(ref fill) = style.fill {
            self.get_or_add_fill(fill)
        } else {
            0
        };
        
        let border_id = if let Some(ref border) = style.border {
            self.get_or_add_border(border)
        } else {
            0
        };
        
        let num_fmt_id = if let Some(ref fmt) = style.number_format {
            fmt.num_fmt_id()
        } else {
            0
        };
        
        let entry = CellXfEntry {
            num_fmt_id,
            font_id,
            fill_id,
            border_id,
            alignment: style.alignment.clone(),
        };
        
        // Check if this combination already exists
        for (idx, xf) in self.cell_xfs.iter().enumerate() {
            if xf.num_fmt_id == entry.num_fmt_id 
                && xf.font_id == entry.font_id 
                && xf.fill_id == entry.fill_id 
                && xf.border_id == entry.border_id 
                && xf.alignment == entry.alignment {
                return idx as u32;
            }
        }
        
        // Add new entry
        self.cell_xfs.push(entry);
        (self.cell_xfs.len() - 1) as u32
    }
    
    fn get_or_add_font(&mut self, font: &FontStyle) -> u32 {
        for (idx, f) in self.fonts.iter().enumerate() {
            if f == font {
                return idx as u32;
            }
        }
        self.fonts.push(font.clone());
        (self.fonts.len() - 1) as u32
    }
    
    fn get_or_add_fill(&mut self, fill: &FillStyle) -> u32 {
        for (idx, f) in self.fills.iter().enumerate() {
            if f == fill {
                return idx as u32;
            }
        }
        self.fills.push(fill.clone());
        (self.fills.len() - 1) as u32
    }
    
    fn get_or_add_border(&mut self, border: &BorderStyle) -> u32 {
        for (idx, b) in self.borders.iter().enumerate() {
            if b == border {
                return idx as u32;
            }
        }
        self.borders.push(border.clone());
        (self.borders.len() - 1) as u32
    }
}

/// Generate enhanced styles.xml with custom cell styles
pub fn generate_styles_xml_enhanced(registry: &StyleRegistry) -> String {
    let mut xml = String::with_capacity(2000 + registry.fonts.len() * 200);
    
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    xml.push_str("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
    
    // 1. Number formats
    xml.push_str("<numFmts count=\"7\">\n");
    xml.push_str("  <numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd hh:mm:ss\"/>\n");
    xml.push_str("  <numFmt numFmtId=\"165\" formatCode=\"0\"/>\n");
    xml.push_str("  <numFmt numFmtId=\"166\" formatCode=\"0.00\"/>\n");
    xml.push_str("  <numFmt numFmtId=\"167\" formatCode=\"0.0000\"/>\n");
    xml.push_str("  <numFmt numFmtId=\"168\" formatCode=\"$#,##0.00\"/>\n");
    xml.push_str("  <numFmt numFmtId=\"169\" formatCode=\"$#,##0\"/>\n");
    xml.push_str("  <numFmt numFmtId=\"170\" formatCode=\"hh:mm:ss\"/>\n");
    xml.push_str("</numFmts>\n");
    
    // 2. Fonts
    xml.push_str(&format!("<fonts count=\"{}\">\n", registry.fonts.len()));
    for font in &registry.fonts {
        xml.push_str("  <font>");
        if font.bold { xml.push_str("<b/>"); }
        if font.italic { xml.push_str("<i/>"); }
        if font.underline { xml.push_str("<u/>"); }
        if let Some(size) = font.size {
            xml.push_str(&format!("<sz val=\"{}\"/>", size));
        }
        if let Some(ref color) = font.color {
            xml.push_str(&format!("<color rgb=\"{}\"/>", color));
        }
        if let Some(ref name) = font.name {
            xml.push_str(&format!("<name val=\"{}\"/>", name));
        }
        xml.push_str("</font>\n");
    }
    xml.push_str("</fonts>\n");
    
    // 3. Fills
    xml.push_str(&format!("<fills count=\"{}\">\n", registry.fills.len()));
    for fill in &registry.fills {
        xml.push_str("  <fill>");
        match fill.pattern_type {
            PatternType::None => xml.push_str("<patternFill patternType=\"none\"/>"),
            PatternType::Gray125 => xml.push_str("<patternFill patternType=\"gray125\"/>"),
            PatternType::Solid => {
                xml.push_str("<patternFill patternType=\"solid\">");
                if let Some(ref fg) = fill.fg_color {
                    xml.push_str(&format!("<fgColor rgb=\"{}\"/>", fg));
                }
                if let Some(ref bg) = fill.bg_color {
                    xml.push_str(&format!("<bgColor rgb=\"{}\"/>", bg));
                }
                xml.push_str("</patternFill>");
            }
        }
        xml.push_str("</fill>\n");
    }
    xml.push_str("</fills>\n");
    
    // 4. Borders
    xml.push_str(&format!("<borders count=\"{}\">\n", registry.borders.len()));
    for border in &registry.borders {
        xml.push_str("  <border>");
        write_border_side(&mut xml, "left", &border.left);
        write_border_side(&mut xml, "right", &border.right);
        write_border_side(&mut xml, "top", &border.top);
        write_border_side(&mut xml, "bottom", &border.bottom);
        xml.push_str("<diagonal/>");
        xml.push_str("</border>\n");
    }
    xml.push_str("</borders>\n");
    
    // cellStyleXfs - REQUIRED element (must come BEFORE cellXfs)
    // This defines the base formatting for cell styles
    xml.push_str("<cellStyleXfs count=\"1\">\n");
    xml.push_str("  <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n");
    xml.push_str("</cellStyleXfs>\n");
    
    // 6. Cell XFs
    xml.push_str(&format!("<cellXfs count=\"{}\">\n", registry.cell_xfs.len()));
    for xf in &registry.cell_xfs {
        xml.push_str(&format!("  <xf numFmtId=\"{}\" fontId=\"{}\" fillId=\"{}\" borderId=\"0\"", 
            xf.num_fmt_id, xf.font_id, xf.fill_id));
        
        let apply_font = xf.font_id > 0;
        let apply_fill = xf.fill_id > 0;
        let apply_num_fmt = xf.num_fmt_id > 0;
        let apply_alignment = xf.alignment.is_some();
        
        if apply_font { xml.push_str(" applyFont=\"1\""); }
        if apply_fill { xml.push_str(" applyFill=\"1\""); }
        if apply_num_fmt { xml.push_str(" applyNumberFormat=\"1\""); }
        if apply_alignment { xml.push_str(" applyAlignment=\"1\""); }
        
        if let Some(ref align) = xf.alignment {
            xml.push_str(">");
            xml.push_str("<alignment");
            if let Some(ref h) = align.horizontal {
                xml.push_str(&format!(" horizontal=\"{}\"", match h {
                    HorizontalAlignment::Left => "left",
                    HorizontalAlignment::Center => "center",
                    HorizontalAlignment::Right => "right",
                    HorizontalAlignment::Justify => "justify",
                }));
            }
            if let Some(ref v) = align.vertical {
                xml.push_str(&format!(" vertical=\"{}\"", match v {
                    VerticalAlignment::Top => "top",
                    VerticalAlignment::Center => "center",
                    VerticalAlignment::Bottom => "bottom",
                }));
            }
            if align.wrap_text {
                xml.push_str(" wrapText=\"1\"");
            }
            if let Some(rotation) = align.text_rotation {
                xml.push_str(&format!(" textRotation=\"{}\"", rotation));
            }
            xml.push_str("/>");
            xml.push_str("</xf>\n");
        } else {
            xml.push_str("/>\n");
        }
    }
    xml.push_str("</cellXfs>\n");
    
    // 7. cellStyles 
    xml.push_str("<cellStyles count=\"1\">\n");
    xml.push_str("  <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>\n");
    xml.push_str("</cellStyles>\n");
    
    // 8. dxfs - MUST come AFTER cellXfs for conditional formatting
    xml.push_str("<dxfs count=\"1\">\n");
    xml.push_str("  <dxf>\n");
    xml.push_str("    <font><b/><color rgb=\"FFFF0000\"/></font>\n");
    xml.push_str("    <fill><patternFill patternType=\"solid\"><bgColor rgb=\"FFFFEB9C\"/></patternFill></fill>\n");
    xml.push_str("  </dxf>\n");
    xml.push_str("</dxfs>\n");
    
    xml.push_str("</styleSheet>");
    xml
}

fn write_border_side(xml: &mut String, side: &str, border: &Option<BorderSide>) {
    if let Some(ref b) = border {
        xml.push_str(&format!("<{} style=\"{}\">", side, match b.style {
            BorderLineStyle::Thin => "thin",
            BorderLineStyle::Medium => "medium",
            BorderLineStyle::Thick => "thick",
            BorderLineStyle::Double => "double",
            BorderLineStyle::Dotted => "dotted",
            BorderLineStyle::Dashed => "dashed",
        }));
        if let Some(ref color) = b.color {
            xml.push_str(&format!("<color rgb=\"{}\"/>", color));
        }
        xml.push_str(&format!("</{}>", side));
    } else {
        xml.push_str(&format!("<{}/>", side));
    }
}

/// Generate the default styles.xml (for backward compatibility)
pub fn generate_styles_xml() -> String {
    let registry = StyleRegistry::new();
    generate_styles_xml_enhanced(&registry)
}

/// Calculate optimal column width from Arrow array
pub fn calculate_column_width(
    array: &dyn Array,
    header: &str,
    max_rows_to_scan: usize,
) -> f64 {
    use arrow_array::{StringArray, LargeStringArray};
    
    let mut max_len = header.len();
    
    if let Some(str_array) = array.as_any().downcast_ref::<StringArray>() {
        let rows_to_check = str_array.len().min(max_rows_to_scan);
        for i in 0..rows_to_check {
            if !str_array.is_null(i) {
                max_len = max_len.max(str_array.value(i).len());
            }
        }
    } else if let Some(str_array) = array.as_any().downcast_ref::<LargeStringArray>() {
        let rows_to_check = str_array.len().min(max_rows_to_scan);
        for i in 0..rows_to_check {
            if !str_array.is_null(i) {
                max_len = max_len.max(str_array.value(i).len());
            }
        }
    } else {
        max_len = match array.data_type() {
            DataType::Int8 | DataType::Int16 => 8,
            DataType::Int32 | DataType::Int64 => 12,
            DataType::UInt8 | DataType::UInt16 => 8,
            DataType::UInt32 | DataType::UInt64 => 12,
            DataType::Float32 | DataType::Float64 => 12,
            DataType::Boolean => 6,
            DataType::Date32 | DataType::Date64 => 12,
            DataType::Timestamp(_, _) => 20,
            _ => 10,
        }.max(header.len());
    }
    
    ((max_len as f64 * 1.2) + 2.0).min(100.0)
}