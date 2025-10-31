use arrow_array::Array;
use arrow_schema::DataType;
use std::collections::{HashMap, HashSet};

fn get_builtin_format_name(code: &str) -> Option<&'static str> {
    match code.to_lowercase().as_str() {
        "0" => Some("integer"),
        "0.00" => Some("decimal2"),
        "0.0000" => Some("decimal4"),
        "0%" => Some("percentage"),
        "0.00%" => Some("percentage_decimal"),
        "$#,##0.00" => Some("currency"),
        "$#,##0" => Some("currency_rounded"),
        "yyyy-mm-dd hh:mm:ss" => Some("datetime"),
        "hh:mm:ss" => Some("time"),
        "0.00e+00" => Some("scientific"),
        "# ?/?" => Some("fraction"),
        "# ??/??" => Some("fraction_two_digits"),
        "#,##0" => Some("thousands"),
        _ => None,
    }
}

fn is_likely_invalid_format(code: &str) -> bool {
    if code.is_empty() {
        return true;
    }
    
    // Pure alphabetic strings are likely user errors like "accounting" instead of format codes
    if code.chars().all(|c| c.is_alphabetic() || c.is_whitespace()) {
        return true;
    }
    
    false
}


#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum NumberFormat {
    // Built-in formats (use Excel's predefined IDs)
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
    
    // Extended built-in formats
    Scientific,
    Fraction,
    FractionTwoDigits,
    ThousandsSeparator,
    PercentageInteger,
    
    // Custom format string
    Custom(String),
}

impl NumberFormat {
    /// Returns the format ID and optional format code
    /// Built-in formats return (id, None), custom formats return (id, Some(code))
    pub fn fmt_info(&self) -> (u32, Option<&str>) {
        match self {
            NumberFormat::General => (0, None),
            NumberFormat::Integer => (165, None),
            NumberFormat::Decimal2 => (166, None),
            NumberFormat::Decimal4 => (167, None),
            NumberFormat::Percentage => (9, None),
            NumberFormat::PercentageDecimal => (10, None),
            NumberFormat::Currency => (168, None),
            NumberFormat::CurrencyRounded => (169, None),
            NumberFormat::Date => (14, None),
            NumberFormat::DateTime => (164, None),
            NumberFormat::Time => (170, None),
            NumberFormat::Scientific => (175, None),           // was: Some("0.00E+00")
            NumberFormat::Fraction => (176, None),             // was: Some("# ?/?")
            NumberFormat::FractionTwoDigits => (177, None),    // was: Some("# ??/??")
            NumberFormat::ThousandsSeparator => (173, None),
            NumberFormat::PercentageInteger => (174, None),
            NumberFormat::Custom(ref code) => (0, Some(code.as_str())), // ID assigned by registry
        }
    }
    
    pub fn is_custom(&self) -> bool {
        matches!(self, NumberFormat::Custom(_))
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
    pub color: Option<String>,
    pub name: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct FillStyle {
    pub pattern_type: PatternType,
    pub fg_color: Option<String>,
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
    pub color: Option<String>,
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
    pub text_rotation: Option<i32>,
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

#[derive(Debug, Clone)]
pub struct Formula {
    pub row: usize,
    pub col: usize,
    pub formula: String,
    pub cached_value: Option<String>,
}

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
pub struct ExcelTable {
    pub name: String,
    pub display_name: String,
    pub range: (usize, usize, usize, usize), // start_row, start_col, end_row, end_col
    pub style_name: Option<String>, // "TableStyleMedium2", etc.
    pub show_first_column: bool,
    pub show_last_column: bool,
    pub show_row_stripes: bool,
    pub show_column_stripes: bool,
    pub show_header_row: bool,
    pub show_totals_row: bool,
    pub column_names: Vec<String>, // Auto-generated from headers if not provided
}

impl ExcelTable {
    pub fn new(name: String, range: (usize, usize, usize, usize)) -> Self {
        Self {
            display_name: name.clone(),
            name,
            range,
            style_name: Some("TableStyleMedium2".to_string()),
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: true,
            show_column_stripes: false,
            show_header_row: true,
            show_totals_row: false,
            column_names: Vec::new(),
        }
    }
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
    pub write_header_row: bool,
    pub column_widths: Option<HashMap<String, ColumnWidth>>,
    pub auto_width: bool,
    pub column_formats: Option<HashMap<String, NumberFormat>>,
    pub merge_cells: Vec<MergeRange>,
    pub data_validations: Vec<DataValidation>,
    pub hyperlinks: Vec<Hyperlink>,
    pub row_heights: Option<HashMap<usize, f64>>,
    pub cell_styles: Vec<CellStyleMap>,
    pub formulas: Vec<Formula>,
    pub conditional_formats: Vec<ConditionalFormat>,
    pub cond_format_dxf_ids: HashMap<usize, u32>,
    pub tables: Vec<ExcelTable>,
    pub charts: Vec<ExcelChart>,
    pub images: Vec<ExcelImage>,
    pub gridlines_visible: bool,
    pub zoom_scale: Option<u16>, // 10-400
    pub tab_color: Option<String>, // RGB like "FFFF0000"
    pub default_row_height: Option<f64>,
    pub hidden_columns: HashSet<usize>,
    pub hidden_rows: HashSet<usize>,
    pub right_to_left: bool,
    pub data_start_row: usize,
    pub header_content: Vec<(usize, usize, String)>,
}

#[derive(Debug, Clone)]
pub enum ColumnWidth {
    Characters(f64),  // Excel native units
    Pixels(f64),      // Convert to characters: px / 7.0
    Auto,             // Calculate from data
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
            write_header_row: true,
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
            cond_format_dxf_ids: HashMap::new(),
            tables: Vec::new(),
            charts: Vec::new(),
            images: Vec::new(),
            gridlines_visible: true,
            zoom_scale: None,
            tab_color: None,
            default_row_height: None,
            hidden_columns: HashSet::new(),
            hidden_rows: HashSet::new(),
            right_to_left: false,
            data_start_row: 0,
            header_content: Vec::new(),
        }
    }
}

pub struct StyleRegistry {
    fonts: Vec<FontStyle>,
    fills: Vec<FillStyle>,
    borders: Vec<BorderStyle>,
    cell_xfs: Vec<CellXfEntry>,
    dxfs: Vec<CellStyle>,
    custom_num_fmts: Vec<(u32, String)>, // (id, format_code)
    next_custom_fmt_id: u32,
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
        let mut registry = Self {
            fonts: vec![
                FontStyle { bold: false, italic: false, underline: false, size: Some(11.0), color: None, name: Some("Calibri".to_string()) },
                FontStyle { bold: true, italic: false, underline: false, size: Some(11.0), color: None, name: Some("Calibri".to_string()) },
                FontStyle { bold: false, italic: false, underline: true, size: Some(11.0), color: Some("FF0000FF".to_string()), name: Some("Calibri".to_string()) },
            ],
            fills: vec![
                FillStyle { pattern_type: PatternType::None, fg_color: None, bg_color: None },
                FillStyle { pattern_type: PatternType::Gray125, fg_color: None, bg_color: None },
                FillStyle { pattern_type: PatternType::Solid, fg_color: Some("FFD9D9D9".to_string()), bg_color: None },
            ],
            borders: vec![
                BorderStyle { left: None, right: None, top: None, bottom: None },
            ],
            cell_xfs: vec![],
            dxfs: Vec::new(),
            custom_num_fmts: Vec::new(),
            next_custom_fmt_id: 178,
        };
        
        registry.build_default_xfs();
        registry
    }
    
    fn build_default_xfs(&mut self) {
        self.cell_xfs = vec![
            CellXfEntry { num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0, alignment: None },
            CellXfEntry { num_fmt_id: 164, font_id: 0, fill_id: 0, border_id: 0, alignment: None }, // datetime
            CellXfEntry { num_fmt_id: 0, font_id: 1, fill_id: 0, border_id: 0, alignment: None },
            CellXfEntry { num_fmt_id: 0, font_id: 1, fill_id: 2, border_id: 0, alignment: None },
            CellXfEntry { num_fmt_id: 168, font_id: 0, fill_id: 0, border_id: 0, alignment: None },
            CellXfEntry { num_fmt_id: 9, font_id: 0, fill_id: 0, border_id: 0, alignment: None },
            CellXfEntry { num_fmt_id: 10, font_id: 0, fill_id: 0, border_id: 0, alignment: None },
            CellXfEntry { num_fmt_id: 165, font_id: 0, fill_id: 0, border_id: 0, alignment: None },
            CellXfEntry { num_fmt_id: 166, font_id: 0, fill_id: 0, border_id: 0, alignment: None },
            CellXfEntry { num_fmt_id: 0, font_id: 2, fill_id: 0, border_id: 0, alignment: None },
            CellXfEntry { num_fmt_id: 14, font_id: 0, fill_id: 0, border_id: 0, alignment: None }, 
        ];
    }
    fn get_or_add_num_fmt(&mut self, fmt: &NumberFormat) -> Result<u32, String> {
        let (base_id, code_opt) = fmt.fmt_info();
        
        match code_opt {
            None => Ok(base_id),
            Some(code) => {
                // Check for duplicates FIRST - O(n) but n is small and avoids validation
                for (id, existing_code) in &self.custom_num_fmts {
                    if existing_code == code {
                        return Ok(*id);
                    }
                }
                
                // Validate only NEW formats
                if is_likely_invalid_format(code) {
                    return Err(format!(
                        "Invalid format code: '{}'. Use valid Excel format codes like '0.00', '#,##0', or named formats like 'currency', 'decimal2'", 
                        code
                    ));
                }
                
                if let Some(builtin_name) = get_builtin_format_name(code) {
                    eprintln!("Warning: Format code '{}' matches built-in format '{}'. Recommend using column_formats={{'column': '{}'}}", 
                        code, builtin_name, builtin_name);
                }
                
                // Add new custom format
                let id = self.next_custom_fmt_id;
                self.custom_num_fmts.push((id, code.to_string()));
                self.next_custom_fmt_id += 1;
                Ok(id)
            }
        }
    }
    
    pub fn register_cell_style(&mut self, style: &CellStyle) -> Result<u32, String> {
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
            self.get_or_add_num_fmt(fmt)?
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
        
        for (idx, xf) in self.cell_xfs.iter().enumerate() {
            if xf.num_fmt_id == entry.num_fmt_id 
                && xf.font_id == entry.font_id 
                && xf.fill_id == entry.fill_id 
                && xf.border_id == entry.border_id 
                && xf.alignment == entry.alignment {
                return Ok(idx as u32);
            }
        }
        
        self.cell_xfs.push(entry);
        Ok((self.cell_xfs.len() - 1) as u32)
    }
    
    pub fn register_dxf(&mut self, style: &CellStyle) -> u32 {
        self.dxfs.push(style.clone());
        (self.dxfs.len() - 1) as u32
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

pub fn generate_styles_xml_enhanced(registry: &StyleRegistry) -> String {
    let base_count = 14; // Base built-in custom formats (164-174)
    let total_count = base_count + registry.custom_num_fmts.len();
    
    let mut xml = String::with_capacity(
        2000 + registry.fonts.len() * 200 + 
        registry.custom_num_fmts.len() * 100
    );
    
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    xml.push_str("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
    
    // Number formats section
    if total_count > 0 {
        xml.push_str(&format!("<numFmts count=\"{}\">\n", total_count));
        
        // Base custom formats (164-174)
        xml.push_str("  <numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd hh:mm:ss\"/>\n");
        xml.push_str("  <numFmt numFmtId=\"165\" formatCode=\"0\"/>\n");
        xml.push_str("  <numFmt numFmtId=\"166\" formatCode=\"0.00\"/>\n");
        xml.push_str("  <numFmt numFmtId=\"167\" formatCode=\"0.0000\"/>\n");
        xml.push_str("  <numFmt numFmtId=\"168\" formatCode=\"$#,##0.00\"/>\n");
        xml.push_str("  <numFmt numFmtId=\"169\" formatCode=\"$#,##0\"/>\n");
        xml.push_str("  <numFmt numFmtId=\"170\" formatCode=\"hh:mm:ss\"/>\n");
        xml.push_str("  <numFmt numFmtId=\"171\" formatCode=\"_(* #,##0.00_);_(* (#,##0.00);_(* &quot;-&quot;??_);_(@_)\"/>\n");
        xml.push_str("  <numFmt numFmtId=\"172\" formatCode=\"_(* #,##0_);_(* (#,##0);_(* &quot;-&quot;_);_(@_)\"/>\n");
        xml.push_str("  <numFmt numFmtId=\"173\" formatCode=\"#,##0\"/>\n");
        xml.push_str("  <numFmt numFmtId=\"174\" formatCode=\"0%\"/>\n");
        xml.push_str("  <numFmt numFmtId=\"175\" formatCode=\"0.00E+00\"/>\n");
        xml.push_str("  <numFmt numFmtId=\"176\" formatCode=\"# ?/?\"/>\n");
        xml.push_str("  <numFmt numFmtId=\"177\" formatCode=\"# ??/??\"/>\n");
        
        // User-defined custom formats (175+)
        for (id, code) in &registry.custom_num_fmts {
            xml.push_str("  <numFmt numFmtId=\"");
            xml.push_str(&id.to_string());
            xml.push_str("\" formatCode=\"");
            xml_escape_format_code(code, &mut xml);
            xml.push_str("\"/>\n");
        }
        
        xml.push_str("</numFmts>\n");
    }
    
    // Fonts section
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
    
    // Fills section
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
    
    // Borders section
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
    
    // Cell style XFs
    xml.push_str("<cellStyleXfs count=\"1\">\n");
    xml.push_str("  <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n");
    xml.push_str("</cellStyleXfs>\n");
    
    // Cell XFs
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
    
    // Cell styles
    xml.push_str("<cellStyles count=\"1\">\n");
    xml.push_str("  <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>\n");
    xml.push_str("</cellStyles>\n");
    
    // DXFs (for conditional formatting)
    xml.push_str(&format!("<dxfs count=\"{}\">\n", registry.dxfs.len()));
    for dxf in &registry.dxfs {
        xml.push_str("  <dxf>");
        
        if let Some(ref font) = dxf.font {
            xml.push_str("<font>");
            if font.bold { xml.push_str("<b/>"); }
            if font.italic { xml.push_str("<i/>"); }
            if font.underline { xml.push_str("<u/>"); }
            if let Some(ref color) = font.color {
                xml.push_str(&format!("<color rgb=\"{}\"/>", color));
            }
            xml.push_str("</font>");
        }
        
        if let Some(ref fmt) = dxf.number_format {
            let (num_fmt_id, _) = fmt.fmt_info();
            xml.push_str(&format!("<numFmt numFmtId=\"{}\" formatCode=\"\"/>", num_fmt_id));
        }
        
        if let Some(ref fill) = dxf.fill {
            xml.push_str("<fill><patternFill patternType=\"solid\">");
            if let Some(ref fg) = fill.fg_color {
                xml.push_str(&format!("<fgColor rgb=\"{}\"/>", fg));
                if fill.bg_color.is_none() {
                    xml.push_str("<bgColor rgb=\"FFFFFFFF\"/>");
                }
            }
            if let Some(ref bg) = fill.bg_color {
                xml.push_str(&format!("<bgColor rgb=\"{}\"/>", bg));
            }
            xml.push_str("</patternFill></fill>");
        }
        
        if let Some(ref align) = dxf.alignment {
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
            xml.push_str("/>");
        }
        
        if let Some(ref border) = dxf.border {
            xml.push_str("<border>");
            write_border_side(&mut xml, "left", &border.left);
            write_border_side(&mut xml, "right", &border.right);
            write_border_side(&mut xml, "top", &border.top);
            write_border_side(&mut xml, "bottom", &border.bottom);
            xml.push_str("</border>");
        }
        
        xml.push_str("</dxf>\n");
    }
    xml.push_str("</dxfs>\n");
    
    xml.push_str("</styleSheet>");
    xml
}

// Helper function for XML escaping format codes
#[inline]
fn xml_escape_format_code(code: &str, xml: &mut String) {
    for ch in code.chars() {
        match ch {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"),
            '"' => xml.push_str("&quot;"),
            '\'' => xml.push_str("&apos;"),
            _ => xml.push(ch),
        }
    }
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

pub fn generate_styles_xml() -> String {
    let registry = StyleRegistry::new();
    generate_styles_xml_enhanced(&registry)
}

pub fn calculate_column_width(
    array: &dyn Array,
    header: &str,
    max_rows_to_scan: usize,
    skip_rows: usize,
) -> f64 {
    use arrow_array::{StringArray, LargeStringArray};
    
 let mut max_len = header.len();
    
    if let Some(str_array) = array.as_any().downcast_ref::<StringArray>() {
        let start_idx = skip_rows.min(str_array.len()); 
        let rows_to_check = str_array.len().min(start_idx + max_rows_to_scan);  
        for i in start_idx..rows_to_check {  
            if !str_array.is_null(i) {
                max_len = max_len.max(str_array.value(i).len());
            }
        }
    } else if let Some(str_array) = array.as_any().downcast_ref::<LargeStringArray>() {
        let start_idx = skip_rows.min(str_array.len());  
        let rows_to_check = str_array.len().min(start_idx + max_rows_to_scan);  
        for i in start_idx..rows_to_check {  
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

#[derive(Debug, Clone)]
pub struct ExcelChart {
    pub chart_type: ChartType,
    pub title: Option<String>,
    pub data_range: (usize, usize, usize, usize), // start_row, start_col, end_row, end_col
    pub position: ChartPosition,
    pub series_names: Vec<String>,
    pub category_col: Option<usize>, // Column index for category labels
    pub show_legend: bool,
    pub legend_position: LegendPosition,
    pub x_axis_title: Option<String>,
    pub y_axis_title: Option<String>, 
}

#[derive(Debug, Clone)]
pub struct ExcelImage {
    pub image_data: Vec<u8>,
    pub extension: String, // "png", "jpeg", etc.
    pub position: ImagePosition,
    pub description: Option<String>,
}

#[derive(Debug, Clone)]
pub struct ImagePosition {
    pub from_col: usize,
    pub from_row: usize,
    pub to_col: usize,
    pub to_row: usize,
}

impl ExcelImage {
    pub fn from_path(path: &str, position: ImagePosition) -> Result<Self, std::io::Error> {
        let data = std::fs::read(path)?;
        let ext = std::path::Path::new(path)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("png")
            .to_lowercase();
        
        Ok(Self {
            image_data: data,
            extension: ext,
            position,
            description: None,
        })
    }

    pub fn from_bytes(data: Vec<u8>, extension: String, position: ImagePosition) -> Self {
        Self {
            image_data: data,
            extension,
            position,
            description: None,
        }
    }
}




#[derive(Debug, Clone)]
pub enum ChartType {
    Column,
    Bar,
    Line,
    Pie,
    Scatter,
    Area,
}

#[derive(Debug, Clone)]
pub struct ChartPosition {
    pub from_col: usize,
    pub from_row: usize,
    pub to_col: usize,
    pub to_row: usize,
}
#[allow(dead_code)]
#[derive(Debug, Clone)]
pub enum LegendPosition {
    Right,
    Left,
    Top,
    Bottom,
    None,
}

impl ExcelChart {
    pub fn new(
        chart_type: ChartType,
        data_range: (usize, usize, usize, usize),
        position: ChartPosition,
    ) -> Self {
        Self {
            chart_type,
            title: None,
            data_range,
            position,
            series_names: Vec::new(),
            category_col: None,
            show_legend: true,
            legend_position: LegendPosition::Right,
            x_axis_title: None,
            y_axis_title: None,  
        }
    }
}