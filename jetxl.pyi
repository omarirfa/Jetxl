"""
JetXL - Fast Excel writer using Arrow and Rust

This module provides high-performance Excel file generation with native support
for Arrow, Polars, and Pandas DataFrames. Built in Rust for maximum speed.

Performance:
    - 10-100x faster than traditional Python Excel libraries
    - Zero-copy Arrow integration
    - Multi-threaded sheet generation
    - Memory-efficient streaming XML generation

Basic Usage:
    >>> import polars as pl
    >>> import jetxl
    >>> df = pl.DataFrame({"Name": ["Alice", "Bob"], "Age": [25, 30]})
    >>> jetxl.write_sheet_arrow(df.to_arrow(), "output.xlsx")

Advanced Usage:
    >>> jetxl.write_sheet_arrow(
    ...     df.to_arrow(),
    ...     "formatted.xlsx",
    ...     sheet_name="Sales Data",
    ...     auto_filter=True,
    ...     freeze_rows=1,
    ...     styled_headers=True,
    ...     column_formats={"Price": "currency", "Date": "date"},
    ...     conditional_formats=[{
    ...         "start_row": 2, "start_col": 0,
    ...         "end_row": 100, "end_col": 0,
    ...         "rule_type": "cell_value",
    ...         "operator": "greater_than",
    ...         "value": "100",
    ...         "priority": 1,
    ...         "style": {"font": {"bold": True, "color": "FFFF0000"}}
    ...     }]
    ... )
"""

from typing import Any, Optional, Literal, TypedDict, List, Dict, Tuple, Union

# =============================================================================
# NUMBER FORMATS
# =============================================================================

NumberFormat = Literal[
    "general",              # Default formatting
    "integer",              # Whole numbers: 0
    "decimal2",             # Two decimal places: 0.00
    "decimal4",             # Four decimal places: 0.0000
    "percentage",           # Percentage: 0%
    "percentage_decimal",   # Percentage with decimal: 0.00%
    "percentage_integer",   # Percentage integer: 0%
    "currency",             # Currency: $#,##0.00
    "currency_rounded",     # Rounded currency: $#,##0
    "date",                 # Date: yyyy-mm-dd
    "datetime",             # Date and time: yyyy-mm-dd hh:mm:ss
    "time",                 # Time: hh:mm:ss
    "scientific",           # Scientific notation: 0.00E+00
    "fraction",             # Fraction: # ?/?
    "fraction_two_digits",  # Fraction with 2 digits: # ??/??
    "thousands",            # Thousands separator: #,##0
] | str                     # Any string not matching above becomes a custom Excel format code

"""
Custom Number Formats
=====================

Any string not matching a built-in format is treated as a custom Excel number format code.
This allows full control over number display using Excel's format code syntax.

Excel Format Code Syntax:
    [Positive];[Negative];[Zero];[Text]
    
    You can specify 1-4 sections separated by semicolons:
    - 1 section: applies to all numbers
    - 2 sections: first for positive/zero, second for negative
    - 3 sections: positive, negative, zero
    - 4 sections: positive, negative, zero, text

Format Code Symbols:
    0       Digit placeholder (shows 0 if no digit)
    #       Digit placeholder (shows nothing if no digit)
    ?       Digit placeholder (adds space for alignment)
    .       Decimal point
    ,       Thousands separator
    %       Multiply by 100 and show percent sign
    E+ E-   Scientific notation
    $       Dollar sign (literal)
    -+()    Math symbols (literal)
    "text"  Literal text in quotes
    @       Text placeholder
    *       Repeat next character to fill cell width
    _       Skip width of next character
    [Color] Color code (e.g., [Red], [Blue], [Green])
    [>=100] Conditional format

Common Custom Format Examples:

Accounting format with negative in parentheses:
    >>> column_formats = {"Amount": "#,##0.00_);(#,##0.00)"}

Show thousands with 'K' suffix:
    >>> column_formats = {"Value": "#,##0,\"K\""}
    # 1500 displays as "1K"

Display millions:
    >>> column_formats = {"Revenue": "$#,##0.0,,\"M\""}
    # 5000000 displays as "$5.0M"

Custom date/time:
    >>> column_formats = {"Date": "dddd, mmmm dd, yyyy"}
    # Displays as "Monday, January 15, 2024"

Conditional coloring:
    >>> column_formats = {"Change": "[Green]#,##0;[Red]-#,##0;[Blue]0"}
    # Green for positive, red for negative, blue for zero

Fraction with specific denominator:
    >>> column_formats = {"Measurement": "# ?/16"}
    # Shows fractions in 16ths

Phone numbers:
    >>> column_formats = {"Phone": "(###) ###-####"}

Pad with zeros:
    >>> column_formats = {"ID": "00000"}
    # 42 displays as "00042"

Show positive/negative indicators:
    >>> column_formats = {"Delta": "+#,##0;-#,##0;0"}

Hide zeros:
    >>> column_formats = {"Value": "#,##0;-#,##0;\"\""}

Limitations and Safeguards:
    - Custom formats MUST be valid Excel format codes
    - No client-side validation - invalid codes may cause Excel errors
    - Special XML characters (<, >, &, ", ') are automatically escaped
    - Maximum format code length: ~255 characters (Excel limitation)
    - Some complex features (e.g., [DBNum1], locale codes) may not work in all Excel versions
    - Color names are limited to Excel's built-in set: [Red], [Blue], [Green], [Yellow], 
      [Cyan], [Magenta], [White], [Black], [Color1]-[Color56]

Examples in Code:

Basic Usage:
    >>> import jetxl
    >>> import polars as pl
    >>> 
    >>> df = pl.DataFrame({
    ...     "Amount": [1234.56, -789.12, 0],
    ...     "Percentage": [0.157, 0.932, 0.005],
    ...     "Code": [1, 42, 999]
    ... })
    >>> 
    >>> jetxl.write_sheet_arrow(
    ...     df.to_arrow(),
    ...     "custom_formats.xlsx",
    ...     column_formats={
    ...         "Amount": "$#,##0.00_);[Red]($#,##0.00)",  # Accounting format
    ...         "Percentage": "0.0%",                       # One decimal percent
    ...         "Code": "000000"                            # Zero-padded
    ...     }
    ... )

Advanced Custom Formats:
    >>> column_formats = {
    ...     "Revenue": "[>=1000000]$#,##0.0,,\"M\";[>=1000]$#,##0,\"K\";$#,##0",
    ...     "Quarter": "\"Q\"0",                           # Q1, Q2, Q3, Q4
    ...     "Ratio": "# ?/?;-# ?/?;\"N/A\"",              # Fractions with N/A for zero
    ...     "Status": "[=1]\"Active\";[=0]\"Inactive\";@", # Text based on value
    ... }

Testing Custom Formats:
    The easiest way to test custom formats is to:
    1. Create the format in Excel manually
    2. Right-click the cell → Format Cells → Custom
    3. Copy the format code from the "Type:" field
    4. Use that exact string in Jetxl

Reference: https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
"""

# =============================================================================
# FONT STYLING
# =============================================================================

class FontStyle(TypedDict, total=False):
    """Font styling options.
    
    Attributes:
        bold: Bold text
        italic: Italic text
        underline: Underlined text
        size: Font size in points (e.g., 11.0)
        color: Font color in ARGB hex format (e.g., "FFFF0000" for red)
        name: Font name (e.g., "Calibri", "Arial")
    
    Color Format (ARGB):
        - AA (Alpha): FF = fully opaque, 00 = fully transparent
        - RR (Red): 00 = no red, FF = maximum red
        - GG (Green): 00 = no green, FF = maximum green
        - BB (Blue): 00 = no blue, FF = maximum blue
    
    Common Colors:
        - "FFFF0000" - Red
        - "FF00FF00" - Green
        - "FF0000FF" - Blue
        - "FFFFFF00" - Yellow
        - "FF000000" - Black
        - "FFFFFFFF" - White
    
    Example:
        >>> font = {
        ...     "bold": True,
        ...     "italic": False,
        ...     "size": 14.0,
        ...     "color": "FFFF0000",  # Red
        ...     "name": "Arial"
        ... }
    """
    bold: bool
    italic: bool
    underline: bool
    size: float
    color: str  # ARGB hex: "FFFF0000" for red
    name: str

# =============================================================================
# FILL/BACKGROUND STYLING
# =============================================================================

class FillStyle(TypedDict, total=False):
    """Cell fill/background styling.
    
    Attributes:
        pattern: Fill pattern type
        fg_color: Foreground color in ARGB hex
        bg_color: Background color in ARGB hex
    
    Pattern Types:
        - "solid": Solid fill
        - "gray125": Light gray pattern
        - "none": No fill
    
    Example:
        >>> fill = {
        ...     "pattern": "solid",
        ...     "fg_color": "FFFFFF00",  # Yellow
        ...     "bg_color": None
        ... }
    """
    pattern: Literal["solid", "gray125", "none"]
    fg_color: str  # ARGB hex
    bg_color: str  # ARGB hex

# =============================================================================
# BORDER STYLING
# =============================================================================

class BorderSide(TypedDict, total=False):
    """Single border side styling.
    
    Attributes:
        style: Line style for the border
        color: Border color in ARGB hex
    
    Border Styles:
        - "thin": Thin line
        - "medium": Medium line
        - "thick": Thick line
        - "double": Double line
        - "dotted": Dotted line
        - "dashed": Dashed line
    
    Example:
        >>> border_side = {
        ...     "style": "thick",
        ...     "color": "FF000000"  # Black
        ... }
    """
    style: Literal["thin", "medium", "thick", "double", "dotted", "dashed"]
    color: str  # ARGB hex

class BorderStyle(TypedDict, total=False):
    """Cell border styling for all four sides.
    
    Attributes:
        left: Left border styling
        right: Right border styling
        top: Top border styling
        bottom: Bottom border styling
    
    Example:
        >>> border = {
        ...     "left": {"style": "thin", "color": "FF000000"},
        ...     "right": {"style": "thin", "color": "FF000000"},
        ...     "top": {"style": "medium", "color": "FF000000"},
        ...     "bottom": {"style": "medium", "color": "FF000000"}
        ... }
    """
    left: BorderSide
    right: BorderSide
    top: BorderSide
    bottom: BorderSide

# =============================================================================
# ALIGNMENT STYLING
# =============================================================================

class AlignmentStyle(TypedDict, total=False):
    """Cell alignment options.
    
    Attributes:
        horizontal: Horizontal alignment
        vertical: Vertical alignment
        wrap_text: Enable text wrapping
        text_rotation: Text rotation in degrees (0-180)
    
    Example:
        >>> alignment = {
        ...     "horizontal": "center",
        ...     "vertical": "center",
        ...     "wrap_text": True,
        ...     "text_rotation": 0
        ... }
    """
    horizontal: Literal["left", "center", "right", "justify"]
    vertical: Literal["top", "center", "bottom"]
    wrap_text: bool
    text_rotation: int  # 0-180 degrees

# =============================================================================
# COMPLETE CELL STYLE
# =============================================================================

class CellStyle(TypedDict, total=False):
    """Complete cell styling combining font, fill, border, alignment, and format.
    
    Attributes:
        font: Font styling options
        fill: Fill/background styling
        border: Border styling for all sides
        alignment: Text alignment options
        number_format: Number format for the cell
    
    Example - Highlighted Header:
        >>> header_style = {
        ...     "font": {
        ...         "bold": True,
        ...         "size": 12.0,
        ...         "color": "FFFFFFFF"  # White
        ...     },
        ...     "fill": {
        ...         "pattern": "solid",
        ...         "fg_color": "FF4472C4"  # Blue
        ...     },
        ...     "alignment": {
        ...         "horizontal": "center",
        ...         "vertical": "center"
        ...     }
        ... }
    
    Example - Currency Cell:
        >>> currency_style = {
        ...     "font": {
        ...         "bold": True,
        ...         "color": "FF00B050"  # Green
        ...     },
        ...     "number_format": "currency",
        ...     "alignment": {
        ...         "horizontal": "right"
        ...     }
        ... }
    """
    font: FontStyle
    fill: FillStyle
    border: BorderStyle
    alignment: AlignmentStyle
    number_format: NumberFormat

class CellStyleMap(TypedDict):
    """Cell style with position for applying to specific cells.
    
    Attributes:
        row: Row number (1-based, where 1 is the first row)
        col: Column number (0-based, where 0 is column A)
        font: Font styling (optional)
        fill: Fill styling (optional)
        border: Border styling (optional)
        alignment: Alignment styling (optional)
        number_format: Number format (optional)
    
    Example - Style cell B3:
        >>> cell_style = {
        ...     "row": 3,      # Row 3
        ...     "col": 1,      # Column B (0-based)
        ...     "font": {"bold": True, "color": "FFFF0000"},
        ...     "fill": {"pattern": "solid", "fg_color": "FFFFFF00"}
        ... }
    """
    row: int      # 1-based row number
    col: int      # 0-based column number
    font: FontStyle
    fill: FillStyle
    border: BorderStyle
    alignment: AlignmentStyle
    number_format: NumberFormat

# =============================================================================
# DATA VALIDATION
# =============================================================================

class DataValidationList(TypedDict):
    """Dropdown list validation.
    
    Attributes:
        start_row: Starting row (1-based)
        start_col: Starting column (0-based)
        end_row: Ending row (1-based)
        end_col: Ending column (0-based)
        type: Validation type (must be "list")
        items: List of valid options
        show_dropdown: Show dropdown arrow
        error_title: Error dialog title
        error_message: Error dialog message
    
    Example - Status dropdown:
        >>> validation = {
        ...     "start_row": 2, "start_col": 3,
        ...     "end_row": 100, "end_col": 3,
        ...     "type": "list",
        ...     "items": ["Pending", "In Progress", "Complete", "Cancelled"],
        ...     "show_dropdown": True,
        ...     "error_title": "Invalid Status",
        ...     "error_message": "Please select a valid status from the list"
        ... }
    """
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    type: Literal["list"]
    items: List[str]
    show_dropdown: bool
    error_title: str
    error_message: str

class DataValidationNumber(TypedDict):
    """Number range validation.
    
    Attributes:
        start_row: Starting row (1-based)
        start_col: Starting column (0-based)
        end_row: Ending row (1-based)
        end_col: Ending column (0-based)
        type: Validation type ("whole_number" or "decimal")
        min: Minimum allowed value
        max: Maximum allowed value
        show_dropdown: Show dropdown arrow
        error_title: Error dialog title
        error_message: Error dialog message
    
    Example - Age validation:
        >>> validation = {
        ...     "start_row": 2, "start_col": 2,
        ...     "end_row": 100, "end_col": 2,
        ...     "type": "whole_number",
        ...     "min": 18,
        ...     "max": 120,
        ...     "show_dropdown": False,
        ...     "error_title": "Invalid Age",
        ...     "error_message": "Age must be between 18 and 120"
        ... }
    """
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    type: Literal["whole_number", "decimal"]
    min: float
    max: float
    show_dropdown: bool
    error_title: str
    error_message: str

class DataValidationTextLength(TypedDict):
    """Text length validation.
    
    Attributes:
        start_row: Starting row (1-based)
        start_col: Starting column (0-based)
        end_row: Ending row (1-based)
        end_col: Ending column (0-based)
        type: Validation type (must be "text_length")
        min: Minimum text length
        max: Maximum text length
        show_dropdown: Show dropdown arrow
        error_title: Error dialog title
        error_message: Error dialog message
    
    Example - Username validation:
        >>> validation = {
        ...     "start_row": 2, "start_col": 0,
        ...     "end_row": 100, "end_col": 0,
        ...     "type": "text_length",
        ...     "min": 3,
        ...     "max": 20,
        ...     "show_dropdown": False,
        ...     "error_title": "Invalid Username",
        ...     "error_message": "Username must be 3-20 characters"
        ... }
    """
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    type: Literal["text_length"]
    min: int
    max: int
    show_dropdown: bool
    error_title: str
    error_message: str

DataValidation = DataValidationList | DataValidationNumber | DataValidationTextLength

# =============================================================================
# CONDITIONAL FORMATTING
# =============================================================================

class ConditionalFormatCellValue(TypedDict):
    """Cell value conditional formatting rule.
    
    Attributes:
        start_row: Starting row (1-based)
        start_col: Starting column (0-based)
        end_row: Ending row (1-based)
        end_col: Ending column (0-based)
        rule_type: Must be "cell_value"
        operator: Comparison operator
        value: Value to compare against (as string)
        priority: Rule priority (lower = higher priority)
        style: Style to apply when condition is met
    
    Operators:
        - "greater_than": Value > threshold
        - "less_than": Value < threshold
        - "equal": Value = threshold
        - "not_equal": Value ≠ threshold
        - "greater_than_or_equal": Value ≥ threshold
        - "less_than_or_equal": Value ≤ threshold
        - "between": min ≤ Value ≤ max
    
    Example - Highlight high values:
        >>> cond_format = {
        ...     "start_row": 2, "start_col": 3,
        ...     "end_row": 100, "end_col": 3,
        ...     "rule_type": "cell_value",
        ...     "operator": "greater_than",
        ...     "value": "1000",
        ...     "priority": 1,
        ...     "style": {
        ...         "font": {"bold": True, "color": "FFFF0000"},
        ...         "fill": {"pattern": "solid", "fg_color": "FFFFFF00"}
        ...     }
        ... }
    """
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    rule_type: Literal["cell_value"]
    operator: Literal[
        "greater_than",
        "less_than",
        "equal",
        "not_equal",
        "greater_than_or_equal",
        "less_than_or_equal",
        "between",
    ]
    value: str
    priority: int
    style: CellStyle

class ConditionalFormatColorScale(TypedDict):
    """Color scale conditional formatting (gradient).
    
    Attributes:
        start_row: Starting row (1-based)
        start_col: Starting column (0-based)
        end_row: Ending row (1-based)
        end_col: Ending column (0-based)
        rule_type: Must be "color_scale"
        min_color: Color for minimum values (ARGB hex)
        max_color: Color for maximum values (ARGB hex)
        mid_color: Optional color for midpoint values (ARGB hex)
        priority: Rule priority (lower = higher priority)
    
    Example - Red-Yellow-Green scale:
        >>> color_scale = {
        ...     "start_row": 2, "start_col": 2,
        ...     "end_row": 100, "end_col": 2,
        ...     "rule_type": "color_scale",
        ...     "min_color": "FFF8696B",  # Red
        ...     "mid_color": "FFFFEB84",  # Yellow
        ...     "max_color": "FF63BE7B",  # Green
        ...     "priority": 1
        ... }
    """
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    rule_type: Literal["color_scale"]
    min_color: str  # ARGB hex
    max_color: str  # ARGB hex
    mid_color: str  # Optional, ARGB hex
    priority: int

class ConditionalFormatDataBar(TypedDict):
    """Data bar conditional formatting (horizontal bars in cells).
    
    Attributes:
        start_row: Starting row (1-based)
        start_col: Starting column (0-based)
        end_row: Ending row (1-based)
        end_col: Ending column (0-based)
        rule_type: Must be "data_bar"
        color: Bar color (ARGB hex)
        show_value: Show cell value alongside bar
        priority: Rule priority (lower = higher priority)
    
    Example - Blue data bars:
        >>> data_bar = {
        ...     "start_row": 2, "start_col": 4,
        ...     "end_row": 50, "end_col": 4,
        ...     "rule_type": "data_bar",
        ...     "color": "FF638EC6",  # Blue
        ...     "show_value": True,
        ...     "priority": 1
        ... }
    """
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    rule_type: Literal["data_bar"]
    color: str  # ARGB hex
    show_value: bool
    priority: int

class ConditionalFormatTop10(TypedDict):
    """Top/Bottom N values conditional formatting.
    
    Attributes:
        start_row: Starting row (1-based)
        start_col: Starting column (0-based)
        end_row: Ending row (1-based)
        end_col: Ending column (0-based)
        rule_type: Must be "top10"
        rank: Number of top/bottom values to highlight
        bottom: If True, highlight bottom N; if False, highlight top N
        priority: Rule priority (lower = higher priority)
        style: Style to apply to top/bottom values
    
    Example - Highlight top 5:
        >>> top5 = {
        ...     "start_row": 2, "start_col": 5,
        ...     "end_row": 100, "end_col": 5,
        ...     "rule_type": "top10",
        ...     "rank": 5,
        ...     "bottom": False,
        ...     "priority": 1,
        ...     "style": {
        ...         "font": {"bold": True, "color": "FF0070C0"}
        ...     }
        ... }
    """
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    rule_type: Literal["top10"]
    rank: int
    bottom: bool  # False = top N, True = bottom N
    priority: int
    style: CellStyle

ConditionalFormat = (
    ConditionalFormatCellValue
    | ConditionalFormatColorScale
    | ConditionalFormatDataBar
    | ConditionalFormatTop10
)

# =============================================================================
# EXCEL TABLES
# =============================================================================

class ExcelTable(TypedDict, total=False):
    """Excel table definition with formatting and filtering capabilities.
    
    Excel tables provide:
    - Automatic header row with filter dropdowns
    - Structured references for formulas
    - Professional styling with banded rows/columns
    - Sorting and filtering capabilities
    
    Attributes:
        name: Internal table identifier (required, must be unique)
        start_row: First row of table including header (required, 1-based)
        start_col: First column of table (required, 0-based)
        end_row: Last row of table (required, 1-based)
        end_col: Last column of table (required, 0-based)
        display_name: User-friendly table name (optional)
        style: Excel table style name (optional)
        show_first_column: Bold first column (optional, default: False)
        show_last_column: Bold last column (optional, default: False)
        show_row_stripes: Alternating row colors (optional, default: True)
        show_column_stripes: Alternating column colors (optional, default: False)
        show_header_row: Show header row with filter dropdowns (optional, default: True)
        show_totals_row: Add totals row at bottom (optional, default: False)
        column_names: Custom column names (optional, auto-detected from data if not provided)

     
    Auto-Calculation:
        If end_row is 0 or omitted, it's automatically calculated as:
            start_row + total_data_rows
        
        If end_col is 0 or omitted, it's automatically calculated as:
            start_col + num_columns - 1
        
        This allows tables to adapt to your DataFrame size without manual counting.    

    Example - Auto-sized table:
        >>> table = {
        ...     "name": "SalesData",
        ...     "start_row": 1,
        ...     "start_col": 0,
        ...     "end_row": 0,      # Auto: will use all data rows
        ...     "end_col": 0       # Auto: will use all columns
        ... }
        

    Available Table Styles:
        Light Styles: TableStyleLight1 through TableStyleLight21
        Medium Styles: TableStyleMedium1 through TableStyleMedium28
        Dark Styles: TableStyleDark1 through TableStyleDark11
        
        Popular choices:
        - "TableStyleMedium2" - Blue theme, balanced design
        - "TableStyleMedium9" - Orange theme, professional
        - "TableStyleLight16" - Minimal gray theme
        - "TableStyleDark1" - High contrast black theme
    
    Example - Basic Sales Table:
        >>> table = {
        ...     "name": "SalesData",
        ...     "display_name": "Q1 Sales",
        ...     "start_row": 1,      # Header row
        ...     "start_col": 0,      # Column A
        ...     "end_row": 100,      # Row 100
        ...     "end_col": 5,        # Column F
        ...     "style": "TableStyleMedium2",
        ...     "show_row_stripes": True
        ... }
    
    Example - Formatted Report Table:
        >>> table = {
        ...     "name": "FinancialReport",
        ...     "start_row": 1,
        ...     "start_col": 0,
        ...     "end_row": 50,
        ...     "end_col": 8,
        ...     "style": "TableStyleMedium9",
        ...     "show_first_column": True,   # Bold first column
        ...     "show_last_column": True,    # Bold totals column
        ...     "show_row_stripes": True,
        ...     "show_column_stripes": False
        ... }
    Example - Table with Totals:
        >>> table = {
        ...     "name": "SalesData",
        ...     "start_row": 1,
        ...     "start_col": 0,
        ...     "end_row": 100,
        ...     "end_col": 3,
        ...     "style": "TableStyleMedium2",
        ...     "show_header_row": True,
        ...     "show_totals_row": True,  # Add totals row
        ...     "show_row_stripes": True
        ... }


    """
    name: str                   # Required: unique table identifier
    start_row: int              # Required: 1-based row
    start_col: int              # Required: 0-based column
    end_row: int                # Required: 1-based row
    end_col: int                # Required: 0-based column
    display_name: str           # Optional: user-friendly name
    style: str                  # Optional: table style name
    show_first_column: bool     # Optional: bold first column
    show_last_column: bool      # Optional: bold last column
    show_row_stripes: bool      # Optional: alternating rows
    show_column_stripes: bool   # Optional: alternating columns
    show_header_row: bool       # Optional: show header row with filters (default: True)
    show_totals_row: bool       # Optional: add totals row at bottom (default: False)
    column_names: List[str]     # Optional: custom column names (auto-detected if not provided)

# =============================================================================
# EXCEL CHARTS
# =============================================================================

class ChartPosition(TypedDict):
    """Chart position on the worksheet.
    
    Charts are positioned using Excel's column/row coordinates:
    - Columns are 0-indexed: A=0, B=1, C=2, etc.
    - Rows are 0-indexed for position: 0=row 1, 1=row 2, etc.
    
    Attributes:
        from_col: Starting column (0-based)
        from_row: Starting row (0-based)
        to_col: Ending column (0-based)
        to_row: Ending row (0-based)
    
    Size Recommendations:
        - Small: 6-8 columns × 12-15 rows
        - Medium: 8-12 columns × 15-20 rows
        - Large: 12-16 columns × 20-30 rows
    
    Example - Position chart from D2 to L16:
        >>> position = {
        ...     "from_col": 3,   # Column D (0-based)
        ...     "from_row": 1,   # Row 2 (0-based)
        ...     "to_col": 11,    # Column L
        ...     "to_row": 15     # Row 16
        ... }
    """
    from_col: int
    from_row: int
    to_col: int
    to_row: int

class ExcelChart(TypedDict, total=False):
    """Excel chart definition with extensive customization options.
    
    Jetxl supports six chart types with advanced styling:
    - column: Vertical bars for comparing categories
    - bar: Horizontal bars for comparing items
    - line: Trends over time or continuous data
    - pie: Proportions of a whole
    - scatter: Relationships between two variables
    - area: Filled areas showing trends
    
    Attributes:
        chart_type: Type of chart (required)
        
        # Data Range (choose one method):
        data_range: Tuple (start_row, start_col, end_row, end_col) - alternative to individual params
        start_row: First data row including header (1-based)
        start_col: First data column (0-based)
        end_row: Last data row (1-based)
        end_col: Last data column (0-based)
        
        # Chart Position (required):
        from_col: Chart start column (0-based)
        from_row: Chart start row (0-based)
        to_col: Chart end column (0-based)
        to_row: Chart end row (0-based)
        
        # Basic Configuration:
        title: Chart title
        category_col: Column for X-axis labels (0-based)
        series_names: Custom names for data series
        show_legend: Show legend (default: True)
        legend_position: Legend placement
        x_axis_title: X-axis label
        y_axis_title: Y-axis label
        
        # Chart Type Options:
        stacked: Stack series (column, bar, line, area)
        percent_stacked: Stack as 100% (column, bar, line, area)
        show_data_labels: Display values on chart elements
        chart_style: Excel chart style number (1-48)
        
        # Axis Scaling:
        axis_min: Minimum Y-axis value
        axis_max: Maximum Y-axis value
        
        # Title Styling:
        title_bold: Bold title text
        title_font_size: Title font size (hundredths of point, e.g., 1800 = 18pt)
        title_color: Title color (ARGB hex without alpha, e.g., "FF0000" for red)
        
        # Axis Title Styling:
        axis_title_bold: Bold axis titles
        axis_title_font_size: Axis title font size (hundredths of point)
        axis_title_color: Axis title color (ARGB hex without alpha)
        
        # Legend Styling:
        legend_bold: Bold legend text
        legend_font_size: Legend font size (hundredths of point)
    
    Example - Basic Column Chart:
        >>> chart = {
        ...     "chart_type": "column",
        ...     "start_row": 1, "start_col": 0,
        ...     "end_row": 12, "end_col": 3,
        ...     "from_col": 5, "from_row": 1,
        ...     "to_col": 13, "to_row": 16,
        ...     "title": "Monthly Sales",
        ...     "category_col": 0,
        ...     "x_axis_title": "Month",
        ...     "y_axis_title": "Sales ($)"
        ... }
    
    Example - Stacked Column Chart with Styling:
        >>> chart = {
        ...     "chart_type": "column",
        ...     "data_range": (0, 0, 3, 3),  # Alternative to start_row/start_col
        ...     "from_col": 5, "from_row": 0,
        ...     "to_col": 15, "to_row": 20,
        ...     "title": "Revenue Components",
        ...     "title_bold": True,
        ...     "title_font_size": 1800,
        ...     "title_color": "0070C0",
        ...     "stacked": True,
        ...     "show_data_labels": True,
        ...     "category_col": 0,
        ...     "series_names": ["Revenue", "Profit", "Expenses"],
        ...     "x_axis_title": "Quarter",
        ...     "y_axis_title": "Amount ($)",
        ...     "chart_style": 104
        ... }
    
    Example - Percent Stacked Area Chart:
        >>> chart = {
        ...     "chart_type": "area",
        ...     "data_range": (0, 0, 3, 3),
        ...     "from_col": 5, "from_row": 0,
        ...     "to_col": 15, "to_row": 20,
        ...     "title": "Market Share Over Time",
        ...     "percent_stacked": True,
        ...     "axis_min": 0.0,
        ...     "axis_max": 1.0,
        ...     "category_col": 0,
        ...     "show_data_labels": True,
        ...     "x_axis_title": "Year",
        ...     "y_axis_title": "Percentage"
        ... }
    
    Example - Styled Bar Chart with Axis Control:
        >>> chart = {
        ...     "chart_type": "bar",
        ...     "start_row": 1, "start_col": 0,
        ...     "end_row": 5, "end_col": 2,
        ...     "from_col": 5, "from_row": 0,
        ...     "to_col": 15, "to_row": 20,
        ...     "title": "Sales Performance",
        ...     "title_bold": True,
        ...     "title_font_size": 1600,
        ...     "title_color": "FF0000",
        ...     "axis_min": 0.0,
        ...     "axis_max": 10000.0,
        ...     "axis_title_bold": True,
        ...     "axis_title_font_size": 1200,
        ...     "axis_title_color": "00B050",
        ...     "legend_bold": True,
        ...     "legend_font_size": 1000,
        ...     "show_data_labels": True,
        ...     "chart_style": 26
        ... }
    
    Font Size Guide:
        Font sizes use Excel's internal unit system based on the Office Open XML (OOXML) standard.
        Values are in HUNDREDTHS OF A POINT (1 point = 100 units).
        
        Formula: excel_points = font_size_value / 100
        
        Common conversions:
        Title Sizes:
        - 800 = 8pt   (Small title)
        - 1000 = 10pt (Compact title)
        - 1200 = 12pt (Standard title)
        - 1400 = 14pt (Medium title)
        - 1600 = 16pt (Large title)
        - 1800 = 18pt (Extra large title)
        - 2400 = 24pt (Presentation title)
        - 3200 = 32pt (Header/banner)
        
        Axis Title Sizes:
        - 900 = 9pt   (Small)
        - 1000 = 10pt (Standard)
        - 1100 = 11pt (Medium)
        - 1200 = 12pt (Large)
        
        Legend Sizes:
        - 800 = 8pt   (Compact)
        - 900 = 9pt   (Small)
        - 1000 = 10pt (Standard)
        - 1100 = 11pt (Medium)
        
        Why hundredths of a point?
        - Precise control over text sizing
        - Matches Excel's internal OOXML format
        - Allows fractional sizes (e.g., 1050 = 10.5pt)
        - Same system used throughout Microsoft Office
    
    Chart Style Numbers:
        Excel provides 48 pre-defined chart styles that apply coordinated colors,
        effects, and formatting. Each chart type interprets these styles differently.
        
        Style Categories (1-48):
        - 1-10: Colorful variations with different color schemes
        - 11-16: Monochrome styles (black, white, gray)
        - 17-32: Colorful outlined styles with borders
        - 33-40: Soft color palettes
        - 41-48: Modern gradient and flat design styles
        
        Popular Styles by Use Case:
        Professional/Corporate:
        - 2: Blue colorful
        - 11: Monochrome gray
        - 26: Dark professional
        
        Modern/Vibrant:
        - 42: Gradient modern
        - 102: Contemporary (if available)
        - 104: Bright modern
        
        Print-Friendly:
        - 11: Black and white
        - 15: High contrast
        
        How to Find Your Preferred Style:
        1. Create a chart in Excel manually
        2. Click the chart → Chart Design → Chart Styles
        3. Preview different styles in the gallery
        4. Note the style number (gallery order)
        5. Use that number in Jetxl's chart_style parameter
        
        Reference:
        https://support.microsoft.com/en-us/office/available-chart-types-in-office-a6187218-807e-4103-9e0a-27cdb19afb90
    
    Axis Scaling Guide:
        Control the Y-axis range to focus on relevant data ranges or maintain
        consistent scales across charts.
        
        When to Use Axis Scaling:
        
        1. Percentage Data (0-100%):
           axis_min=0.0, axis_max=100.0
        
        2. Normalized Data (0-1) for percent_stacked:
           axis_min=0.0, axis_max=1.0
        
        3. Focus on Variance (zoom in):
           Data ranges 85-95, use axis_min=80.0, axis_max=100.0
        
        4. Compare Multiple Charts:
           Use same scale across all charts for fair comparison
        
        5. Symmetric Ranges (profit/loss):
           axis_min=-10000.0, axis_max=10000.0
        
        Best Practices:
        - Always start at 0.0 for bar/column charts (avoid misleading visuals)
        - Use axis_min > 0 only for line charts showing trends
        - For percent_stacked, always use axis_min=0.0, axis_max=1.0
        - Omit axis_min/axis_max to let Excel auto-scale
        - Use consistent scales when comparing multiple charts
    """
    # Required
    chart_type: Literal["column", "bar", "line", "pie", "scatter", "area"]
    
    # Data Range - Option 1: Tuple format
    data_range: Tuple[int, int, int, int]  # (start_row, start_col, end_row, end_col)
    
    # Data Range - Option 2: Individual parameters
    start_row: int                      # 1-based
    start_col: int                      # 0-based
    end_row: int                        # 1-based
    end_col: int                        # 0-based
    
    # Chart Position (required)
    from_col: int                       # Chart position
    from_row: int                       # Chart position
    to_col: int                         # Chart position
    to_row: int                         # Chart position
    
    # Basic Configuration
    title: str                          # Chart title
    category_col: int                   # 0-based column for categories
    series_names: List[str]             # Custom series names
    show_legend: bool                   # Show legend
    legend_position: Literal["right", "left", "top", "bottom", "none"]
    x_axis_title: str                   # X-axis label
    y_axis_title: str                   # Y-axis label
    
    # Chart Type Options
    stacked: bool                       # Stack series (column, bar, line, area)
    percent_stacked: bool               # Stack as 100% (column, bar, line, area)
    show_data_labels: bool              # Display values on chart elements
    chart_style: int                    # Excel chart style (1-48)
    
    # Axis Scaling
    axis_min: float                     # Minimum Y-axis value
    axis_max: float                     # Maximum Y-axis value
    
    # Title Styling
    title_bold: bool                    # Bold title text
    title_font_size: int                # Title font size (hundredths of point)
    title_color: str                    # Title color (ARGB hex without alpha)
    
    # Axis Title Styling
    axis_title_bold: bool               # Bold axis titles
    axis_title_font_size: int           # Axis title font size (hundredths of point)
    axis_title_color: str               # Axis title color (ARGB hex without alpha)
    
    # Legend Styling
    legend_bold: bool                   # Bold legend text
    legend_font_size: int               # Legend font size (hundredths of point)

# =============================================================================
# EXCEL IMAGES
# =============================================================================

class ImagePosition(TypedDict):
    """Image position on the worksheet.
    
    Images are positioned using Excel's column/row coordinates:
    - Columns are 0-indexed: A=0, B=1, C=2, etc.
    - Rows are 0-indexed for position: 0=row 1, 1=row 2, etc.
    
    Attributes:
        from_col: Starting column (0-based)
        from_row: Starting row (0-based)
        to_col: Ending column (0-based)
        to_row: Ending row (0-based)
    
    Size Recommendations:
        - Small: 2-4 columns × 5-8 rows
        - Medium: 4-6 columns × 8-12 rows
        - Large: 6-10 columns × 12-20 rows
    
    Example - Position image from B3 to E10:
        >>> position = {
        ...     "from_col": 1,   # Column B (0-based)
        ...     "from_row": 2,   # Row 3 (0-based)
        ...     "to_col": 4,     # Column E
        ...     "to_row": 9      # Row 10
        ... }
    """
    from_col: int
    from_row: int
    to_col: int
    to_row: int

class ExcelImageFromPath(TypedDict):
    """Excel image loaded from a file path.
    
    This is the recommended method for adding images as it's simpler
    and automatically detects the image format from the file extension.
    
    Attributes:
        path: File path to the image (required)
        from_col: Starting column (required, 0-based)
        from_row: Starting row (required, 0-based)
        to_col: Ending column (required, 0-based)
        to_row: Ending row (required, 0-based)
    
    Supported Formats:
        - PNG (.png)
        - JPEG (.jpg, .jpeg)
        - GIF (.gif)
        - BMP (.bmp)
        - TIFF (.tiff, .tif)
    
    Example - Add logo from file:
        >>> image = {
        ...     "path": "company_logo.png",
        ...     "from_col": 0,   # Column A
        ...     "from_row": 0,   # Row 1
        ...     "to_col": 2,     # Column C
        ...     "to_row": 5      # Row 6
        ... }
    """
    path: str           # Required: file path
    from_col: int       # Required: 0-based column
    from_row: int       # Required: 0-based row
    to_col: int         # Required: 0-based column
    to_row: int         # Required: 0-based row

class ExcelImageFromBytes(TypedDict):
    """Excel image from raw image data bytes.
    
    Use this method when you have image data in memory (e.g., from
    an API response, database, or generated programmatically).
    
    Attributes:
        data: Raw image bytes (required)
        extension: Image format extension (required)
        from_col: Starting column (required, 0-based)
        from_row: Starting row (required, 0-based)
        to_col: Ending column (required, 0-based)
        to_row: Ending row (required, 0-based)
    
    Supported Extensions:
        - "png"
        - "jpg" or "jpeg"
        - "gif"
        - "bmp"
        - "tiff" or "tif"
    
    Example - Add image from bytes:
        >>> import requests
        >>> response = requests.get("https://example.com/image.png")
        >>> image_bytes = response.content
        >>> 
        >>> image = {
        ...     "data": image_bytes,
        ...     "extension": "png",
        ...     "from_col": 5,
        ...     "from_row": 2,
        ...     "to_col": 8,
        ...     "to_row": 10
        ... }
    
    Example - Read from file manually:
        >>> with open("logo.png", "rb") as f:
        ...     image_data = f.read()
        >>> 
        >>> image = {
        ...     "data": image_data,
        ...     "extension": "png",
        ...     "from_col": 0,
        ...     "from_row": 0,
        ...     "to_col": 3,
        ...     "to_row": 6
        ... }
    """
    data: bytes         # Required: raw image bytes
    extension: str      # Required: image format (png, jpg, gif, etc.)
    from_col: int       # Required: 0-based column
    from_row: int       # Required: 0-based row
    to_col: int         # Required: 0-based column
    to_row: int         # Required: 0-based row

ExcelImage = Union[ExcelImageFromPath, ExcelImageFromBytes]

# =============================================================================
# MAIN FUNCTIONS
# =============================================================================

def write_sheet_arrow(
    arrow_data: Any,
    filename: str,
    sheet_name: Optional[str] = None,
    auto_filter: bool = False,
    freeze_rows: int = 0,
    freeze_cols: int = 0,
    auto_width: bool = False,
    styled_headers: bool = False,
    write_header_row: bool = True,
    column_widths: Optional[Dict[str, Union[float, str]]] = None,
    column_formats: Optional[Dict[str, str]] = None,
    merge_cells: Optional[List[Tuple[int, int, int, int]]] = None,
    data_validations: Optional[List[DataValidation]] = None,
    hyperlinks: Optional[List[Tuple[int, int, str, Optional[str]]]] = None,
    row_heights: Optional[Dict[int, float]] = None,
    cell_styles: Optional[List[CellStyleMap]] = None,
    formulas: Optional[List[Tuple[int, int, str, Optional[str]]]] = None,
    conditional_formats: Optional[List[ConditionalFormat]] = None,
    tables: Optional[List[ExcelTable]] = None,
    charts: Optional[List[ExcelChart]] = None,
    images: Optional[List[ExcelImage]] = None,
    gridlines_visible: bool = True,              
    zoom_scale: Optional[int] = None,            
    tab_color: Optional[str] = None,             
    default_row_height: Optional[float] = None,  
    hidden_columns: Optional[List[int]] = None,  
    hidden_rows: Optional[List[int]] = None,     
    right_to_left: bool = False,                 
    data_start_row: int = 0,        
    header_content: Optional[List[Tuple[int, int, str]]] = None,             
) -> None:
    """Write Arrow data to Excel with advanced formatting.
    
    This is the primary high-performance API for writing Excel files. It uses
    zero-copy Arrow integration for maximum speed and minimal memory usage.
    
    Supports:
    - Polars: df.to_arrow()
    - Pandas: df.to_arrow() (requires pyarrow)
    - PyArrow: native Table/RecordBatch
    
    Args:
        arrow_data: PyArrow Table or RecordBatch (from DataFrame.to_arrow())
        filename: Output Excel file path (.xlsx)
        sheet_name: Sheet name (default: "Sheet1")
        auto_filter: Enable autofilter dropdowns on header row
        freeze_rows: Number of rows to freeze (typically 1 for headers)
        freeze_cols: Number of columns to freeze
        auto_width: Automatically calculate column widths from content
        styled_headers: Apply bold text + gray background to headers
        write_header_row: Write column names as first row (default: True)
        column_widths: Manual column widths by name, e.g. {"Name": 20.0, "Age": 10.0}
        column_formats: Number formats by column, e.g. {"Price": "currency", "Date": "date"}
        merge_cells: List of (start_row, start_col, end_row, end_col) to merge
        data_validations: List of validation rules (dropdowns, number ranges, etc.)
        hyperlinks: List of (row, col, url, display_text) for clickable links
        row_heights: Custom row heights by row number, e.g. {1: 30.0, 5: 25.0}
        cell_styles: Custom styles with positions for individual cells
        formulas: List of (row, col, formula, cached_value) for Excel formulas
        conditional_formats: Conditional formatting rules (color scales, data bars, etc.)
        tables: Excel table definitions with filtering and styling
        charts: Excel chart definitions (column, bar, line, pie, scatter, area)
        images: Excel image definitions (from file path or bytes)
        gridlines_visible: Show worksheet gridlines (default: True)
        zoom_scale: Zoom percentage 10-400 (default: 100)
        tab_color: Sheet tab color in ARGB hex format (e.g., "FFFF0000" for red)
        default_row_height: Default height for all rows in points
        hidden_columns: List of column indices to hide (0-based)
        hidden_rows: List of row indices to hide (1-based)
        right_to_left: Enable right-to-left worksheet layout
        data_start_row: Skip this many rows when auto-calculating column widths
        header_content: Optional[List[Tuple[int, int, str]]] - Arbitrary text content to write before data rows.
        Each tuple is (row, col, text). Useful for titles, metadata, or template headers.
        These rows are written BEFORE the DataFrame data and headers.
        Example: [(1, 0, "Company Report"), (2, 0, "Q4 2024")]
    
    Examples:
        Basic Usage (Polars):
            >>> import polars as pl
            >>> import jetxl
            >>> df = pl.DataFrame({"Name": ["Alice", "Bob"], "Age": [25, 30]})
            >>> jetxl.write_sheet_arrow(df.to_arrow(), "output.xlsx")
        
        Basic Usage (Pandas):
            >>> import pandas as pd
            >>> df = pd.DataFrame({"Name": ["Alice", "Bob"], "Age": [25, 30]})
            >>> jetxl.write_sheet_arrow(df.to_arrow(), "output.xlsx")
        
        Data Only (No Headers):
            >>> jetxl.write_sheet_arrow(
            >>>    df.to_arrow(),
            >>>     "data_only.xlsx",
            >>>     write_header_row=False)  # Skip column names

        
        Formatted Output:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "formatted.xlsx",
            ...     sheet_name="Sales Data",
            ...     auto_filter=True,
            ...     freeze_rows=1,
            ...     styled_headers=True,
            ...     auto_width=True,
            ...     column_formats={
            ...         "Date": "date",
            ...         "Amount": "currency",
            ...         "Percentage": "percentage"
            ...     }
            ... )
        
        Manual Column Widths:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "sized.xlsx",
            ...     column_widths={
            ...         "Name": 25.0,
            ...         "Description": 50.0,
            ...         "Price": 12.0
            ...     },
            ...     row_heights={
            ...         1: 30.0,  # Header row
            ...         2: 25.0   # First data row
            ...     }
            ... )
        
        Hyperlinks:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "links.xlsx",
            ...     hyperlinks=[
            ...         (2, 0, "https://example.com", "Visit Website"),
            ...         (3, 0, "mailto:user@example.com", None)
            ...     ]
            ... )
        
        Data Validation Dropdown:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "validation.xlsx",
            ...     data_validations=[{
            ...         "start_row": 2, "start_col": 0,
            ...         "end_row": 100, "end_col": 0,
            ...         "type": "list",
            ...         "items": ["Pending", "Approved", "Rejected"],
            ...         "show_dropdown": True,
            ...         "error_title": "Invalid Status",
            ...         "error_message": "Please select from dropdown"
            ...     }]
            ... )
        
        Custom Cell Styles:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "styled.xlsx",
            ...     cell_styles=[{
            ...         "row": 2, "col": 0,
            ...         "font": {
            ...             "bold": True,
            ...             "size": 14.0,
            ...             "color": "FFFF0000"  # Red
            ...         },
            ...         "fill": {
            ...             "pattern": "solid",
            ...             "fg_color": "FFFFFF00"  # Yellow
            ...         },
            ...         "alignment": {
            ...             "horizontal": "center",
            ...             "vertical": "center"
            ...         }
            ...     }]
            ... )
        
        Formulas:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "formulas.xlsx",
            ...     formulas=[
            ...         (2, 3, "=SUM(A2:C2)", None),
            ...         (5, 3, "=AVERAGE(D2:D4)", "45.5")
            ...     ]
            ... )
        
        Conditional Formatting:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "conditional.xlsx",
            ...     conditional_formats=[{
            ...         "start_row": 2, "start_col": 2,
            ...         "end_row": 100, "end_col": 2,
            ...         "rule_type": "cell_value",
            ...         "operator": "greater_than",
            ...         "value": "1000",
            ...         "priority": 1,
            ...         "style": {
            ...             "font": {"bold": True, "color": "FFFF0000"}
            ...         }
            ...     }]
            ... )
        
        Color Scale:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "colors.xlsx",
            ...     conditional_formats=[{
            ...         "start_row": 2, "start_col": 3,
            ...         "end_row": 50, "end_col": 3,
            ...         "rule_type": "color_scale",
            ...         "min_color": "FFF8696B",  # Red
            ...         "mid_color": "FFFFEB84",  # Yellow
            ...         "max_color": "FF63BE7B",  # Green
            ...         "priority": 1
            ...     }]
            ... )
        
        Excel Table:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "table.xlsx",
            ...     tables=[{
            ...         "name": "SalesData",
            ...         "display_name": "Q1 Sales",
            ...         "start_row": 1, "start_col": 0,
            ...         "end_row": 100, "end_col": 5,
            ...         "style": "TableStyleMedium2",
            ...         "show_row_stripes": True
            ...     }]
            ... )
        
        Chart:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "chart.xlsx",
            ...     charts=[{
            ...         "chart_type": "column",
            ...         "start_row": 1, "start_col": 0,
            ...         "end_row": 12, "end_col": 3,
            ...         "from_col": 5, "from_row": 1,
            ...         "to_col": 13, "to_row": 16,
            ...         "title": "Monthly Sales",
            ...         "category_col": 0,
            ...         "x_axis_title": "Month",
            ...         "y_axis_title": "Revenue ($)"
            ...     }]
            ... )
        
        With Image from File:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "report.xlsx",
            ...     images=[{
            ...         "path": "logo.png",
            ...         "from_col": 0,
            ...         "from_row": 0,
            ...         "to_col": 3,
            ...         "to_row": 5
            ...     }]
            ... )
        
        With Image from Bytes:
            >>> import requests
            >>> response = requests.get("https://example.com/chart.png")
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "report.xlsx",
            ...     images=[{
            ...         "data": response.content,
            ...         "extension": "png",
            ...         "from_col": 5,
            ...         "from_row": 1,
            ...         "to_col": 12,
            ...         "to_row": 15
            ...     }]
            ... )
        
        Multiple Images:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "dashboard.xlsx",
            ...     images=[
            ...         {
            ...             "path": "logo.png",
            ...             "from_col": 0,
            ...             "from_row": 0,
            ...             "to_col": 2,
            ...             "to_row": 4
            ...         },
            ...         {
            ...             "path": "chart.png",
            ...             "from_col": 8,
            ...             "from_row": 1,
            ...             "to_col": 15,
            ...             "to_row": 20
            ...         }
            ...     ]
            ... )
        
        Images with Charts:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "visual_report.xlsx",
            ...     charts=[{
            ...         "chart_type": "column",
            ...         "start_row": 1, "start_col": 0,
            ...         "end_row": 10, "end_col": 2,
            ...         "from_col": 4, "from_row": 1,
            ...         "to_col": 12, "to_row": 15,
            ...         "title": "Sales Chart"
            ...     }],
            ...     images=[{
            ...         "path": "company_logo.png",
            ...         "from_col": 0,
            ...         "from_row": 0,
            ...         "to_col": 2,
            ...         "to_row": 4
            ...     }]
            ... )
        Sheet Appearance:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "styled_sheet.xlsx",
            ...     gridlines_visible=False,
            ...     zoom_scale=150,
            ...     tab_color="FF4472C4",  # Blue tab
            ...     default_row_height=18.0
            ... )
        
        Hidden Rows/Columns:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "hidden.xlsx",
            ...     hidden_columns=[2, 3],  # Hide columns C and D
            ...     hidden_rows=[5, 6, 7]   # Hide rows 5, 6, 7
            ... )
        
        Right-to-Left Layout:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "rtl.xlsx",
            ...     right_to_left=True  # For Hebrew, Arabic, etc.
            ... )
        
        Auto-Width with Dummy Rows:
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "report.xlsx",
            ...     auto_width=True,
            ...     data_start_row=5  # Skip first 5 rows in width calculation
            ... )
        Complete Example:
            >>> import polars as pl
            >>> import jetxl
            >>> 
            >>> df = pl.DataFrame({
            ...     "Date": ["2024-01-01", "2024-01-02", "2024-01-03"],
            ...     "Product": ["Widget A", "Widget B", "Widget C"],
            ...     "Quantity": [100, 150, 75],
            ...     "Price": [19.99, 29.99, 39.99],
            ...     "Revenue": [1999.0, 4498.5, 2999.25]
            ... })
            >>> 
            >>> jetxl.write_sheet_arrow(
            ...     df.to_arrow(),
            ...     "complete_report.xlsx",
            ...     sheet_name="Q1 Sales",
            ...     styled_headers=True,
            ...     freeze_rows=1,
            ...     auto_width=True,
            ...     column_formats={
            ...         "Date": "date",
            ...         "Price": "currency",
            ...         "Revenue": "currency"
            ...     },
            ...     tables=[{
            ...         "name": "SalesData",
            ...         "start_row": 1, "start_col": 0,
            ...         "end_row": 4, "end_col": 4,
            ...         "style": "TableStyleMedium9"
            ...     }],
            ...     conditional_formats=[{
            ...         "start_row": 2, "start_col": 4,
            ...         "end_row": 4, "end_col": 4,
            ...         "rule_type": "data_bar",
            ...         "color": "FF638EC6",
            ...         "show_value": True,
            ...         "priority": 1
            ...     }],
            ...     charts=[{
            ...         "chart_type": "column",
            ...         "start_row": 1, "start_col": 0,
            ...         "end_row": 4, "end_col": 4,
            ...         "from_col": 6, "from_row": 1,
            ...         "to_col": 14, "to_row": 18,
            ...         "title": "Revenue by Product",
            ...         "category_col": 1,
            ...         "x_axis_title": "Product",
            ...         "y_axis_title": "Revenue ($)"
            ...     }]
            ... )
    
    Raises:
        IOError: If file cannot be written or image cannot be read
        ValueError: If arrow_data is empty or invalid
    
    Notes:
        - Row numbers are 1-based (row 1 is the first row)
        - Column numbers are 0-based (column 0 is 'A')
        - Colors use ARGB hex format: "AARRGGBB"
        - Images are embedded in the Excel file
        - Supported image formats: PNG, JPEG, GIF, BMP, TIFF
        - Performance: 10-100x faster than openpyxl/xlsxwriter
        - Memory: Minimal overhead with streaming XML generation
    """
    ...

def write_sheets_arrow(
    arrow_sheets: List[Dict[str, Any]],
    filename: str,
    num_threads: int,
) -> None:
    """Write multiple Arrow tables to Excel sheets with parallel processing.
    
    Full feature parity with write_sheet_arrow() - each sheet can have completely
    independent formatting, styles, charts, tables, images, and all other options.
    
    Args:
        arrow_sheets: List of sheet configurations, each containing:
            - data: PyArrow Table/RecordBatch (required)
            - name: Sheet name (required)
            - auto_filter: Enable autofilter (optional)
            - freeze_rows: Rows to freeze (optional)
            - freeze_cols: Columns to freeze (optional)
            - auto_width: Auto-calculate column widths (optional)
            - styled_headers: Bold headers with gray background (optional)
            - write_header_row: Write column names as first row (optional)
            - column_widths: Dict[str, float|str] - manual widths (optional)
            - column_formats: Dict[str, str] - number formats (optional)
            - merge_cells: List[(row, col, row, col)] - merge ranges (optional)
            - data_validations: List[dict] - validation rules (optional)
            - hyperlinks: List[(row, col, url, display)] (optional)
            - row_heights: Dict[int, float] - row heights (optional)
            - cell_styles: List[dict] - individual cell styles (optional)
            - formulas: List[(row, col, formula, cached_value)] (optional)
            - conditional_formats: List[dict] - conditional formatting (optional)
            - tables: List[dict] - Excel table definitions (optional)
            - charts: List[dict] - chart definitions (optional)
            - images: List[dict] - image definitions (optional)
            - gridlines_visible: Show gridlines (optional)
            - zoom_scale: Zoom percentage 10-400 (optional)
            - tab_color: Sheet tab color in ARGB hex (optional)
            - default_row_height: Default row height (optional)
            - hidden_columns: List[int] - column indices to hide (optional)
            - hidden_rows: List[int] - row indices to hide (optional)
            - right_to_left: RTL layout (optional)
            - data_start_row: Skip rows for auto-width calculation (optional)
        filename: Output Excel file path (.xlsx)
        num_threads: Number of parallel threads for XML generation
    
    Examples:
        Basic Multi-Sheet:
            >>> import polars as pl
            >>> import jetxl
            >>> 
            >>> df_sales = pl.DataFrame({"Product": ["A", "B"], "Revenue": [100, 200]})
            >>> df_costs = pl.DataFrame({"Product": ["A", "B"], "Cost": [50, 80]})
            >>> 
            >>> sheets = [
            ...     {"data": df_sales.to_arrow(), "name": "Sales"},
            ...     {"data": df_costs.to_arrow(), "name": "Costs"}
            ... ]
            >>> 
            >>> jetxl.write_sheets_arrow(sheets, "report.xlsx", num_threads=2)
        
        Independent Sheet Formatting:
            >>> sheets = [
            ...     {
            ...         "data": df_sales.to_arrow(),
            ...         "name": "Sales",
            ...         "auto_filter": True,
            ...         "styled_headers": True,
            ...         "column_formats": {"Revenue": "currency"},
            ...         "tables": [{
            ...             "name": "SalesTable",
            ...             "start_row": 1, "start_col": 0,
            ...             "end_row": 100, "end_col": 2,
            ...             "style": "TableStyleMedium2"
            ...         }]
            ...     },
            ...     {
            ...         "data": df_costs.to_arrow(),
            ...         "name": "Costs",
            ...         "freeze_rows": 1,
            ...         "conditional_formats": [{
            ...             "start_row": 2, "start_col": 1,
            ...             "end_row": 50, "end_col": 1,
            ...             "rule_type": "data_bar",
            ...             "color": "FF638EC6",
            ...             "show_value": True,
            ...             "priority": 1
            ...         }]
            ...     },
            ...     {
            ...         "data": df_profit.to_arrow(),
            ...         "name": "Profit",
            ...         "auto_width": True,
            ...         "cell_styles": [{
            ...             "row": 2, "col": 0,
            ...             "font": {"bold": True, "color": "FF00B050"}
            ...         }]
            ...     }
            ... ]
            >>> 
            >>> jetxl.write_sheets_arrow(sheets, "complete.xlsx", num_threads=4)
        
        Charts, Tables, and Images Per Sheet:
            >>> sheets = [
            ...     {
            ...         "data": df_monthly.to_arrow(),
            ...         "name": "Monthly Sales",
            ...         "styled_headers": True,
            ...         "tables": [{
            ...             "name": "MonthlyData",
            ...             "start_row": 1, "start_col": 0,
            ...             "end_row": 12, "end_col": 3,
            ...             "style": "TableStyleMedium9"
            ...         }],
            ...         "charts": [{
            ...             "chart_type": "column",
            ...             "start_row": 1, "start_col": 0,
            ...             "end_row": 12, "end_col": 2,
            ...             "from_col": 5, "from_row": 1,
            ...             "to_col": 13, "to_row": 16,
            ...             "title": "Sales Trend",
            ...             "category_col": 0
            ...         }],
            ...         "images": [{
            ...             "path": "logo.png",
            ...             "from_col": 0, "from_row": 0,
            ...             "to_col": 2, "to_row": 4
            ...         }]
            ...     },
            ...     {
            ...         "data": df_quarterly.to_arrow(),
            ...         "name": "Quarterly Summary",
            ...         "auto_filter": True,
            ...         "charts": [{
            ...             "chart_type": "pie",
            ...             "start_row": 1, "start_col": 0,
            ...             "end_row": 4, "end_col": 1,
            ...             "from_col": 3, "from_row": 1,
            ...             "to_col": 10, "to_row": 15,
            ...             "title": "Market Share"
            ...         }]
            ...     }
            ... ]
            >>> 
            >>> jetxl.write_sheets_arrow(sheets, "dashboard.xlsx", num_threads=2)
        
        Advanced Formatting Per Sheet:
            >>> sheets = [
            ...     {
            ...         "data": df1.to_arrow(),
            ...         "name": "Formatted",
            ...         "auto_width": True,
            ...         "styled_headers": True,
            ...         "freeze_rows": 1,
            ...         "column_formats": {
            ...             "Date": "date",
            ...             "Amount": "currency",
            ...             "Rate": "percentage"
            ...         },
            ...         "merge_cells": [(1, 0, 1, 3)],
            ...         "data_validations": [{
            ...             "start_row": 2, "start_col": 4,
            ...             "end_row": 100, "end_col": 4,
            ...             "type": "list",
            ...             "items": ["Active", "Inactive"],
            ...             "show_dropdown": True
            ...         }],
            ...         "hyperlinks": [(2, 0, "https://example.com", "Link")],
            ...         "formulas": [(5, 5, "=SUM(A2:A4)", None)],
            ...         "gridlines_visible": False,
            ...         "zoom_scale": 120,
            ...         "tab_color": "FF00B050"
            ...     },
            ...     {
            ...         "data": df2.to_arrow(),
            ...         "name": "Raw Data",
            ...         "write_header_row": False,  # Data only, no headers
            ...         "hidden_columns": [2, 3],
            ...         "default_row_height": 20.0
            ...     }
            ... ]
            >>> 
            >>> jetxl.write_sheets_arrow(sheets, "advanced.xlsx", num_threads=2)
    
    Raises:
        IOError: If file cannot be written or images cannot be read
        ValueError: If any sheet data is invalid
        KeyError: If required keys (data, name) are missing
    
    Notes:
        - Full feature parity with write_sheet_arrow()
        - Each sheet has completely independent configuration
        - Style registry is shared across sheets for deduplication
        - Parallel XML generation across num_threads
        - Performance: minimal overhead vs single-sheet (<1%)
    """
    ...

def write_sheet_arrow_to_bytes(
    arrow_data: Any,
    sheet_name: Optional[str] = None,
    auto_filter: bool = False,
    freeze_rows: int = 0,
    freeze_cols: int = 0,
    auto_width: bool = False,
    styled_headers: bool = False,
    write_header_row: bool = True,
    column_widths: Optional[Dict[str, Union[float, int, str]]] = None,
    column_formats: Optional[Dict[str, str]] = None,
    merge_cells: Optional[List[Tuple[int, int, int, int]]] = None,
    data_validations: Optional[List[DataValidation]] = None,
    hyperlinks: Optional[List[Tuple[int, int, str, Optional[str]]]] = None,
    row_heights: Optional[Dict[int, float]] = None,
    cell_styles: Optional[List[CellStyleDict]] = None,
    formulas: Optional[List[Tuple[int, int, str, Optional[str]]]] = None,
    conditional_formats: Optional[List[ConditionalFormat]] = None,
    tables: Optional[List[ExcelTable]] = None,
    charts: Optional[List[ExcelChart]] = None,
    images: Optional[List[ExcelImage]] = None,
    gridlines_visible: bool = True,
    zoom_scale: Optional[int] = None,
    tab_color: Optional[str] = None,
    default_row_height: Optional[float] = None,
    hidden_columns: Optional[List[int]] = None,
    hidden_rows: Optional[List[int]] = None,
    right_to_left: bool = False,
    data_start_row: int = 0,
    header_content: Optional[List[Tuple[int, int, str]]] = None,
) -> bytes:
    """Write Arrow data to Excel bytes (in-memory, no file I/O).
    
    Identical to write_sheet_arrow() but returns bytes instead of writing to file.
    Perfect for web APIs, Lambda, streaming scenarios.
    
    Returns:
        bytes: Complete Excel file as bytes (.xlsx format)
   
    Examples:
        Flask endpoint:
            >>> from flask import Response
            >>> 
            >>> @app.route('/export')
            >>> def export():
            ...     df = get_sales_data()
            ...     excel = jet.write_sheet_arrow_to_bytes(
            ...         df.to_arrow(),
            ...         styled_headers=True,
            ...         auto_width=True
            ...     )
            ...     return Response(
            ...         excel,
            ...         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            ...         headers={'Content-Disposition': 'attachment;filename=sales.xlsx'}
            ...     )
        
        FastAPI with streaming:
            >>> from fastapi.responses import StreamingResponse
            >>> from io import BytesIO
            >>> 
            >>> @app.get("/download")
            >>> async def download():
            ...     excel = jet.write_sheet_arrow_to_bytes(df.to_arrow())
            ...     return StreamingResponse(
            ...         BytesIO(excel),
            ...         media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            ...         headers={'Content-Disposition': 'attachment; filename=report.xlsx'}
            ...     )
        
        AWS Lambda:
            >>> import base64
            >>> 
            >>> def lambda_handler(event, context):
            ...     df = query_database()
            ...     excel = jet.write_sheet_arrow_to_bytes(df.to_arrow())
            ...     return {
            ...         'statusCode': 200,
            ...         'body': base64.b64encode(excel).decode('utf-8'),
            ...         'isBase64Encoded': True,
            ...         'headers': {
            ...             'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            ...             'Content-Disposition': 'attachment; filename=data.xlsx'
            ...         }
            ...     }
        
        S3 upload without temp file:
            >>> import boto3
            >>> 
            >>> s3 = boto3.client('s3')
            >>> excel = jet.write_sheet_arrow_to_bytes(
            ...     df.to_arrow(),
            ...     sheet_name="Report"
            ... )
            >>> s3.put_object(
            ...     Bucket='reports',
            ...     Key='2024/sales.xlsx',
            ...     Body=excel,
            ...     ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            ... )
        
        JSON API response:
            >>> import base64
            >>> 
            >>> excel = jet.write_sheet_arrow_to_bytes(df.to_arrow())
            >>> return jsonify({
            ...     'filename': 'report.xlsx',
            ...     'data': base64.b64encode(excel).decode('utf-8'),
            ...     'size': len(excel)
            ... })
        
        Email attachment:
            >>> from email.mime.application import MIMEApplication
            >>> 
            >>> excel = jet.write_sheet_arrow_to_bytes(df.to_arrow())
            >>> attachment = MIMEApplication(excel)
            >>> attachment.add_header('Content-Disposition', 'attachment', filename='report.xlsx')
            >>> msg.attach(attachment)
        
        Django view:
            >>> from django.http import HttpResponse
            >>> 
            >>> def export_view(request):
            ...     excel = jet.write_sheet_arrow_to_bytes(df.to_arrow())
            ...     response = HttpResponse(
            ...         excel,
            ...         content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            ...     )
            ...     response['Content-Disposition'] = 'attachment; filename=export.xlsx'
            ...     return response
        
        Azure Functions:
            >>> import azure.functions as func
            >>> 
            >>> def main(req: func.HttpRequest) -> func.HttpResponse:
            ...     excel = jet.write_sheet_arrow_to_bytes(df.to_arrow())
            ...     return func.HttpResponse(
            ...         excel,
            ...         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            ...         headers={'Content-Disposition': 'attachment; filename=data.xlsx'}
            ...     )
            
            """
    ... 

def write_sheets_arrow_to_bytes(
    sheets_data: List[SheetConfig],
    num_threads: int = 1,
) -> bytes:
    """Write multiple sheets to Excel bytes (in-memory, no file I/O).
    
    Identical to write_sheets_arrow() but returns bytes instead of writing to file.
    
    Returns:
        bytes: Complete Excel file as bytes (.xlsx format)
    

    Examples:
        Multi-sheet API endpoint:
            >>> @app.route('/quarterly-report')
            >>> def quarterly():
            ...     sheets = [
            ...         {"data": q1_df.to_arrow(), "name": "Q1", "styled_headers": True},
            ...         {"data": q2_df.to_arrow(), "name": "Q2", "styled_headers": True},
            ...         {"data": q3_df.to_arrow(), "name": "Q3", "styled_headers": True},
            ...         {"data": q4_df.to_arrow(), "name": "Q4", "styled_headers": True}
            ...     ]
            ...     excel = jet.write_sheets_arrow_to_bytes(sheets, num_threads=4)
            ...     return Response(excel, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        Google Cloud Storage:
            >>> from google.cloud import storage
            >>> 
            >>> client = storage.Client()
            >>> bucket = client.bucket('reports')
            >>> blob = bucket.blob('monthly/report.xlsx')
            >>> 
            >>> excel = jet.write_sheets_arrow_to_bytes(sheets, num_threads=2)
            >>> blob.upload_from_string(
            ...     excel,
            ...     content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            ... )
        
        Celery background task:
            >>> @celery.task
            >>> def generate_report(user_id):
            ...     sheets = fetch_user_data(user_id)
            ...     excel = jet.write_sheets_arrow_to_bytes(sheets, num_threads=4)
            ...     
            ...     # Email it
            ...     send_email(
            ...         to=user.email,
            ...         subject='Your Report',
            ...         attachments=[('report.xlsx', excel)]
            ...     )
            """

    ...


def write_sheet(
    columns: Dict[str, List[Any]],
    filename: str,
    sheet_name: Optional[str] = None,
    charts: Optional[List[ExcelChart]] = None,
) -> None:
    """Write dict-based data to Excel (legacy API, slower than Arrow).
    
    This is the legacy dictionary-based API maintained for backward compatibility.
    For new code, prefer write_sheet_arrow() which is 10-100x faster.
    
    Args:
        columns: Dictionary mapping column names to lists of values
        filename: Output Excel file path (.xlsx)
        sheet_name: Sheet name (default: "Sheet1")
        charts: List of chart definitions (optional)
    
    Supported Value Types:
        - str: Text values
        - int/float: Numeric values
        - bool: Boolean values
        - datetime: Date/time values
        - None: Empty cells
    
    Examples:
        Basic Usage:
            >>> import jetxl
            >>> 
            >>> data = {
            ...     "Name": ["Alice", "Bob", "Charlie"],
            ...     "Age": [25, 30, 35],
            ...     "Salary": [50000.0, 60000.0, 75000.0]
            ... }
            >>> 
            >>> jetxl.write_sheet(data, "output.xlsx")
        
        With Custom Sheet Name:
            >>> jetxl.write_sheet(
            ...     data,
            ...     "employees.xlsx",
            ...     sheet_name="Employee Data"
            ... )
        
        With Chart:
            >>> data = {
            ...     "Month": ["Jan", "Feb", "Mar", "Apr"],
            ...     "Sales": [1000, 1200, 1100, 1300]
            ... }
            >>> 
            >>> charts = [{
            ...     "chart_type": "column",
            ...     "start_row": 1, "start_col": 0,
            ...     "end_row": 5, "end_col": 1,
            ...     "from_col": 3, "from_row": 1,
            ...     "to_col": 10, "to_row": 12,
            ...     "title": "Monthly Sales",
            ...     "category_col": 0
            ... }]
            >>> 
            >>> jetxl.write_sheet(data, "sales.xlsx", charts=charts)
        
        With DateTime Values:
            >>> from datetime import datetime
            >>> 
            >>> data = {
            ...     "Date": [
            ...         datetime(2024, 1, 1),
            ...         datetime(2024, 1, 2),
            ...         datetime(2024, 1, 3)
            ...     ],
            ...     "Value": [100, 150, 125]
            ... }
            >>> 
            >>> jetxl.write_sheet(data, "dates.xlsx")
    
    Raises:
        IOError: If file cannot be written
        ValueError: If columns have different lengths
        TypeError: If unsupported value types are provided
    
    Notes:
        - Images are not supported in the legacy dict API
        - Use write_sheet_arrow() for image support
        - All column lists must have the same length
        - Sheet names limited to 31 characters
        - Invalid characters in sheet names: [ ] : * ? / \\
        - For better performance, use write_sheet_arrow() instead
    """
    ...

def write_sheets(
    sheets_data: List[Dict[str, Any]],
    filename: str,
    num_threads: int,
) -> None:
    """Write multiple dict-based sheets to Excel (legacy API).
    
    This is the legacy multi-sheet dictionary-based API maintained for
    backward compatibility. For new code, prefer write_sheets_arrow().
    
    Args:
        sheets_data: List of dicts with "name" and "columns" keys
        filename: Output Excel file path (.xlsx)
        num_threads: Number of parallel threads for generation
    
    Examples:
        Basic Multi-Sheet:
            >>> import jetxl
            >>> 
            >>> sheets = [
            ...     {
            ...         "name": "Sales",
            ...         "columns": {
            ...             "Product": ["A", "B", "C"],
            ...             "Revenue": [100, 200, 150]
            ...         }
            ...     },
            ...     {
            ...         "name": "Costs",
            ...         "columns": {
            ...             "Product": ["A", "B", "C"],
            ...             "Cost": [50, 80, 60]
            ...         }
            ...     }
            ... ]
            >>> 
            >>> jetxl.write_sheets(sheets, "report.xlsx", num_threads=2)
        
        With Multiple Data Types:
            >>> from datetime import datetime
            >>> 
            >>> sheets = [
            ...     {
            ...         "name": "January",
            ...         "columns": {
            ...             "Date": [datetime(2024, 1, i) for i in range(1, 8)],
            ...             "Sales": [100, 120, 110, 130, 125, 140, 135],
            ...             "Active": [True, True, True, False, True, True, True]
            ...         }
            ...     },
            ...     {
            ...         "name": "February",
            ...         "columns": {
            ...             "Date": [datetime(2024, 2, i) for i in range(1, 8)],
            ...             "Sales": [150, 160, 155, 170, 165, 180, 175],
            ...             "Active": [True] * 7
            ...         }
            ...     }
            ... ]
            >>> 
            >>> jetxl.write_sheets(sheets, "monthly.xlsx", num_threads=2)
    
    Raises:
        IOError: If file cannot be written
        ValueError: If any sheet data is invalid
        KeyError: If required keys are missing
    
    Notes:
        - Images are not supported in the legacy dict API
        - Use write_sheets_arrow() for image support
        - Each sheet dict must have "name" and "columns" keys
        - All columns in a sheet must have the same length
        - For better performance, use write_sheets_arrow() instead
    """
    ...