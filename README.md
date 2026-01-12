# Jetxl ✈️
**Blazingly fast Excel (XLSX) writer for Python, powered by Rust**

Jetxl is a high-performance library for creating Excel files from Python with native support for Arrow, Polars, and Pandas DataFrames. Built from the ground up in Rust for maximum speed and efficiency.

## ✨ Features

- 🚀 **Ultra-fast**: 5-40x faster than other Python Excel libraries
- 🔄 **Zero-copy Arrow integration**: Direct DataFrame → Excel with no intermediate conversions
- 🎨 **Rich formatting**: Fonts, colors, borders, alignment, number formats
- 📊 **Advanced features**: Conditional formatting, data validation, formulas, hyperlinks, Excel tables, charts, images
- 🧵 **Multi-threaded**: Parallel sheet generation for multi-sheet workbooks
- 💾 **Memory efficient**: Streaming XML generation with minimal memory overhead
- 🐻‍❄️🐼 **Framework agnostic**: Works seamlessly with Polars, Pandas, PyArrow, and native Python dicts



## ⚡ Performance Comparison

*Benchmark environment: Python 3.13, AMD Ryzen 9 7900x, 64 GB RAM*

At the time of the test, the following library versions were used:
- Jetxl: 0.1.0
- Polars: 1.37.0
- Pandas: 2.3.3
- Pyexcelerate: 0.13.0
- Rustpy-xlsxwriter: 0.7.0
- Openpyxl: 3.1.5
- Xlsxwriter: 3.2.9


### Library comparison summary
| Library | 1M Rows | Speedup | Throughput | Memory |
|---------|---------|---------|------------|--------|
| **jetxl (arrow)** | **2.06s** | **1.0x** | **526K rows/s** | **~0 MB** |
| jetxl (dict) | 3.57s | 1.7x slower | 286K rows/s | ~0 MB |
| xlsxwriter | 10.05s | 4.9x slower | 96K rows/s | 0.4 MB |
| rustpy_xlsxwriter | 11.27s | 5.5x slower | 89K rows/s | - |
| pyexcelerate | 35.55s | 17x slower | 28K rows/s | - |
| polars.write_excel | 40.85s | **20x slower** | 27K rows/s | 2.1 GB |
| openpyxl | 56.25s | **27x slower** | 18K rows/s | 0.4 MB |
| pandas+xlsxwriter | 56.30s | 27x slower | 18K rows/s | - |
| pandas+openpyxl | 83.42s | **40x slower** | 12K rows/s | - |



The chart below shows the execution time comparison with popular libraries for a single sheet file against Jetxl. Performance will vary on different hardware.
![exec_time_comparison](benchmark\execution_time.png)


### Execution Time (seconds)

| Library | 10K rows | 100K rows | 1M rows |
|---------|----------|-----------|---------|
| **jetxl (arrow)** | **0.022** | **0.19** | **2.06** |
| jetxl (dict) | 0.032 | 0.31 | 3.57 |
| rustpy_xlsxwriter | 0.107 | 1.06 | 11.27 |
| xlsxwriter | 0.112 | 1.00 | 10.05 |
| pyexcelerate | 0.333 | 3.86 | 35.55 |
| polars.write_excel | 0.349 | 3.53 | 40.85 |
| openpyxl | 0.516 | 5.10 | 56.25 |
| pandas+xlsxwriter | 0.523 | 5.59 | 56.30 |
| pandas+openpyxl | 0.756 | 8.33 | 83.42 |




## 📦 Installation

```bash
pip install jetxl

# Install with uv (recommended)
# uv pip install jetxl
```

## 🚀 Quick Start


> [!IMPORTANT]
> Jetxl is an **experimental xlsx writer**. There are still
> bugs or functionality that is being worked on. Existing 
> functionality is subject to change.



### Using Polars (Recommended)

```python
import polars as pl
import jetxl as jet

# Create a DataFrame
df = pl.DataFrame({
    "Name": ["Alice", "Bob", "Charlie"],
    "Age": [25, 30, 35],
    "Salary": [50000.0, 60000.0, 75000.0]
})

# Write to Excel (requires to_arrow() conversion)
jet.write_sheet_arrow(df.to_arrow(), "output.xlsx")
```

### Using Pandas

```python
import pandas as pd
import jetxl as jet

df = pd.DataFrame({
    "Name": ["Alice", "Bob", "Charlie"],
    "Age": [25, 30, 35],
    "Salary": [50000.0, 60000.0, 75000.0]
})

# Convert to Arrow for zero-copy performance
jet.write_sheet_arrow(df.to_arrow(), "output.xlsx")
```

### Using PyArrow

```python
import pyarrow as pa
import jetxl as jet

# Create an Arrow table
table = pa.table({
    "Name": ["Alice", "Bob", "Charlie"],
    "Age": [25, 30, 35],
    "Salary": [50000.0, 60000.0, 75000.0]
})

# Write directly from Arrow table
jet.write_sheet_arrow(table, "output.xlsx")
```

### Using Python Dicts (Legacy API)

```python
import jetxl as jet

data = {
    "Name": ["Alice", "Bob", "Charlie"],
    "Age": [25, 30, 35],
    "Salary": [50000.0, 60000.0, 75000.0]
}

jet.write_sheet(data, "output.xlsx")
```

## 📚 API Reference

### Arrow API (Recommended - Fastest)

#### `write_sheet_arrow()`

Write a single sheet from Arrow-compatible data (Polars, PyArrow, Pandas).

```python
jet.write_sheet_arrow(
    arrow_data,                    # DataFrame or Arrow RecordBatch
    filename,                       # Output file path
    sheet_name=None,               # Sheet name (default: "Sheet1")
    auto_filter=False,             # Enable autofilter on headers
    freeze_rows=0,                 # Number of rows to freeze
    freeze_cols=0,                 # Number of columns to freeze
    auto_width=False,              # Auto-calculate column widths
    styled_headers=False,          # Apply bold styling to headers
    write_header_row=True,         # Write column names as first row
    column_widths=None,            # Dict[str, float|str] - manual widths
    column_formats=None,           # Dict[str, str] - number formats
    merge_cells=None,              # List[(row, col, row, col)] - merge ranges
    data_validations=None,         # List[dict] - validation rules
    hyperlinks=None,               # List[(row, col, url, display)]
    row_heights=None,              # Dict[int, float] - row heights
    cell_styles=None,              # List[dict] - individual cell styles
    formulas=None,                 # List[(row, col, formula, cached_value)]
    conditional_formats=None,      # List[dict] - conditional formatting
    tables=None,                   # List[dict] - Excel table definitions
    charts=None,                   # List[dict] - Excel chart definitions
    images=None,                   # List[dict] - Excel image definitions
    gridlines_visible=True,        # Show worksheet gridlines
    zoom_scale=None,               # Zoom percentage 10-400
    tab_color=None,                # Sheet tab color (ARGB hex)
    default_row_height=None,       # Default row height in points
    hidden_columns=None,           # List[int] - column indices to hide
    hidden_rows=None,              # List[int] - row indices to hide
    right_to_left=False,           # Enable RTL layout
    data_start_row=0,              # Skip rows for auto-width calculation
    header_content=None            # List[(row, col, text)] - custom header rows
)
```

#### `write_sheets_arrow()`

Write multiple sheets with parallel processing. **Full feature parity** with `write_sheet_arrow()` - each sheet supports all formatting options independently.
```python
jet.write_sheets_arrow(
    sheets,                         # List[dict] with data, name, and any formatting options
    filename,                       # Output file path
    num_threads                     # Parallel threads for XML generation
)
```

**Each sheet dict supports all `write_sheet_arrow()` parameters:**
```python
{
    "data": arrow_data,                         # Required: Arrow Table/RecordBatch
    "name": "Sheet1",                           # Required: Sheet name
    
    # All write_sheet_arrow() options available:
    "auto_filter": bool,
    "freeze_rows": int,
    "freeze_cols": int,
    "auto_width": bool,
    "styled_headers": bool,
    "write_header_row": bool,
    "column_widths": Dict[str, float|str],
    "column_formats": Dict[str, str],
    "merge_cells": List[Tuple[int, int, int, int]],
    "data_validations": List[dict],
    "hyperlinks": List[Tuple[int, int, str, str]],
    "row_heights": Dict[int, float],
    "cell_styles": List[dict],
    "formulas": List[Tuple[int, int, str, str]],
    "conditional_formats": List[dict],
    "tables": List[dict],
    "charts": List[dict],
    "images": List[dict],
    "gridlines_visible": bool,
    "zoom_scale": int,
    "tab_color": str,
    "default_row_height": float,
    "hidden_columns": List[int],
    "hidden_rows": List[int],
    "right_to_left": bool,
    "data_start_row": int,
    "header_content": List[Tuple[int, int, str]]
}
```

**Example with independent sheet configurations:**
```python
sheets = [
    {
        "data": df_sales.to_arrow(),
        "name": "Sales",
        "styled_headers": True,
        "tables": [{"name": "SalesTable", ...}],
        "charts": [{"chart_type": "column", ...}],
        "tab_color": "FF00B050"
    },
    {
        "data": df_costs.to_arrow(),
        "name": "Costs",
        "conditional_formats": [{...}],
        "hidden_columns": [2, 3],
        "tab_color": "FFFF0000"
    }
]

jet.write_sheets_arrow(sheets, "report.xlsx", num_threads=4)
```


### In-Memory Bytes API (No File I/O)

#### `write_sheet_arrow_to_bytes()`

Returns Excel file as bytes instead of writing to disk. Identical parameters to `write_sheet_arrow()` except returns bytes instead of writing to a file.

```python
excel_bytes = jet.write_sheet_arrow_to_bytes(
    arrow_data,                    # DataFrame or Arrow RecordBatch
    sheet_name=None,               # Sheet name (default: "Sheet1")
    auto_filter=False,             # Enable autofilter on headers
    freeze_rows=0,                 # Number of rows to freeze
    freeze_cols=0,                 # Number of columns to freeze
    auto_width=False,              # Auto-calculate column widths
    styled_headers=False,          # Apply bold styling to headers
    write_header_row=True,         # Write column names as first row
    column_widths=None,            # Dict[str, float|str] - manual widths
    column_formats=None,           # Dict[str, str] - number formats
    merge_cells=None,              # List[(row, col, row, col)] - merge ranges
    data_validations=None,         # List[dict] - validation rules
    hyperlinks=None,               # List[(row, col, url, display)]
    row_heights=None,              # Dict[int, float] - row heights
    cell_styles=None,              # List[dict] - individual cell styles
    formulas=None,                 # List[(row, col, formula, cached_value)]
    conditional_formats=None,      # List[dict] - conditional formatting
    tables=None,                   # List[dict] - Excel table definitions
    charts=None,                   # List[dict] - Excel chart definitions
    images=None,                   # List[dict] - Excel image definitions
    gridlines_visible=True,        # Show worksheet gridlines
    zoom_scale=None,               # Zoom percentage 10-400
    tab_color=None,                # Sheet tab color (ARGB hex)
    default_row_height=None,       # Default row height in points
    hidden_columns=None,           # List[int] - column indices to hide
    hidden_rows=None,              # List[int] - row indices to hide
    right_to_left=False,           # Enable RTL layout
    data_start_row=0,              # Skip rows for auto-width calculation
    header_content=None            # List[(row, col, text)] - custom header rows
)
```

**Use Cases:**
- Web APIs and HTTP responses
- Cloud functions (AWS Lambda, Google Cloud Functions)
- Streaming scenarios
- In-memory processing
- Base64 encoding for JSON APIs

**Examples:**

```python
import polars as pl
import jetxl as jet

df = pl.DataFrame({
    "Name": ["Alice", "Bob"],
    "Age": [25, 30],
    "Salary": [50000, 60000]
})

# Generate Excel in memory
excel_bytes = jet.write_sheet_arrow_to_bytes(
    df.to_arrow(),
    sheet_name="Employees",
    styled_headers=True,
    auto_width=True
)

# Save to file
with open("output.xlsx", "wb") as f:
    f.write(excel_bytes)

# Or use in web framework (Flask)
from flask import Response

@app.route('/download')
def download():
    excel_bytes = jet.write_sheet_arrow_to_bytes(df.to_arrow())
    return Response(
        excel_bytes,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': 'attachment;filename=data.xlsx'}
    )

# Or base64 encode for API
import base64
encoded = base64.b64encode(excel_bytes).decode('utf-8')

# Or return from Lambda
def lambda_handler(event, context):
    excel_bytes = jet.write_sheet_arrow_to_bytes(df.to_arrow())
    return {
        'statusCode': 200,
        'body': base64.b64encode(excel_bytes).decode('utf-8'),
        'isBase64Encoded': True,
        'headers': {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
    }
```

#### `write_sheets_arrow_to_bytes()`

Write multiple sheets to bytes. Identical to `write_sheets_arrow()` but returns bytes.

```python
excel_bytes = jet.write_sheets_arrow_to_bytes(
    sheets,        # List[dict] with data, name, and any formatting options
    num_threads=1  # Parallel threads for XML generation
)
```

**Example:**

```python
sheets = [
    {
        "data": df1.to_arrow(),
        "name": "Sales",
        "styled_headers": True,
        "freeze_rows": 1
    },
    {
        "data": df2.to_arrow(),
        "name": "Costs",
        "auto_width": True
    }
]

# Generate multi-sheet Excel in memory
excel_bytes = jet.write_sheets_arrow_to_bytes(sheets, num_threads=2)

# FastAPI example
from fastapi.responses import Response

@app.get("/report")
async def generate_report():
    excel_bytes = jet.write_sheets_arrow_to_bytes(sheets, num_threads=2)
    return Response(
        content=excel_bytes,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': 'attachment; filename=report.xlsx'}
    )

# S3 upload without local file
import boto3
s3 = boto3.client('s3')
s3.put_object(
    Bucket='my-bucket',
    Key='reports/monthly.xlsx',
    Body=excel_bytes,
    ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)
```

### Dict API (Legacy - Backward Compatible)

#### `write_sheet()`

```python
jet.write_sheet(
    columns,       # Dict[str, List] - column name to values
    filename,      # Output file path
    sheet_name=None,  # Sheet name
    charts=None    # List[dict] - Excel chart definitions
)
```

#### `write_sheets()`

```python
sheets = [
    {"name": "Sales", "columns": sales_data},
    {"name": "Expenses", "columns": expenses_data}
]
jet.write_sheets(sheets, "output.xlsx", num_threads=4)
```

## 🎨 Formatting & Styling

### Basic Formatting

```python
import polars as pl
import jetxl as jet

df = pl.DataFrame({
    "Product": ["Apple", "Banana", "Cherry"],
    "Price": [1.50, 0.75, 2.25],
    "Quantity": [100, 150, 80]
})

jet.write_sheet_arrow(
    df.to_arrow(),
    "formatted.xlsx",
    auto_filter=True,           # Add filter dropdowns
    freeze_rows=1,              # Freeze header row
    styled_headers=True,        # Bold headers
    auto_width=True             # Auto-size columns
)

# Without headers (data only)
jet.write_sheet_arrow(
    df.to_arrow(),
    "no_headers.xlsx",
    write_header_row=False  # Skip writing column names
)

```

### Column Formats

Jetxl supports both built-in format shortcuts and custom Excel format codes for complete control over number display.

#### Built-in Format Shortcuts

```python
jet.write_sheet_arrow(
    df.to_arrow(),
    "formatted.xlsx",
    column_formats={
        "Price": "currency",           # $#,##0.00
        "Quantity": "integer",         # 0
        "Growth": "percentage",        # 0%
        "Timestamp": "datetime",       # yyyy-mm-dd hh:mm:ss
        "Score": "decimal2",           # 0.00
        "Rate": "scientific",          # 0.00E+00
        "Measurement": "fraction"      # # ?/?
    }
)
```

**Available built-in formats:**
- `general` - Default formatting
- `integer` - Whole numbers (0)
- `decimal2` - Two decimal places (0.00)
- `decimal4` - Four decimal places (0.0000)
- `percentage` - Percentage (0%)
- `percentage_decimal` - Percentage with decimal (0.00%)
- `percentage_integer` - Percentage as integer (0%)
- `currency` - Currency ($#,##0.00)
- `currency_rounded` - Rounded currency ($#,##0)
- `date` - Date (yyyy-mm-dd)
- `datetime` - Date and time (yyyy-mm-dd hh:mm:ss)
- `time` - Time (hh:mm:ss)
- `scientific` - Scientific notation (0.00E+00)
- `fraction` - Fraction (# ?/?)
- `fraction_two_digits` - Fraction with 2 digits (# ??/??)
- `thousands` - Thousands separator (#,##0)

#### Custom Format Codes

Any string not matching a built-in format becomes a custom Excel format code, giving you full control:

```python
column_formats = {
    # Accounting format with negative in parentheses
    "Amount": "$#,##0.00_);[Red]($#,##0.00)",
    
    # Thousands with 'K' suffix
    "Visitors": "#,##0,\"K\"",
    
    # Millions with 'M' suffix  
    "Revenue": "$#,##0.0,,\"M\"",
    
    # Custom date format
    "Date": "dddd, mmmm dd, yyyy",
    
    # Conditional coloring
    "Change": "[Green]#,##0;[Red]-#,##0;[Blue]0",
    
    # Fractions in sixteenths
    "Measurement": "# ?/16",
    
    # Phone numbers
    "Phone": "(###) ###-####",
    
    # Zero-padded IDs
    "ID": "00000",
    
    # Hide zeros
    "Optional": "#,##0;-#,##0;\"\""
}
```

#### Custom Format Syntax

Excel format codes use this structure:
```
[Positive];[Negative];[Zero];[Text]
```

**Format symbols:**
- `0` - Digit placeholder (shows 0 if no digit)
- `#` - Digit placeholder (shows nothing if no digit)  
- `?` - Digit placeholder (adds space for alignment)
- `.` - Decimal point
- `,` - Thousands separator (or divider in millions/thousands)
- `%` - Multiply by 100 and show percent sign
- `E+` `E-` - Scientific notation
- `"text"` - Literal text in quotes
- `@` - Text placeholder
- `[Color]` - Color code (Red, Blue, Green, etc.)
- `[>=100]` - Conditional formatting

**Scaling numbers:**
- One comma `,` after number divides by 1,000
- Two commas `,,` divide by 1,000,000
- Example: `#,##0,` shows 1500 as "2" (rounded thousands)
- Example: `#,##0.0,,` shows 5000000 as "5.0" (millions)

#### Complete Custom Format Examples

```python
import polars as pl
import jetxl as jet

df = pl.DataFrame({
    "Revenue": [1500000, 500000, 75000],
    "Change": [150, -75, 0],
    "Ratio": [0.333, 0.125, 0.875],
    "Code": [1, 42, 999],
    "Date": ["2024-01-15", "2024-02-20", "2024-03-25"]
})

jet.write_sheet_arrow(
    df.to_arrow(),
    "custom_formats.xlsx",
    column_formats={
        # Show millions with conditional formatting
        "Revenue": "[>=1000000]$#,##0.0,,\"M\";[>=1000]$#,##0,\"K\";$#,##0",
        
        # Color-coded changes with +/- indicators
        "Change": "[Green]+#,##0;[Red]-#,##0;[Blue]0",
        
        # Fractions with fallback
        "Ratio": "# ?/?;-# ?/?;\"N/A\"",
        
        # Zero-padded codes
        "Code": "000000",
        
        # Custom date format
        "Date": "dddd, mmmm dd, yyyy"
    }
)
```

#### Testing Custom Formats

The easiest way to create custom formats:

1. Open Excel and format a cell manually
2. Right-click → Format Cells → Custom
3. Copy the format code from the "Type:" field
4. Use that exact string in Jetxl

#### Limitations

- **No validation**: Custom format codes aren't validated client-side. Invalid codes may cause Excel errors when opening the file.
- **XML escaping**: Special characters (`<`, `>`, `&`, `"`, `'`) are automatically escaped - you don't need to worry about them.
- **Length limit**: Format codes are limited to ~255 characters (Excel limitation).
- **Compatibility**: Some advanced features (locale codes, DBNum) may not work in all Excel versions.
- **Color names**: Limited to Excel's built-in set: `[Red]`, `[Blue]`, `[Green]`, `[Yellow]`, `[Cyan]`, `[Magenta]`, `[White]`, `[Black]`, `[Color1]`-`[Color56]`.

**Reference**: [Excel Number Format Codes - Microsoft](https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68)

### Advanced Number Format Examples

#### Dynamic Scaling

Automatically scale numbers based on magnitude:

```python
# Show millions, thousands, or regular numbers
column_formats = {
    "Value": "[>=1000000]#,##0.0,,\"M\";[>=1000]#,##0.0,\"K\";#,##0"
}
# 5000000 → "5.0M"
# 15000 → "15.0K"  
# 500 → "500"
```

#### Conditional Text

Display custom text based on values:

```python
column_formats = {
    "Status": "[=1]\"✓ Complete\";[=0]\"✗ Pending\";\"Unknown\"",
    "Grade": "[>=90]\"A\";[>=80]\"B\";[>=70]\"C\";\"F\""
}
```

#### Accounting Formats

Professional financial formatting:

```python
column_formats = {
    # Negative in parentheses, aligned decimals
    "P&L": "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
    
    # Simple accounting with red negatives
    "Balance": "$#,##0.00_);[Red]($#,##0.00)"
}
```

#### Custom Date/Time Formats

```python
column_formats = {
    "FullDate": "dddd, mmmm dd, yyyy",        # Monday, January 15, 2024
    "ShortDate": "mm/dd/yy",                  # 01/15/24
    "MonthYear": "mmmm yyyy",                 # January 2024
    "Quarter": "\"Q\"Q yyyy",                 # Q1 2024
    "TimeOnly": "h:mm AM/PM",                 # 3:45 PM
    "Timestamp": "yyyy-mm-dd hh:mm:ss"       # 2024-01-15 15:45:30
}
```

#### Fractions and Measurements

```python
column_formats = {
    "Inches": "# ?/16\"",           # Fractions in sixteenths with inch mark
    "Simple": "# ?/?",              # Simplest fraction
    "Eighths": "# ?/8",             # Fractions in eighths
    "Mixed": "# ??/??",             # Up to two-digit fractions
    "Feet": "#' ?/16\"",            # 5' 3/16"
}
```

#### Percentage Variations

```python
column_formats = {
    "Basic": "0%",                  # 15%
    "OneDecimal": "0.0%",           # 15.7%
    "TwoDecimal": "0.00%",          # 15.73%
    "WithSign": "+0.0%;-0.0%;0%",   # +15.7%, -3.2%, 0%
}
```

### Column Widths & Row Heights
```python
# Manual column widths
jet.write_sheet_arrow(
    df.to_arrow(),
    "sized.xlsx",
    column_widths={
        "Product": 20.0,      # 20 character units
        "Description": 50.0,
        "Price": 12.0
    },
    row_heights={
        1: 25.0,    # Header row height
        2: 18.0,    # First data row
        5: 30.0     # Fifth row
    }
)

# Column widths in pixels (converted automatically)
jet.write_sheet_arrow(
    df.to_arrow(),
    "pixel_widths.xlsx",
    column_widths={
        "Name": "150px",      # 150 pixels
        "Email": "200px",
        "Status": "80px"
    }
)

# Mix of manual and auto
jet.write_sheet_arrow(
    df.to_arrow(),
    "mixed_widths.xlsx",
    auto_width=True,          # Auto-calculate most columns
    column_widths={
        "ID": 8.0,            # Override: fixed width for ID
        "Notes": 60.0         # Override: extra wide for notes
    }
)
```

**Column Width Units:**
- Float (e.g., `20.0`) - Excel character units (width of '0' in standard font)
- String with "px" (e.g., `"150px"`) - Pixels (converted to character units)
- `"auto"` - Calculate from content (same as `auto_width=True`)

### Cell Styles

```python
cell_styles = [
    {
        "row": 2,
        "col": 1,
        "font": {
            "bold": True,
            "italic": False,
            "size": 14.0,
            "color": "FFFF0000",  # Red (ARGB format: AA=alpha, RR=red, GG=green, BB=blue)
            "name": "Arial"
        },
        "fill": {
            "pattern": "solid",  # Options: "solid", "gray125", "none"
            "fg_color": "FFFFFF00",  # Yellow
            "bg_color": None
        },
        "border": {
            "left": {"style": "thin", "color": "FF000000"},
            "right": {"style": "thick", "color": "FF000000"},
            "top": {"style": "medium", "color": "FF000000"},
            "bottom": {"style": "double", "color": "FF000000"}
        },
        "alignment": {
            "horizontal": "center",  # left, center, right, justify
            "vertical": "center",    # top, center, bottom
            "wrap_text": True,
            "text_rotation": 45      # 0-180 degrees, 255 for vertical
        },
        "number_format": "currency"
    }
]

jet.write_sheet_arrow(df.to_arrow(), "styled.xlsx", cell_styles=cell_styles)
```

### Text Rotation

Rotate text in cells for compact headers or labels:
```python
cell_styles = [{
    "row": 1,
    "col": 0,
    "alignment": {
        "horizontal": "center",
        "vertical": "center",
        "text_rotation": 45  # 0-180 degrees, or 255 for vertical text
    }
}]

jet.write_sheet_arrow(df.to_arrow(), "rotated.xlsx", cell_styles=cell_styles)
```

**Rotation values:**
- `0` - No rotation (default)
- `1-90` - Counterclockwise rotation
- `91-180` - Clockwise rotation (91 = -89°)
- `255` - Vertical text (top to bottom)

### Fill Patterns

Excel supports different fill patterns:
```python
# Solid fill (most common)
cell_styles = [{
    "row": 2,
    "col": 0,
    "fill": {
        "pattern": "solid",
        "fg_color": "FFFFFF00"  # Yellow
    }
}]

# Gray pattern (subtle shading)
cell_styles = [{
    "row": 2,
    "col": 0,
    "fill": {
        "pattern": "gray125",
        "fg_color": "FFD9D9D9"  # Light gray
    }
}]

# No fill (transparent)
cell_styles = [{
    "row": 2,
    "col": 0,
    "fill": {
        "pattern": "none"
    }
}]
```

### Complete Border Example

Apply different border styles to all four sides:
```python
cell_styles = [{
    "row": 2,
    "col": 0,
    "border": {
        "left": {"style": "thin", "color": "FF000000"},
        "right": {"style": "medium", "color": "FF000000"},
        "top": {"style": "thick", "color": "FF0070C0"},
        "bottom": {"style": "double", "color": "FF000000"}
    }
}]

jet.write_sheet_arrow(df.to_arrow(), "borders.xlsx", cell_styles=cell_styles)
```

**Available border styles:**
- `"thin"` - Standard thin line
- `"medium"` - Medium weight line
- `"thick"` - Thick line
- `"double"` - Double line
- `"dotted"` - Dotted line
- `"dashed"` - Dashed line

**Color Format Guide:**
- Colors use ARGB hexadecimal format: `AARRGGBB`
- `AA` = Alpha (transparency): `FF` = fully opaque, `00` = fully transparent
- `RR` = Red component: `00` = no red, `FF` = maximum red
- `GG` = Green component: `00` = no green, `FF` = maximum green
- `BB` = Blue component: `00` = no blue, `FF` = maximum blue

Common colors: `FFFF0000` (red), `FF00FF00` (green), `FF0000FF` (blue), `FFFFFF00` (yellow), `FF000000` (black), `FFFFFFFF` (white)

For more colors and an interactive picker, see the [External Resources](#-external-resources--references) section below.

### Excel Tables

Create formatted Excel tables with built-in styles, sorting, and filtering capabilities.

### Basic Table
```python
tables = [{
    "name": "ProductTable",
    "display_name": "Product Data",
    "start_row": 1,
    "start_col": 0,
    "end_row": 0,      # NEW: 0 means auto-calculate from data
    "end_col": 0,      # NEW: 0 means auto-calculate from data
    "style": "TableStyleMedium2"
}]

jet.write_sheet_arrow(df.to_arrow(), "table.xlsx", tables=tables)
```

### Auto-Sizing Tables

Let Jetxl automatically calculate table dimensions based on your DataFrame:
```python
import polars as pl
import jetxl as jet

df = pl.DataFrame({
    "Product": ["A", "B", "C", "D", "E"],  # 5 rows
    "Price": [10, 20, 30, 40, 50],
    "Qty": [100, 200, 150, 300, 250]       # 3 columns
})

tables = [{
    "name": "AutoTable",
    "start_row": 1,    # Table starts at row 1 (header)
    "start_col": 0,    # Column A
    "end_row": 0,      # Auto: becomes 6 (1 header + 5 data rows)
    "end_col": 0,      # Auto: becomes 2 (columns A, B, C = indices 0, 1, 2)
    "style": "TableStyleMedium2"
}]

jet.write_sheet_arrow(df.to_arrow(), "auto_table.xlsx", tables=tables)
```

**Manual vs Auto-Sizing:**
```python
# Manual (explicit range)
table = {
    "name": "ManualTable",
    "start_row": 1,
    "start_col": 0,
    "end_row": 100,    # Exactly 100 rows
    "end_col": 5       # Columns A-F
}

# Auto (adapts to DataFrame)
table = {
    "name": "AutoTable", 
    "start_row": 1,
    "start_col": 0,
    "end_row": 0,      # Uses all DataFrame rows
    "end_col": 0       # Uses all DataFrame columns
}

# Mixed (partial auto)
table = {
    "name": "MixedTable",
    "start_row": 1,
    "start_col": 0,
    "end_row": 50,     # Fixed 50 rows
    "end_col": 0       # Auto-calculate columns
}
```

**Auto-calculation rules:**
- `end_row = 0` → calculated as `start_row + num_data_rows`
- `end_col = 0` → calculated as `start_col + num_columns - 1`
- If table starts after row 1, a header row is automatically inserted
- Manual values (non-zero) are used as-is

### Available Table Styles

Excel provides many built-in table styles that you can use with Jetxl. The styles are organized into three categories:

**Light Table Styles** (Minimal emphasis, subtle colors)
- `TableStyleLight1` through `TableStyleLight21`
- Best for: Professional reports, financial statements, clean presentations

**Medium Table Styles** (Moderate emphasis, balanced design)
- `TableStyleMedium1` through `TableStyleMedium28`
- Best for: Data analysis, dashboards, general-purpose tables

**Dark Table Styles** (Strong emphasis, high contrast)
- `TableStyleDark1` through `TableStyleDark11`
- Best for: Executive summaries, presentations, highlighting key data

**Visual Reference**: To see examples of all table styles, visit [Microsoft's Format an Excel Table guide](https://support.microsoft.com/en-us/office/format-an-excel-table-6789619f-c889-495c-99c2-2f971c0e2370) which includes screenshots of each style.

**Additional Resources**:
- [Create and format Excel tables](https://support.microsoft.com/en-us/office/create-and-format-tables-e81aa349-b006-4f8a-9806-5af9df0ac664) - Official Microsoft documentation
- [Excel table overview](https://support.microsoft.com/en-us/office/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c) - Complete guide to table features

### Multiple Tables in One Sheet

```python
# Create two separate tables in the same sheet
tables = [
    {
        "name": "SalesTable",
        "start_row": 1,
        "start_col": 0,
        "end_row": 10,
        "end_col": 3,
        "style": "TableStyleMedium9"
    },
    {
        "name": "SummaryTable",
        "start_row": 12,
        "start_col": 0,
        "end_row": 15,
        "end_col": 2,
        "style": "TableStyleLight16"
    }
]

jet.write_sheet_arrow(df.to_arrow(), "multi_tables.xlsx", tables=tables)
```

### Table Configuration Options

```python
table = {
    "name": "MyTable",                # Required: Unique table identifier
    "display_name": "My Data",        # Optional: User-friendly name
    "start_row": 1,                   # Required: First row (1-indexed)
    "start_col": 0,                   # Required: First column (0-indexed)
    "end_row": 100,                   # Required: Last row
    "end_col": 5,                     # Required: Last column
    "style": "TableStyleMedium2",     # Optional: Table style name
    "show_first_column": False,       # Optional: Bold first column (default: False)
    "show_last_column": False,        # Optional: Bold last column (default: False)
    "show_row_stripes": True,         # Optional: Alternating rows (default: True)
    "show_column_stripes": False      # Optional: Alternating columns (default: False)
}
```

**Note:** Excel tables automatically include:
- Header row with filter dropdowns
- Structured references for formulas
- Automatic formatting and styling
- Sort and filter capabilities

## 📊 Excel Charts

Create professional charts and visualizations directly in your Excel files. Jetxl supports six chart types with extensive customization options including stacked charts, data labels, styling, and advanced formatting.

### Chart Types

Jetxl supports the following chart types:
- **Column Chart** - Vertical bars, ideal for comparing values across categories
- **Bar Chart** - Horizontal bars, good for comparing items
- **Line Chart** - Shows trends over time or continuous data
- **Pie Chart** - Displays proportions of a whole
- **Scatter Chart** - Shows relationships between two numerical variables
- **Area Chart** - Similar to line chart but with filled areas

### Basic Column Chart

```python
import polars as pl
import jetxl as jet

# Create sample data
df = pl.DataFrame({
    "Month": ["Jan", "Feb", "Mar", "Apr", "May"],
    "Sales": [1000, 1500, 1200, 1800, 2000],
    "Costs": [800, 900, 850, 1000, 1100]
})

# Define a column chart
charts = [{
    "chart_type": "column",
    "start_row": 1,           # Data starts at row 1 (header)
    "start_col": 0,           # First column (Month)
    "end_row": 5,             # Last data row
    "end_col": 2,             # Last column (Costs)
    "from_col": 4,            # Chart position: start column
    "from_row": 1,            # Chart position: start row
    "to_col": 12,             # Chart position: end column
    "to_row": 15,             # Chart position: end row
    "title": "Monthly Sales and Costs",
    "category_col": 0,        # Use first column (Month) for X-axis
    "show_legend": True,
    "x_axis_title": "Month",
    "y_axis_title": "Amount ($)"
}]

jet.write_sheet_arrow(
    df.to_arrow(),
    "chart_example.xlsx",
    charts=charts
)
```

### Chart Configuration

Every chart requires these basic parameters:

```python
chart = {
    # Required: Chart type
    "chart_type": "column",  # column, bar, line, pie, scatter, area
    
    # Required: Data range (1-indexed for rows, 0-indexed for columns)
    # Option 1: Individual parameters
    "start_row": 1,          # First data row (including header)
    "start_col": 0,          # First data column
    "end_row": 10,           # Last data row
    "end_col": 3,            # Last data column
    
    # Option 2: Tuple format (alternative to above)
    "data_range": (0, 0, 9, 3),  # (start_row, start_col, end_row, end_col)
    
    # Required: Chart position on worksheet
    "from_col": 5,           # Starting column for chart
    "from_row": 1,           # Starting row for chart
    "to_col": 15,            # Ending column for chart
    "to_row": 20,            # Ending row for chart
    
    # Optional: Chart customization
    "title": "My Chart",                # Chart title
    "category_col": 0,                  # Column to use for category axis (X-axis)
    "series_names": ["Series 1", "Series 2"],  # Custom series names
    "show_legend": True,                # Show/hide legend
    "legend_position": "right",         # Legend position: "right", "left", "top", "bottom", "none"
    "x_axis_title": "Categories",       # X-axis label
    "y_axis_title": "Values",           # Y-axis label
    
    # Optional: Advanced styling
    "stacked": True,                    # Stack series (column, bar, line, area)
    "percent_stacked": True,            # Stack as 100% (column, bar, line, area)
    "show_data_labels": True,           # Show data labels on chart
    "chart_style": 104,                 # Excel chart style (1-48)
    
    # Optional: Axis scaling
    "axis_min": 0.0,                    # Minimum Y-axis value
    "axis_max": 100.0,                  # Maximum Y-axis value
    
    # Optional: Title styling
    "title_bold": True,                 # Bold title text
    "title_font_size": 1800,            # Title font size (hundredths of point, e.g., 1800 = 18pt)
    "title_color": "FF0000",            # Title color (ARGB hex)
    
    # Optional: Axis title styling
    "axis_title_bold": True,            # Bold axis title text
    "axis_title_font_size": 1200,       # Axis title font size (hundredths of point)
    "axis_title_color": "00B050",       # Axis title color (ARGB hex)
    
    # Optional: Legend styling
    "legend_bold": True,                # Bold legend text
    "legend_font_size": 1000,           # Legend font size (hundredths of point)
}
```

### Stacked Charts

Create stacked charts to show composition or percent stacked to show proportions:

```python
import polars as pl
import jetxl as jet

df = pl.DataFrame({
    "Quarter": ["Q1", "Q2", "Q3", "Q4"],
    "Revenue": [25000, 28000, 31000, 35000],
    "Profit": [5000, 6500, 7200, 8500],
    "Expenses": [20000, 21500, 23800, 26500]
})

# Regular stacked column chart
stacked_chart = [{
    "chart_type": "column",
    "data_range": (0, 0, 3, 3),
    "from_col": 5,
    "from_row": 0,
    "to_col": 15,
    "to_row": 20,
    "title": "Stacked Revenue Components",
    "category_col": 0,
    "series_names": ["Revenue", "Profit", "Expenses"],
    "stacked": True,  # Stack the series
    "show_data_labels": True,
    "x_axis_title": "Quarter",
    "y_axis_title": "Amount ($)"
}]

# Percent stacked column chart
percent_chart = [{
    "chart_type": "column",
    "data_range": (0, 0, 3, 3),
    "from_col": 5,
    "from_row": 22,
    "to_col": 15,
    "to_row": 42,
    "title": "Percentage Distribution",
    "category_col": 0,
    "series_names": ["Revenue", "Profit", "Expenses"],
    "percent_stacked": True,  # Stack as 100%
    "show_data_labels": True,
    "x_axis_title": "Quarter",
    "y_axis_title": "Percentage"
}]

jet.write_sheet_arrow(
    df.to_arrow(),
    "stacked_charts.xlsx",
    charts=[*stacked_chart, *percent_chart]
)
```

**Note:** `stacked` and `percent_stacked` work with column, bar, line, and area charts.

### Understanding Chart Styles

Excel provides 48 pre-defined chart styles that apply coordinated colors, effects, and formatting. Each chart type interprets these styles differently.

**Chart Style Numbers (1-48):**

Chart styles are organized into categories:
- **1-10**: Colorful variations with different color schemes
- **11-16**: Monochrome styles (black, white, gray variations)
- **17-32**: Colorful outlined styles with borders
- **33-40**: Soft color palettes
- **41-48**: Modern gradient and flat design styles

**Popular Styles by Use Case:**
```python
# Professional/Corporate
"chart_style": 2    # Blue colorful
"chart_style": 11   # Monochrome gray
"chart_style": 26   # Dark professional

# Modern/Vibrant
"chart_style": 42   # Gradient modern
"chart_style": 102  # Contemporary (if available)
"chart_style": 104  # Bright modern

# Print-Friendly
"chart_style": 11   # Black and white
"chart_style": 15   # High contrast
```

**How to Find Your Preferred Style:**

The best way to discover chart styles is to test them in Excel:
1. Create a chart in Excel manually
2. Click the chart and go to Chart Design → Chart Styles
3. Preview different styles in the gallery
4. Note the style you like - styles are numbered in the gallery order
5. Use that number in Jetxl's `chart_style` parameter

**Example - Testing Multiple Styles:**
```python
# Create the same chart with different styles to compare
for style_num in [2, 11, 26, 42, 104]:
    charts = [{
        "chart_type": "column",
        "data_range": (0, 0, 3, 2),
        "from_col": 5,
        "from_row": 0,
        "to_col": 15,
        "to_row": 20,
        "title": f"Chart Style {style_num}",
        "chart_style": style_num,
        "category_col": 0
    }]
    jet.write_sheet_arrow(df.to_arrow(), f"style_{style_num}.xlsx", charts=charts)
```

**Chart Style Reference:**
- [Excel Chart Styles Gallery](https://support.microsoft.com/en-us/office/available-chart-types-in-office-a6187218-807e-4103-9e0a-27cdb19afb90) - Microsoft's official guide
- Styles are part of the Office Open XML standard
- Different Excel versions may render styles slightly differently
- Not all style numbers work with all chart types

### Understanding Font Sizes

Font sizes in charts use Excel's internal unit system based on the Office Open XML (OOXML) standard.

**Font Size Format:**
- Values are in **hundredths of a point**
- 1 point = 100 units
- Standard Excel points = units ÷ 100

**Common Font Size Conversions:**
```python
# Title Sizes
"title_font_size": 800   # 8pt  - Small title
"title_font_size": 1000  # 10pt - Compact title
"title_font_size": 1200  # 12pt - Standard title
"title_font_size": 1400  # 14pt - Medium title
"title_font_size": 1600  # 16pt - Large title
"title_font_size": 1800  # 18pt - Extra large title
"title_font_size": 2400  # 24pt - Presentation title
"title_font_size": 3200  # 32pt - Header/banner

# Axis Title Sizes
"axis_title_font_size": 900   # 9pt  - Small
"axis_title_font_size": 1000  # 10pt - Standard
"axis_title_font_size": 1100  # 11pt - Medium
"axis_title_font_size": 1200  # 12pt - Large

# Legend Sizes
"legend_font_size": 800   # 8pt  - Compact
"legend_font_size": 900   # 9pt  - Small
"legend_font_size": 1000  # 10pt - Standard
"legend_font_size": 1100  # 11pt - Medium
```

**Why Hundredths of a Point?**
- Precise control over text sizing
- Matches Excel's internal OOXML format
- Allows fractional point sizes (e.g., 1050 = 10.5pt)
- Same system used throughout Microsoft Office

**Quick Conversion Formula:**
```python
# Points to hundredths
hundredths = points * 100

# Hundredths to points
points = hundredths / 100

# Example: 14pt title
title_font_size = 14 * 100  # = 1400
```

### Understanding Axis Scaling

Control the Y-axis range to focus on relevant data ranges or maintain consistent scales across charts.

**Basic Axis Scaling:**
```python
charts = [{
    "chart_type": "line",
    "data_range": (0, 0, 3, 1),
    "from_col": 5, "from_row": 0,
    "to_col": 15, "to_row": 20,
    "title": "Test Scores",
    "category_col": 0,
    "axis_min": 0.0,      # Y-axis starts at 0
    "axis_max": 100.0,    # Y-axis ends at 100
}]
```

**When to Use Axis Scaling:**

1. **Percentage Data (0-100%):**
```python
"axis_min": 0.0,
"axis_max": 100.0,  # Perfect for showing 0-100% range
```

2. **Normalized Data (0-1):**
```python
"axis_min": 0.0,
"axis_max": 1.0,  # For decimal percentages (use with percent_stacked)
```

3. **Focus on Variance:**
```python
# Data ranges from 85-95, default would show 0-100
"axis_min": 80.0,
"axis_max": 100.0,  # Zoom in to show meaningful differences
```

4. **Compare Multiple Charts:**
```python
# Use same scale across all charts for fair comparison
"axis_min": 0.0,
"axis_max": 50000.0,  # All revenue charts use same scale
```

5. **Symmetric Ranges:**
```python
# For profit/loss or change data
"axis_min": -10000.0,
"axis_max": 10000.0,  # Symmetric around zero
```

**Special Case - Percent Stacked Charts:**
```python
# When using percent_stacked, values are 0.0 to 1.0
charts = [{
    "chart_type": "column",
    "data_range": (0, 0, 3, 3),
    "percent_stacked": True,
    "axis_min": 0.0,    # 0% = 0.0
    "axis_max": 1.0,    # 100% = 1.0
    "y_axis_title": "Percentage"
}]
```

**Axis Scaling Best Practices:**
- Always start at 0.0 for bar/column charts to avoid misleading visuals
- Use axis_min > 0 only for line charts showing trends
- For percent_stacked, always use axis_min=0.0 and axis_max=1.0
- Omit axis_min/axis_max to let Excel auto-scale
- Use consistent scales when comparing multiple charts

### Advanced Chart Styling

Customize chart appearance with styling options:

```python
import polars as pl
import jetxl as jet

df = pl.DataFrame({
    "Month": ["Jan", "Feb", "Mar", "Apr"],
    "Sales": [10000, 12000, 11000, 13000],
    "Target": [9000, 10000, 12000, 14000]
})

charts = [{
    "chart_type": "bar",
    "data_range": (0, 0, 3, 2),
    "from_col": 5,
    "from_row": 0,
    "to_col": 15,
    "to_row": 20,
    
    # Title styling
    "title": "Sales vs Target",
    "title_bold": True,
    "title_font_size": 1800,      # 18pt (in hundredths of point)
    "title_color": "0070C0",      # Blue (ARGB hex without alpha)
    
    # Axis titles with styling
    "x_axis_title": "Amount ($)",
    "y_axis_title": "Month",
    "axis_title_bold": True,
    "axis_title_font_size": 1200,  # 12pt
    "axis_title_color": "00B050",  # Green
    
    # Legend styling
    "show_legend": True,
    "legend_position": "bottom",
    "legend_bold": True,
    "legend_font_size": 1000,      # 10pt
    
    # Chart style and data labels
    "chart_style": 104,             # Apply Excel chart style
    "show_data_labels": True,
    
    # Axis scaling
    "axis_min": 0.0,
    "axis_max": 15000.0,
    
    "category_col": 0
}]

jet.write_sheet_arrow(df.to_arrow(), "styled_chart.xlsx", charts=charts)
```

**Chart Style Numbers:**
- Excel supports chart styles numbered 1-48
- Each chart type has different style variations
- Common styles: 2 (colorful), 11 (monochrome), 26 (dark), 42 (gradient)
- Experiment with different numbers to find your preferred style

**Font Sizes:**
- Font sizes are specified in hundredths of a point
- Examples: 800 = 8pt, 1000 = 10pt, 1200 = 12pt, 1800 = 18pt

### Data Labels on Charts

Add data labels to show values directly on chart elements:

```python
df = pl.DataFrame({
    "Product": ["A", "B", "C", "D"],
    "Sales": [250, 380, 420, 290]
})

charts = [{
    "chart_type": "column",
    "data_range": (0, 0, 3, 1),
    "from_col": 3,
    "from_row": 0,
    "to_col": 11,
    "to_row": 15,
    "title": "Product Sales",
    "category_col": 0,
    "show_data_labels": True,  # Display values on bars
    "x_axis_title": "Product",
    "y_axis_title": "Units Sold"
}]

jet.write_sheet_arrow(df.to_arrow(), "chart_with_labels.xlsx", charts=charts)
```

### Axis Scaling

Control the Y-axis range for better visualization:

```python
df = pl.DataFrame({
    "Month": ["Jan", "Feb", "Mar", "Apr"],
    "Score": [92, 88, 95, 91]
})

charts = [{
    "chart_type": "line",
    "data_range": (0, 0, 3, 1),
    "from_col": 3,
    "from_row": 0,
    "to_col": 11,
    "to_row": 15,
    "title": "Test Scores",
    "category_col": 0,
    "axis_min": 80.0,     # Start Y-axis at 80
    "axis_max": 100.0,    # End Y-axis at 100
    "x_axis_title": "Month",
    "y_axis_title": "Score",
    "show_data_labels": True
}]

jet.write_sheet_arrow(df.to_arrow(), "scaled_chart.xlsx", charts=charts)
```

### Pie Chart with Styling

```python
df = pl.DataFrame({
    "Category": ["North", "South", "East", "West"],
    "Sales": [2500, 1800, 2200, 1500]
})

charts = [{
    "chart_type": "pie",
    "data_range": (0, 0, 3, 1),
    "from_col": 3,
    "from_row": 0,
    "to_col": 10,
    "to_row": 15,
    "title": "Regional Distribution",
    "title_bold": True,
    "title_font_size": 1800,
    "category_col": 0,
    "show_data_labels": True,  # Show percentages/values
    "legend_font_size": 1100
}]

jet.write_sheet_arrow(df.to_arrow(), "pie_chart.xlsx", charts=charts)
```

### Scatter Chart with Axis Control

```python
df = pl.DataFrame({
    "X": [1, 2, 3, 4, 5],
    "Y1": [10, 25, 30, 45, 60],
    "Y2": [15, 20, 35, 40, 55]
})

charts = [{
    "chart_type": "scatter",
    "data_range": (0, 0, 4, 2),
    "from_col": 5,
    "from_row": 0,
    "to_col": 15,
    "to_row": 20,
    "title": "Correlation Analysis",
    "axis_min": 0.0,
    "axis_max": 70.0,
    "x_axis_title": "X Values",
    "y_axis_title": "Y Values",
    "show_data_labels": False,
    "series_names": ["Series 1", "Series 2"]
}]

jet.write_sheet_arrow(df.to_arrow(), "scatter_chart.xlsx", charts=charts)
```

### Multiple Charts in One Sheet

You can add multiple charts to the same worksheet:

```python
df = pl.DataFrame({
    "Month": ["Jan", "Feb", "Mar", "Apr"],
    "Revenue": [10000, 12000, 11000, 13000],
    "Expenses": [7000, 8000, 7500, 8500],
    "Profit": [3000, 4000, 3500, 4500]
})

charts = [
    {
        # Column chart for Revenue and Expenses
        "chart_type": "column",
        "data_range": (0, 0, 3, 2),
        "from_col": 5,
        "from_row": 0,
        "to_col": 13,
        "to_row": 15,
        "title": "Revenue & Expenses",
        "title_bold": True,
        "category_col": 0,
        "series_names": ["Revenue", "Expenses"],
        "show_data_labels": True,
        "x_axis_title": "Month",
        "y_axis_title": "Amount ($)"
    },
    {
        # Line chart for Profit trend
        "chart_type": "line",
        "data_range": (0, 0, 3, 3),
        "from_col": 5,
        "from_row": 17,
        "to_col": 13,
        "to_row": 32,
        "title": "Profit Trend",
        "title_color": "00B050",
        "category_col": 0,
        "series_names": ["Profit"],
        "x_axis_title": "Month",
        "y_axis_title": "Profit ($)",
        "chart_style": 26
    }
]

jet.write_sheet_arrow(df.to_arrow(), "multiple_charts.xlsx", charts=charts)
```

### Complete Advanced Chart Example

Here's a comprehensive example using all available chart features:

```python
import polars as pl
import jetxl as jet

df = pl.DataFrame({
    "Quarter": ["Q1", "Q2", "Q3", "Q4"],
    "Revenue": [25000, 28000, 31000, 35000],
    "Profit": [5000, 6500, 7200, 8500],
    "Expenses": [20000, 21500, 23800, 26500]
})

charts = [{
    "chart_type": "column",
    
    # Data range - use tuple or individual parameters
    "data_range": (0, 0, 3, 3),
    # OR: "start_row": 1, "start_col": 0, "end_row": 4, "end_col": 3,
    
    # Chart position
    "from_col": 5,
    "from_row": 0,
    "to_col": 15,
    "to_row": 20,
    
    # Basic settings
    "title": "Quarterly Financial Performance",
    "category_col": 0,
    "series_names": ["Revenue", "Profit", "Expenses"],
    
    # Stacking
    "percent_stacked": True,  # Show as percentages
    
    # Title styling
    "title_bold": True,
    "title_font_size": 1600,
    "title_color": "0070C0",
    
    # Axis configuration
    "x_axis_title": "Quarter",
    "y_axis_title": "Percentage",
    "axis_title_bold": True,
    "axis_title_font_size": 1200,
    "axis_title_color": "00B050",
    "axis_min": 0.0,
    "axis_max": 1.0,
    
    # Legend
    "show_legend": True,
    "legend_position": "bottom",
    "legend_bold": True,
    "legend_font_size": 1000,
    
    # Visual enhancements
    "show_data_labels": True,
    "chart_style": 102
}]

jet.write_sheet_arrow(
    df.to_arrow(),
    "complete_chart.xlsx",
    charts=charts
)
```

## 🖼️ Excel Images

Add images (logos, charts, diagrams) to your Excel sheets with precise positioning control.

### Supported Image Formats

| Format | Extensions | Best For | Notes |
|--------|------------|----------|-------|
| PNG | `.png` | Logos, screenshots | Lossless, supports transparency |
| JPEG | `.jpg`, `.jpeg` | Photos | Smaller file size, no transparency |
| GIF | `.gif` | Simple graphics | Limited colors, supports animation |
| BMP | `.bmp` | Windows bitmaps | Large file size, uncompressed |
| TIFF | `.tiff`, `.tif` | High-quality images | Professional printing |

### Adding Images from Files

### Adding Images from Files

The simplest way to add images is from file paths:

```python
import polars as pl
import jetxl as jet

df = pl.DataFrame({
    "Product": ["Widget A", "Widget B", "Widget C"],
    "Sales": [1000, 1500, 1200]
})

images = [{
    "path": "company_logo.png",
    "from_col": 0,   # Column A (0-based)
    "from_row": 0,   # Row 1 (0-based)
    "to_col": 2,     # Column C
    "to_row": 5      # Row 6
}]

jet.write_sheet_arrow(
    df.to_arrow(),
    "report_with_logo.xlsx",
    images=images
)
```

### Adding Images from Bytes

Load images from memory (useful for API responses, databases, or generated images):

```python
import requests
import jetxl as jet

# Download image from URL
response = requests.get("https://example.com/chart.png")
image_bytes = response.content

# Or read from file
with open("logo.png", "rb") as f:
    image_bytes = f.read()

images = [{
    "data": image_bytes,
    "extension": "png",  # Required when using bytes
    "from_col": 5,
    "from_row": 1,
    "to_col": 12,
    "to_row": 15
}]

jet.write_sheet_arrow(
    df.to_arrow(),
    "report.xlsx",
    images=images
)
```

### Multiple Images

Add multiple images to the same sheet:

```python
images = [
    {
        # Company logo in top-left
        "path": "company_logo.png",
        "from_col": 0,
        "from_row": 0,
        "to_col": 2,
        "to_row": 4
    },
    {
        # Product image on the right
        "path": "product_photo.jpg",
        "from_col": 8,
        "from_row": 2,
        "to_col": 12,
        "to_row": 10
    },
    {
        # Chart at the bottom
        "path": "sales_chart.png",
        "from_col": 0,
        "from_row": 15,
        "to_col": 10,
        "to_row": 30
    }
]

jet.write_sheet_arrow(
    df.to_arrow(),
    "multi_image_report.xlsx",
    images=images
)
```

### Image Positioning Guide

Images are positioned using Excel's column/row coordinates:
- **Columns** are 0-indexed: A=0, B=1, C=2, etc.
- **Rows** are 0-indexed: 0=row 1, 1=row 2, etc.

```python
# Position image from B3 to F10
image = {
    "path": "image.png",
    "from_col": 1,   # Column B (0-based)
    "from_row": 2,   # Row 3 (0-based)
    "to_col": 5,     # Column F
    "to_row": 9      # Row 10
}
```

**Size Recommendations:**
- **Small**: 2-4 columns × 5-8 rows (logos, icons)
- **Medium**: 4-6 columns × 8-12 rows (product photos)
- **Large**: 6-10 columns × 12-20 rows (charts, diagrams)

### Combining Images with Data

Create professional reports with logos, data, and visualizations:

```python
import polars as pl
import jetxl as jet

# Sample data
df = pl.DataFrame({
    "Month": ["Jan", "Feb", "Mar", "Apr"],
    "Revenue": [10000, 12000, 11000, 13000],
    "Costs": [7000, 8000, 7500, 8500]
})

# Add company logo, data table, and chart image
jet.write_sheet_arrow(
    df.to_arrow(),
    "monthly_report.xlsx",
    sheet_name="Financial Report",
    styled_headers=True,
    freeze_rows=1,
    column_formats={
        "Revenue": "currency",
        "Costs": "currency"
    },
    images=[
        {
            # Logo at top
            "path": "company_logo.png",
            "from_col": 0,
            "from_row": 0,
            "to_col": 2,
            "to_row": 3
        },
        {
            # Visualization chart
            "path": "revenue_chart.png",
            "from_col": 5,
            "from_row": 5,
            "to_col": 15,
            "to_row": 25
        }
    ]
)
```

### Images with Charts and Tables

Combine all visualization features:

```python
df = pl.DataFrame({
    "Product": ["A", "B", "C", "D"],
    "Q1": [100, 150, 120, 180],
    "Q2": [110, 160, 130, 190],
    "Q3": [120, 170, 140, 200]
})

jet.write_sheet_arrow(
    df.to_arrow(),
    "complete_dashboard.xlsx",
    tables=[{
        "name": "SalesTable",
        "start_row": 1,
        "start_col": 0,
        "end_row": 4,
        "end_col": 3,
        "style": "TableStyleMedium2"
    }],
    charts=[{
        "chart_type": "column",
        "start_row": 1,
        "start_col": 0,
        "end_row": 4,
        "end_col": 3,
        "from_col": 5,
        "from_row": 5,
        "to_col": 13,
        "to_row": 20,
        "title": "Quarterly Sales"
    }],
    images=[{
        "path": "company_logo.png",
        "from_col": 0,
        "from_row": 0,
        "to_col": 2,
        "to_row": 3
    }]
)
```

### Images Across Multiple Sheets

Each sheet can have its own images:

```python
df_summary = pl.DataFrame({"Metric": ["Total Sales"], "Value": [50000]})
df_details = pl.DataFrame({"Product": ["A", "B"], "Sales": [30000, 20000]})

sheets = [
    {
        "data": df_summary.to_arrow(),
        "name": "Summary",
        "images": [{
            "path": "company_logo.png",
            "from_col": 0,
            "from_row": 0,
            "to_col": 2,
            "to_row": 4
        }]
    },
    {
        "data": df_details.to_arrow(),
        "name": "Details",
        "images": [{
            "path": "product_breakdown.png",
            "from_col": 4,
            "from_row": 1,
            "to_col": 12,
            "to_row": 15
        }]
    }
]

jet.write_sheets_arrow(sheets, "multi_sheet_report.xlsx", num_threads=2)
```

### Working with Generated Images

Combine with image generation libraries:

```python
import matplotlib.pyplot as plt
import io
import jetxl as jet

# Generate a chart with matplotlib
fig, ax = plt.subplots()
ax.plot([1, 2, 3, 4], [10, 20, 15, 25])
ax.set_title("Sales Trend")

# Save to bytes
img_buffer = io.BytesIO()
fig.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
img_bytes = img_buffer.getvalue()
plt.close(fig)

# Add to Excel
jet.write_sheet_arrow(
    df.to_arrow(),
    "report_with_chart.xlsx",
    images=[{
        "data": img_bytes,
        "extension": "png",
        "from_col": 5,
        "from_row": 1,
        "to_col": 15,
        "to_row": 20
    }]
)
```

### Image Best Practices

1. **File Formats**
   - Use PNG for logos and screenshots (lossless, supports transparency)
   - Use JPEG for photos (smaller file size)
   - Use GIF for simple animations (limited color palette)

2. **Image Size**
   - Optimize images before embedding to reduce file size
   - Use appropriate dimensions for your target (don't embed 4K images for small displays)
   - Consider using PIL/Pillow to resize images programmatically

3. **Performance**
   - Large images increase Excel file size
   - Multiple large images can slow down Excel opening time
   - Compress images before embedding when possible

4. **Positioning**
   - Leave space around images for readability
   - Align images with data columns when possible
   - Use consistent sizing for professional appearance

## 🔗 Hyperlinks

```python
hyperlinks = [
    (2, 0, "https://example.com", "Visit Example"),  # Row 2, Col 0
    (3, 0, "https://google.com", None),              # Display URL as text
    (4, 2, "mailto:user@example.com", "Email Us")
]

jet.write_sheet_arrow(df.to_arrow(), "links.xlsx", hyperlinks=hyperlinks)
```

## 📢 Formulas
```python
formulas = [
    (2, 3, "=SUM(A2:C2)", None),           # Simple formula
    (5, 3, "=AVERAGE(D2:D4)", "45.5"),     # Formula with cached value
    (6, 3, "=IF(D5>50,\"High\",\"Low\")", None)
]

jet.write_sheet_arrow(df.to_arrow(), "formulas.xlsx", formulas=formulas)
```

### Understanding Cached Values

The cached value is the pre-calculated result shown before Excel recalculates the formula:
```python
formulas = [
    # No cached value - Excel calculates on open
    (2, 3, "=SUM(A2:C2)", None),
    
    # With cached value - shows "45.5" until Excel recalculates
    (5, 3, "=AVERAGE(D2:D4)", "45.5"),
]
```

**When to use cached values:**
- Formulas that reference external data sources
- Complex calculations that take time to compute
- When you want to show a result before Excel opens
- Cross-workbook references that may not be available

**When to use None:**
- Simple formulas (SUM, AVERAGE of local cells)
- When you want Excel to always calculate fresh
- Formulas with volatile functions (NOW, RAND)

## 🔀 Merge Cells

```python
merge_cells = [
    (1, 0, 1, 3),  # Merge A1:D1 (start_row, start_col, end_row, end_col)
    (2, 0, 5, 0),  # Merge A2:A5
]

jet.write_sheet_arrow(df.to_arrow(), "merged.xlsx", merge_cells=merge_cells)
```

## ✅ Data Validation

### Dropdown Lists

```python
validations = [{
    "start_row": 2,
    "start_col": 0,
    "end_row": 100,
    "end_col": 0,
    "type": "list",
    "items": ["Option A", "Option B", "Option C"],
    "show_dropdown": True,
    "error_title": "Invalid Selection",
    "error_message": "Please select from the dropdown"
}]

jet.write_sheet_arrow(df.to_arrow(), "validation.xlsx", data_validations=validations)
```

### Number Ranges

```python
validations = [{
    "start_row": 2,
    "start_col": 1,
    "end_row": 100,
    "end_col": 1,
    "type": "whole_number",
    "min": 1,
    "max": 100,
    "error_title": "Out of Range",
    "error_message": "Value must be between 1 and 100"
}]
```

### Decimal Ranges

```python
validations = [{
    "start_row": 2,
    "start_col": 2,
    "end_row": 100,
    "end_col": 2,
    "type": "decimal",
    "min": 0.0,
    "max": 100.0
}]
```

### Text Length

Validate text input length:
```python
validations = [{
    "start_row": 2,
    "start_col": 0,
    "end_row": 100,
    "end_col": 0,
    "type": "text_length",
    "min": 3,
    "max": 20,
    "error_title": "Invalid Username",
    "error_message": "Username must be 3-20 characters long"
}]

jet.write_sheet_arrow(df.to_arrow(), "validation.xlsx", data_validations=validations)
```

## 🎨 Conditional Formatting

### Cell Value Rules

```python
conditional_formats = [{
    "start_row": 2,
    "start_col": 2,
    "end_row": 100,
    "end_col": 2,
    "rule_type": "cell_value",
    "operator": "greater_than",  # less_than, equal, not_equal, etc.
    "value": "50",
    "priority": 1,
    "style": {
        "font": {
            "bold": True,
            "color": "FFFF0000"  # Red text
        },
        "fill": {
            "pattern": "solid",
            "fg_color": "FFFFFF00"  # Yellow background
        }
    }
}]

jet.write_sheet_arrow(df.to_arrow(), "conditional.xlsx", conditional_formats=conditional_formats)
```
### All Comparison Operators

The `cell_value` rule type supports these operators:
```python
# Greater than
"operator": "greater_than",  "value": "100"

# Less than
"operator": "less_than",  "value": "50"

# Equal to
"operator": "equal",  "value": "0"

# Not equal to
"operator": "not_equal",  "value": "0"

# Greater than or equal
"operator": "greater_than_or_equal",  "value": "100"

# Less than or equal
"operator": "less_than_or_equal",  "value": "50"

# Between (use comma-separated values)
"operator": "between",  "value": "10,100"
```

### Color Scale Variations
```python
# Two-color scale (min to max)
conditional_formats = [{
    "start_row": 2,
    "start_col": 2,
    "end_row": 100,
    "end_col": 2,
    "rule_type": "color_scale",
    "min_color": "FFF8696B",  # Red
    "max_color": "FF63BE7B",  # Green
    "priority": 1
}]

# Three-color scale (min to mid to max)
# Better for showing deviation from average/target
conditional_formats = [{
    "start_row": 2,
    "start_col": 2,
    "end_row": 100,
    "end_col": 2,
    "rule_type": "color_scale",
    "min_color": "FFF8696B",  # Red for low values
    "mid_color": "FFFFEB84",  # Yellow for medium values
    "max_color": "FF63BE7B",  # Green for high values
    "priority": 1
}]
```

### Top/Bottom N Values

Highlight the highest or lowest values in a range:
```python
# Highlight top 10 values
conditional_formats = [{
    "start_row": 2,
    "start_col": 2,
    "end_row": 100,
    "end_col": 2,
    "rule_type": "top10",
    "rank": 10,
    "bottom": False,  # Top 10 (set to True for bottom 10)
    "priority": 1,
    "style": {
        "font": {"bold": True, "color": "FF00B050"},
        "fill": {"pattern": "solid", "fg_color": "FFC6EFCE"}
    }
}]

# Highlight bottom 5 values
conditional_formats = [{
    "start_row": 2,
    "start_col": 2,
    "end_row": 100,
    "end_col": 2,
    "rule_type": "top10",
    "rank": 5,
    "bottom": True,  # Bottom 5
    "priority": 1,
    "style": {
        "font": {"bold": True, "color": "FFFF0000"},
        "fill": {"pattern": "solid", "fg_color": "FFFFC7CE"}
    }
}]
```

## 📊 Multiple Sheets

Create multi-sheet workbooks with full independent formatting per sheet. Each sheet supports **all features** from `write_sheet_arrow()` including tables, charts, images, conditional formatting, data validation, formulas, cell styles, and more.

### Basic Multi-Sheet
```python
import polars as pl
import jetxl as jet

df_sales = pl.DataFrame({"Product": ["A", "B"], "Revenue": [100, 200]})
df_costs = pl.DataFrame({"Product": ["A", "B"], "Cost": [50, 80]})
df_profit = pl.DataFrame({"Product": ["A", "B"], "Profit": [50, 120]})

sheets = [
    {
        "data": df_sales.to_arrow(),
        "name": "Sales",
        "auto_filter": True
    },
    {
        "data": df_costs.to_arrow(),
        "name": "Costs",
        "freeze_rows": 1
    },
    {
        "data": df_profit.to_arrow(),
        "name": "Profit",
        "styled_headers": True
    }
]

jet.write_sheets_arrow(
    sheets,
    "report.xlsx",
    num_threads=4  # Use 4 threads for parallel generation
)
```

### Independent Formatting Per Sheet

Each sheet can have completely different formatting:
```python
sheets = [
    {
        "data": df_sales.to_arrow(),
        "name": "Sales",
        "styled_headers": True,
        "auto_filter": True,
        "freeze_rows": 1,
        "column_formats": {
            "Date": "date",
            "Revenue": "currency",
            "Tax": "percentage"
        },
        "tables": [{
            "name": "SalesTable",
            "start_row": 1,
            "start_col": 0,
            "end_row": 100,
            "end_col": 5,
            "style": "TableStyleMedium2"
        }],
        "tab_color": "FF00B050"  # Green tab
    },
    {
        "data": df_costs.to_arrow(),
        "name": "Costs",
        "auto_width": True,
        "conditional_formats": [{
            "start_row": 2,
            "start_col": 2,
            "end_row": 100,
            "end_col": 2,
            "rule_type": "data_bar",
            "color": "FFFF0000",
            "show_value": True,
            "priority": 1
        }],
        "tab_color": "FFFF0000"  # Red tab
    },
    {
        "data": df_profit.to_arrow(),
        "name": "Profit",
        "write_header_row": False,  # Data only, no headers
        "hidden_columns": [2, 3],
        "gridlines_visible": False,
        "zoom_scale": 150
    }
]

jet.write_sheets_arrow(sheets, "advanced.xlsx", num_threads=3)
```

### Multi-Sheet with Charts, Tables, and Images
```python
sheets = [
    {
        "data": df_monthly.to_arrow(),
        "name": "Monthly Sales",
        "styled_headers": True,
        "freeze_rows": 1,
        
        # Excel table
        "tables": [{
            "name": "MonthlySales",
            "start_row": 1,
            "start_col": 0,
            "end_row": 12,
            "end_col": 3,
            "style": "TableStyleMedium9"
        }],
        
        # Chart
        "charts": [{
            "chart_type": "column",
            "start_row": 1,
            "start_col": 0,
            "end_row": 12,
            "end_col": 2,
            "from_col": 5,
            "from_row": 1,
            "to_col": 13,
            "to_row": 16,
            "title": "Monthly Sales Trend",
            "category_col": 0,
            "x_axis_title": "Month",
            "y_axis_title": "Revenue ($)"
        }],
        
        # Logo
        "images": [{
            "path": "company_logo.png",
            "from_col": 0,
            "from_row": 0,
            "to_col": 2,
            "to_row": 4
        }]
    },
    {
        "data": df_quarterly.to_arrow(),
        "name": "Quarterly",
        "auto_filter": True,
        
        # Different chart type
        "charts": [{
            "chart_type": "pie",
            "start_row": 1,
            "start_col": 0,
            "end_row": 4,
            "end_col": 1,
            "from_col": 3,
            "from_row": 1,
            "to_col": 10,
            "to_row": 15,
            "title": "Market Share"
        }]
    }
]

jet.write_sheets_arrow(sheets, "dashboard.xlsx", num_threads=2)
```

### All Features Per Sheet

Every sheet supports the full API from `write_sheet_arrow()`:
```python
sheets = [
    {
        "data": df.to_arrow(),
        "name": "Complete Example",
        
        # Basic formatting
        "auto_filter": True,
        "freeze_rows": 1,
        "freeze_cols": 0,
        "auto_width": True,
        "styled_headers": True,
        "write_header_row": True,
        
        # Column formatting
        "column_widths": {"Name": 25.0, "Email": "200px", "Notes": "auto"},
        "column_formats": {"Date": "date", "Amount": "currency", "Rate": "percentage"},
        
        # Cell operations
        "merge_cells": [(1, 0, 1, 3), (5, 0, 8, 0)],
        "row_heights": {1: 30.0, 5: 25.0},
        
        # Cell styles
        "cell_styles": [{
            "row": 2,
            "col": 0,
            "font": {"bold": True, "color": "FFFF0000", "size": 14.0},
            "fill": {"pattern": "solid", "fg_color": "FFFFFF00"},
            "alignment": {"horizontal": "center", "vertical": "center"}
        }],
        
        # Data validation
        "data_validations": [{
            "start_row": 2, "start_col": 4,
            "end_row": 100, "end_col": 4,
            "type": "list",
            "items": ["Active", "Pending", "Closed"],
            "show_dropdown": True
        }],
        
        # Hyperlinks
        "hyperlinks": [(2, 0, "https://example.com", "Visit Site")],
        
        # Formulas
        "formulas": [(5, 5, "=SUM(A2:A4)", None)],
        
        # Conditional formatting
        "conditional_formats": [{
            "start_row": 2, "start_col": 3,
            "end_row": 100, "end_col": 3,
            "rule_type": "color_scale",
            "min_color": "FFF8696B",
            "mid_color": "FFFFEB84",
            "max_color": "FF63BE7B",
            "priority": 1
        }],
        
        # Excel tables
        "tables": [{
            "name": "DataTable",
            "start_row": 1, "start_col": 0,
            "end_row": 100, "end_col": 5,
            "style": "TableStyleMedium2"
        }],
        
        # Charts
        "charts": [{
            "chart_type": "column",
            "start_row": 1, "start_col": 0,
            "end_row": 12, "end_col": 2,
            "from_col": 7, "from_row": 1,
            "to_col": 15, "to_row": 18,
            "title": "Sales Chart"
        }],
        
        # Images
        "images": [{
            "path": "logo.png",
            "from_col": 0, "from_row": 0,
            "to_col": 2, "to_row": 4
        }],
        
        # Appearance
        "gridlines_visible": False,
        "zoom_scale": 120,
        "tab_color": "FF4472C4",
        "default_row_height": 18.0,
        "hidden_columns": [2],
        "hidden_rows": [5, 6],
        "right_to_left": False,
        "data_start_row": 0
    }
]

jet.write_sheets_arrow(sheets, "everything.xlsx", num_threads=1)
```

**Performance Notes:**
- XML generation is fully parallel across `num_threads`
- Each sheet can have independent formatting with minimal overhead (<1%)
- Style registry is shared for deduplication
- Recommended: `num_threads = min(cpu_count, len(sheets))`

## 🎨 Sheet Appearance & Layout

### Gridlines and Zoom

Control worksheet visibility settings:
```python
import polars as pl
import jetxl as jet

df = pl.DataFrame({
    "Product": ["A", "B", "C"],
    "Price": [10.0, 20.0, 30.0]
})

# Hide gridlines and set zoom
jet.write_sheet_arrow(
    df.to_arrow(),
    "clean_view.xlsx",
    gridlines_visible=False,  # Hide gridlines for cleaner look
    zoom_scale=150            # Zoom to 150% (range: 10-400)
)
```

### Sheet Tab Colors

Color-code your sheets for better organization:
```python
# Single sheet with colored tab
jet.write_sheet_arrow(
    df.to_arrow(),
    "colored_tab.xlsx",
    tab_color="FFFF0000"  # Red tab (ARGB format)
)

# Multiple sheets with different colors
sheets = [
    {
        "data": df_sales.to_arrow(),
        "name": "Sales",
        "tab_color": "FF00B050"  # Green
    },
    {
        "data": df_costs.to_arrow(),
        "name": "Costs",
        "tab_color": "FFFF0000"  # Red
    },
    {
        "data": df_profit.to_arrow(),
        "name": "Profit",
        "tab_color": "FF0070C0"  # Blue
    }
]

jet.write_sheets_arrow(sheets, "colored_tabs.xlsx", num_threads=2)
```

**Common Tab Colors:**
- `"FF4472C4"` - Blue
- `"FF00B050"` - Green
- `"FFFF0000"` - Red
- `"FFFFC000"` - Orange
- `"FF7030A0"` - Purple

### Default Row Height

Set a consistent row height for all rows:
```python
jet.write_sheet_arrow(
    df.to_arrow(),
    "tall_rows.xlsx",
    default_row_height=25.0,  # 25 points (default is 15)
    row_heights={
        1: 35.0,  # Override: make header taller
        5: 20.0   # Override: specific row
    }
)
```

### Hidden Rows and Columns

Hide sensitive or intermediate data:
```python
df = pl.DataFrame({
    "ID": [1, 2, 3],
    "Name": ["Alice", "Bob", "Charlie"],
    "Secret": ["X", "Y", "Z"],
    "Salary": [50000, 60000, 75000],
    "Bonus": [5000, 6000, 7500]
})

jet.write_sheet_arrow(
    df.to_arrow(),
    "hidden_data.xlsx",
    hidden_columns=[2, 4],  # Hide "Secret" (col 2) and "Bonus" (col 4)
    hidden_rows=[3]         # Hide row 3
)
```

**Note:** Hidden data is still in the file - it's just not visible by default. Users can unhide it in Excel.

### Right-to-Left Layout

For languages like Arabic, Hebrew, Persian, etc.:
```python
df = pl.DataFrame({
    "שם": ["אליס", "בוב", "צ'רלי"],
    "גיל": [25, 30, 35]
})

jet.write_sheet_arrow(
    df.to_arrow(),
    "hebrew.xlsx",
    right_to_left=True  # Sheet flows from right to left
)
```

### Auto-Width with Complex Headers

When your Excel file has multiple header rows, dummy rows, or template rows, exclude them from width calculation:
```python
# Scenario: Your file structure is:
# Row 1: Company logo (merged cells with long text)
# Row 2: Report title "Q4 2024 Financial Summary - Confidential"
# Row 3: Date range
# Row 4: Empty spacing row
# Row 5: Column headers (Name, Amount, Status)
# Row 6+: Actual data

# Without data_start_row, auto_width uses ALL rows including dummy rows
# This makes columns unnecessarily wide to fit the title text

jet.write_sheet_arrow(
    df.to_arrow(),
    "complex_report.xlsx",
    auto_width=True,
    data_start_row=5  # Start width calculation from row 5 (actual data)
    # Now columns are sized based on data + headers only
)
```

**Common use cases:**
- Reports with title rows, logos, or metadata at the top
- Templates with pre-existing formatting rows
- Multi-section reports where only one section should determine width
- Files with merged header rows that contain long text

### Header Content (Template Rows)

Write arbitrary content above your DataFrame data - perfect for report titles, metadata, logos in merged cells, or template headers:
```python
import polars as pl
import jetxl as jet

df = pl.DataFrame({
    "Name": ["Alice", "Bob"],
    "Sales": [1000, 1500]
})

# Add title rows, metadata, spacing before DataFrame
jet.write_sheet_arrow(
    df.to_arrow(),
    "report.xlsx",
    header_content=[
        (1, 0, "ACME Corporation"),           # Row 1, Col A
        (1, 2, "Confidential"),               # Row 1, Col C
        (2, 0, "Q4 2024 Sales Report"),       # Row 2, Col A
        (3, 0, "Generated: 2024-10-17"),      # Row 3, Col A
        # Row 4 is empty (spacing)
    ],
    data_start_row=5,  # DataFrame starts at row 5
    write_header_row=True,  # Row 5 will have column headers
    # Actual data starts at row 6
    merge_cells=[
        (1, 0, 1, 1),  # Merge A1:B1 for company name
    ]
)
```

**Common use cases:**
- Report headers with company name, logo placeholder, dates
- Multi-line titles with merged cells
- Metadata rows (author, generated date, version)
- Template text that shouldn't come from DataFrame
- Section dividers in complex reports

**Coordinates:**
- Row numbers are 1-based (row 1 is first row)
- Column numbers are 0-based (0=A, 1=B, 2=C, etc.)
- `header_content` rows are written BEFORE DataFrame data
- Use `data_start_row` to position DataFrame below header content




### Professional Dashboard Example

Combine appearance settings for a polished look:
```python
import polars as pl
import jetxl as jet

df = pl.DataFrame({
    "Quarter": ["Q1", "Q2", "Q3", "Q4"],
    "Revenue": [100000, 120000, 115000, 140000],
    "Target": [95000, 110000, 120000, 135000]
})

jet.write_sheet_arrow(
    df.to_arrow(),
    "executive_dashboard.xlsx",
    sheet_name="Performance",
    
    # Clean appearance
    gridlines_visible=False,
    zoom_scale=120,
    tab_color="FF0070C0",
    default_row_height=20.0,
    
    # Formatting
    styled_headers=True,
    freeze_rows=1,
    auto_width=True,
    column_formats={
        "Revenue": "currency",
        "Target": "currency"
    },
    
    # Visualization
    charts=[{
        "chart_type": "column",
        "start_row": 1,
        "start_col": 0,
        "end_row": 4,
        "end_col": 2,
        "from_col": 4,
        "from_row": 1,
        "to_col": 12,
        "to_row": 18,
        "title": "Revenue vs Target",
        "category_col": 0,
        "x_axis_title": "Quarter",
        "y_axis_title": "Amount ($)"
    }]
)
```




## 📋 Complete Example

Here's a comprehensive example using multiple features:

```python
import polars as pl
import jetxl as jet

# Create sample data
df = pl.DataFrame({
    "Date": ["2024-01-01", "2024-01-02", "2024-01-03"],
    "Product": ["Widget A", "Widget B", "Widget C"],
    "Quantity": [100, 150, 75],
    "Price": [19.99, 29.99, 39.99],
    "Revenue": [1999.0, 4498.5, 2999.25]
})

# Tables auto-size to data
tables = [{
    "name": "SalesData",
    "display_name": "Q1 Sales",
    "start_row": 1,
    "start_col": 0,
    "end_row": 0,      # Auto-calculate from DataFrame rows
    "end_col": 0,      # Auto-calculate from DataFrame columns
    "style": "TableStyleMedium9",
    "show_row_stripes": True
}]

# Add conditional formatting
conditional_formats = [{
    "start_row": 2,
    "start_col": 4,
    "end_row": 4,
    "end_col": 4,
    "rule_type": "data_bar",
    "color": "FF638EC6",
    "show_value": True,
    "priority": 1
}]

# Add chart
charts = [{
    "chart_type": "column",
    "start_row": 1,
    "start_col": 0,
    "end_row": 4,
    "end_col": 4,
    "from_col": 6,
    "from_row": 1,
    "to_col": 14,
    "to_row": 18,
    "title": "Revenue by Product",
    "category_col": 1,  # Product column
    "x_axis_title": "Product",
    "y_axis_title": "Revenue ($)",
    "show_legend": False
}]

# Add logo image
images = [{
    "path": "company_logo.png",
    "from_col": 0,
    "from_row": 0,
    "to_col": 2,
    "to_row": 3
}]

# Write to Excel
jet.write_sheet_arrow(
    df.to_arrow(),
    "sales_report.xlsx",
    sheet_name="Q1 Sales",
    styled_headers=True,
    freeze_rows=1,
    auto_width=True,
    column_formats={
        "Date": "date",
        "Price": "currency",
        "Revenue": "currency"
    },
    tables=tables,
    conditional_formats=conditional_formats,
    charts=charts,
    images=images
)
```



## 🗃️ Architecture

Jetxl achieves its performance through several key optimizations:

1. **Zero-copy Arrow Integration**: Direct memory access to DataFrame buffers without copying
2. **SIMD XML Escaping**: Hardware-accelerated string processing
3. **Pre-calculated Buffer Sizing**: Single allocation per sheet with exact size calculation
4. **Parallel Sheet Generation**: Multi-threaded XML generation for multiple sheets
5. **Optimized Number Formatting**: Fast integer/float detection and conversion
6. **Streaming Compression**: On-the-fly ZIP compression with minimal memory overhead

## 🔧 Advanced Usage

### Working with Large Datasets

```python
import polars as pl
import jetxl as jet

# For very large datasets, use batched reading
df = pl.scan_csv("huge_file.csv").collect()

# Jetxl handles large datasets efficiently
jet.write_sheet_arrow(
    df.to_arrow(),
    "large_output.xlsx",
    auto_width=False  # Disable auto-width for faster generation
)
```

### Custom Styling Templates

```python
def create_report_style():
    return {
        "styled_headers": True,
        "auto_filter": True,
        "freeze_rows": 1,
        "column_formats": {
            "Date": "date",
            "Amount": "currency",
            "Percentage": "percentage"
        }
    }

# Apply consistent styling across reports
jet.write_sheet_arrow(df.to_arrow(), "report.xlsx", **create_report_style())
```

### Error Handling

```python
try:
    jet.write_sheet_arrow(df.to_arrow(), "output.xlsx")
except IOError as e:
    print(f"Failed to write file: {e}")
except ValueError as e:
    print(f"Invalid data: {e}")
```

## 🤝 Comparison with Other Libraries

### vs xlsxwriter
- ✅ **5x faster** (1M rows: 2.06s vs 10.05s)
- ✅ Near-zero Python memory overhead
- ✅ Zero-copy DataFrame integration
- ✅ Multi-threaded sheet generation
- ✅ Modern Python API with type hints
- ❌ Larger output files (less aggressive compression)
- ❌ Fewer advanced chart customizations

### vs openpyxl
- ✅ **27x faster** (1M rows: 2.06s vs 56.25s)
- ✅ Dramatically lower memory usage
- ✅ Native Arrow/Polars/Pandas support
- ❌ Write-only (openpyxl supports reading)
- ❌ Fewer cell-level features

### vs polars.write_excel
- ✅ **20x faster** (1M rows: 2.06s vs 40.85s)
- ✅ **2000x lower memory** (~0 MB vs 2.1 GB at 1M rows)
- ✅ More formatting options (conditional formatting, tables, charts)
- ✅ Multi-sheet threading support
- ❌ Requires `.to_arrow()` conversion

### vs pandas.to_excel
- ✅ **27-40x faster** depending on engine
- ✅ Direct Polars support (no pandas dependency)
- ✅ Richer formatting options
- ✅ Multi-threading support
- ✅ Dramatically lower memory footprint

### vs rustpy_xlsxwriter
- ✅ **5.5x faster** (1M rows: 2.06s vs 11.27s)
- ✅ Native Arrow support (no data conversion needed)
- ✅ More formatting options
- ✅ Multi-threaded sheet generation


## 📋 Supported Data Types

### Arrow/Polars Types

- Numeric: `Int8/16/32/64`, `UInt8/16/32/64`, `Float32/64`
- String: `Utf8`, `LargeUtf8`
- Boolean: `Bool`
- Temporal: `Date32/64`, `Timestamp` (all units), `Time32/64`

### Python Types (Dict API)

- `str`, `int`, `float`, `bool`, `datetime`, `None`

## 📚 External Resources & References

### Official Microsoft Documentation

**Excel Tables**
- [Format an Excel Table](https://support.microsoft.com/en-us/office/format-an-excel-table-6789619f-c889-495c-99c2-2f971c0e2370) - Complete guide with visual examples of all table styles
- [Overview of Excel Tables](https://support.microsoft.com/en-us/office/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c) - Features and capabilities
- [Using Structured References](https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e) - Advanced table formulas

**Excel Charts**
- [Available Chart Types in Office](https://support.microsoft.com/en-us/office/available-chart-types-in-office-a6187218-807e-4103-9e0a-27cdb19afb90) - Complete reference for all chart types
- [Create a Chart from Start to Finish](https://support.microsoft.com/en-us/office/create-a-chart-from-start-to-finish-0baf399e-dd61-4e18-8a73-b3fd5d5680c2) - Step-by-step guide

**Number Formats**
- [Excel Number Format Codes](https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68) - Complete reference for custom format codes
- [Available Number Formats](https://support.microsoft.com/en-us/office/available-number-formats-in-excel-0afe8f52-97db-41f1-b972-4b46e9f1e8d2) - Built-in format options
- [Custom Number Format Guidelines](https://support.microsoft.com/en-us/office/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) - Creating custom formats
- [Create Custom Number Formats](https://support.microsoft.com/en-us/office/create-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4) - Detailed tutorial

### Color Resources

**Understanding Excel Colors**
- Excel uses **ARGB format** for colors: `AARRGGBB` where:
  - `AA` = Alpha channel (transparency) - usually `FF` for fully opaque
  - `RR` = Red component (00-FF in hexadecimal)
  - `GG` = Green component (00-FF in hexadecimal)  
  - `BB` = Blue component (00-FF in hexadecimal)

**Example Colors**:
```python
"FFFF0000"  # Red (FF = opaque, FF0000 = red)
"FF00FF00"  # Green (FF = opaque, 00FF00 = green)
"FF0000FF"  # Blue (FF = opaque, 0000FF = blue)
"FFFFFF00"  # Yellow (red + green)
"FFFF00FF"  # Magenta (red + blue)
"FF00FFFF"  # Cyan (green + blue)
"FF000000"  # Black
"FFFFFFFF"  # White
```

**Common Conditional Formatting Colors**:
```python
# Red-Yellow-Green color scale (default Excel)
"FFF8696B"  # Red for low values
"FFFFEB84"  # Yellow for middle values
"FF63BE7B"  # Green for high values

# Data bar colors
"FF638EC6"  # Blue (Excel default data bar)
"FF5687C5"  # Dark blue
"FFFF6347"  # Tomato red
```

**Color Picker Tools**:
- [RapidTables RGB Color Picker](https://www.rapidtables.com/web/color/RGB_Color.html) - Interactive color selection with RGB/hex codes
- [W3Schools Color Picker](https://www.w3schools.com/colors/colors_picker.asp) - Simple online color chooser
- [Microsoft RGB Function](https://support.microsoft.com/en-us/office/rgb-function-aa04db19-fb8a-4f58-9ad6-71a1f5a43e94) - Excel's RGB function documentation

### Community Resources

**Tutorials & Guides**
- [ExcelJet Custom Number Formats](https://exceljet.net/articles/custom-number-formats) - Comprehensive formatting guide
- [Ablebits Excel Tables Guide](https://www.ablebits.com/office-addins-blog/excel-tables-styles/) - Advanced table formatting
- [W3Schools Excel Tutorial](https://www.w3schools.com/excel/) - Beginner-friendly Excel basics

---


Made with ❤️ and 🦀 by the Jetxl team