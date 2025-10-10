# """
# JetXL - Fast Excel writer using Arrow and Rust

# Examples:
#     Basic usage:
#     >>> import polars as pl
#     >>> import jetxl
#     >>> df = pl.DataFrame({"Name": ["Alice", "Bob"], "Age": [25, 30]})
#     >>> jetxl.write_sheet_arrow(df.to_arrow(), "output.xlsx")

#     With formatting:
#     >>> jetxl.write_sheet_arrow(
#     ...     df.to_arrow(),
#     ...     "output.xlsx",
#     ...     sheet_name="Sales Data",
#     ...     auto_filter=True,
#     ...     freeze_rows=1,
#     ...     styled_headers=True,
#     ...     column_formats={"Age": "integer", "Salary": "currency"}
#     ... )
# """

# from typing import Any, Optional, Literal

# NumberFormat = Literal[
#     "general",
#     "integer",
#     "decimal2",
#     "decimal4",
#     "percentage",
#     "percentage_decimal",
#     "currency",
#     "currency_rounded",
#     "date",
#     "datetime",
#     "time",
# ]

# def write_sheet_arrow(
#     arrow_data: Any,
#     filename: str,
#     sheet_name: Optional[str] = None,
#     auto_filter: bool = False,
#     freeze_rows: int = 0,
#     freeze_cols: int = 0,
#     auto_width: bool = False,
#     styled_headers: bool = False,
#     column_widths: Optional[dict[str, float]] = None,
#     column_formats: Optional[dict[str, NumberFormat]] = None,
#     merge_cells: Optional[list[tuple[int, int, int, int]]] = None,
#     data_validations: Optional[list[dict[str, Any]]] = None,
#     hyperlinks: Optional[list[tuple[int, int, str, Optional[str]]]] = None,
#     row_heights: Optional[dict[int, float]] = None,
#     cell_styles: Optional[list[dict[str, Any]]] = None,
#     formulas: Optional[list[tuple[int, int, str, Optional[str]]]] = None,
#     conditional_formats: Optional[list[dict[str, Any]]] = None,
#     tables: Optional[list[dict[str, Any]]] = None,
# ) -> None:
#     """
#     Write Arrow data to Excel with advanced formatting.

#     Args:
#         arrow_data: PyArrow Table or RecordBatch (from polars: df.to_arrow())
#         filename: Output Excel file path
#         sheet_name: Sheet name (default: "Sheet1")
#         auto_filter: Enable autofilter on header row
#         freeze_rows: Number of rows to freeze (e.g., 1 for header)
#         freeze_cols: Number of columns to freeze
#         auto_width: Auto-calculate column widths based on content
#         styled_headers: Apply bold + gray background to headers
#         column_widths: Manual widths by column name, e.g., {"Name": 20.0, "Age": 10.0}
#         column_formats: Number formats by column, e.g., {"Price": "currency", "Date": "date"}
#         merge_cells: List of (start_row, start_col, end_row, end_col), e.g., [(1, 0, 1, 2)]
#         data_validations: Validation rules (see examples)
#         hyperlinks: List of (row, col, url, display_text), e.g., [(2, 0, "https://example.com", "Click here")]
#         row_heights: Custom heights by row number, e.g., {1: 30.0, 5: 25.0}
#         cell_styles: Custom styles (see examples)
#         formulas: List of (row, col, formula, cached_value), e.g., [(2, 3, "=A2+B2", "100")]
#         conditional_formats: Conditional formatting rules (see examples)
#         tables: Excel table definitions (see examples)

#     Examples:
#         Basic write:
#         >>> import polars as pl
#         >>> df = pl.DataFrame({"Name": ["Alice"], "Age": [25]})
#         >>> write_sheet_arrow(df.to_arrow(), "out.xlsx")

#         With formatting:
#         >>> write_sheet_arrow(
#         ...     df.to_arrow(),
#         ...     "out.xlsx",
#         ...     sheet_name="Report",
#         ...     auto_filter=True,
#         ...     freeze_rows=1,
#         ...     styled_headers=True,
#         ...     column_widths={"Name": 15.0, "Age": 8.0},
#         ...     column_formats={"Age": "integer"}
#         ... )

#         With hyperlinks (row/col are 1-based):
#         >>> write_sheet_arrow(
#         ...     df.to_arrow(),
#         ...     "out.xlsx",
#         ...     hyperlinks=[
#         ...         (2, 1, "https://example.com", "Website"),
#         ...         (3, 1, "mailto:test@example.com", None)
#         ...     ]
#         ... )

#         With data validation (dropdown):
#         >>> write_sheet_arrow(
#         ...     df.to_arrow(),
#         ...     "out.xlsx",
#         ...     data_validations=[{
#         ...         "start_row": 2, "start_col": 0,
#         ...         "end_row": 10, "end_col": 0,
#         ...         "type": "list",
#         ...         "items": ["Yes", "No", "Maybe"],
#         ...         "show_dropdown": True,
#         ...         "error_title": "Invalid",
#         ...         "error_message": "Pick from list"
#         ...     }]
#         ... )

#         With cell styles:
#         >>> write_sheet_arrow(
#         ...     df.to_arrow(),
#         ...     "out.xlsx",
#         ...     cell_styles=[{
#         ...         "row": 2, "col": 0,
#         ...         "font": {"bold": True, "color": "FFFF0000", "size": 14},
#         ...         "fill": {"pattern": "solid", "fg_color": "FFFFFF00"},
#         ...         "alignment": {"horizontal": "center", "vertical": "center"}
#         ...     }]
#         ... )

#         With formulas (row/col are 1-based):
#         >>> write_sheet_arrow(
#         ...     df.to_arrow(),
#         ...     "out.xlsx",
#         ...     formulas=[
#         ...         (2, 3, "=A2+B2", None),  # A2+B2 in cell C2
#         ...         (3, 3, "=SUM(A2:B2)", "100")  # With cached value
#         ...     ]
#         ... )

#         With conditional formatting (highlight cells > 100):
#         >>> write_sheet_arrow(
#         ...     df.to_arrow(),
#         ...     "out.xlsx",
#         ...     conditional_formats=[{
#         ...         "start_row": 2, "start_col": 0,
#         ...         "end_row": 10, "end_col": 0,
#         ...         "rule_type": "cell_value",
#         ...         "operator": "greater_than",
#         ...         "value": "100",
#         ...         "priority": 1,
#         ...         "style": {
#         ...             "font": {"bold": True, "color": "FFFF0000"}
#         ...         }
#         ...     }]
#         ... )

#         With Excel tables:
#         >>> write_sheet_arrow(
#         ...     df.to_arrow(),
#         ...     "out.xlsx",
#         ...     tables=[{
#         ...         "name": "Table1",
#         ...         "start_row": 1, "start_col": 0,
#         ...         "end_row": 10, "end_col": 2,
#         ...         "style": "TableStyleMedium2",
#         ...         "show_row_stripes": True
#         ...     }]
#         ... )
#     """
#     ...

# def write_sheets_arrow(
#     arrow_sheets: list[tuple[Any, str]],
#     filename: str,
#     num_threads: int,
# ) -> None:
#     """
#     Write multiple Arrow tables to Excel sheets in parallel.

#     Args:
#         arrow_sheets: List of (arrow_data, sheet_name) tuples
#         filename: Output Excel file path
#         num_threads: Number of parallel threads for XML generation

#     Example:
#         >>> import polars as pl
#         >>> df1 = pl.DataFrame({"A": [1, 2]})
#         >>> df2 = pl.DataFrame({"B": [3, 4]})
#         >>> write_sheets_arrow(
#         ...     [(df1.to_arrow(), "Sheet1"), (df2.to_arrow(), "Sheet2")],
#         ...     "output.xlsx",
#         ...     num_threads=4
#         ... )
#     """
#     ...

# def write_sheet(
#     columns: dict[str, list[Any]],
#     filename: str,
#     sheet_name: Optional[str] = None,
# ) -> None:
#     """
#     Write dict-based data to Excel (legacy API, slower than Arrow).

#     Args:
#         columns: Dictionary mapping column names to lists of values
#         filename: Output Excel file path
#         sheet_name: Sheet name (default: "Sheet1")

#     Example:
#         >>> write_sheet(
#         ...     {"Name": ["Alice", "Bob"], "Age": [25, 30]},
#         ...     "output.xlsx",
#         ...     sheet_name="People"
#         ... )
#     """
#     ...

# def write_sheets(
#     sheets_data: list[dict[str, Any]],
#     filename: str,
#     num_threads: int,
# ) -> None:
#     """
#     Write multiple dict-based sheets to Excel (legacy API).

#     Args:
#         sheets_data: List of dicts with "name" and "columns" keys
#         filename: Output Excel file path
#         num_threads: Number of parallel threads

#     Example:
#         >>> write_sheets(
#         ...     [
#         ...         {"name": "Sheet1", "columns": {"A": [1, 2]}},
#         ...         {"name": "Sheet2", "columns": {"B": [3, 4]}}
#         ...     ],
#         ...     "output.xlsx",
#         ...     num_threads=2
#         ... )
#     """
#     ...
"""
JetXL - Fast Excel writer using Arrow and Rust

Examples:
    Basic usage:
    >>> import polars as pl
    >>> import jetxl
    >>> df = pl.DataFrame({"Name": ["Alice", "Bob"], "Age": [25, 30]})
    >>> jetxl.write_sheet_arrow(df.to_arrow(), "output.xlsx")

    With formatting:
    >>> jetxl.write_sheet_arrow(
    ...     df.to_arrow(),
    ...     "output.xlsx",
    ...     sheet_name="Sales Data",
    ...     auto_filter=True,
    ...     freeze_rows=1,
    ...     styled_headers=True,
    ...     column_formats={"Age": "integer", "Salary": "currency"}
    ... )
"""

from typing import Any, Optional, Literal, TypedDict

NumberFormat = Literal[
    "general",
    "integer",
    "decimal2",
    "decimal4",
    "percentage",
    "percentage_decimal",
    "currency",
    "currency_rounded",
    "date",
    "datetime",
    "time",
]

class FontStyle(TypedDict, total=False):
    """Font styling options."""
    bold: bool
    italic: bool
    underline: bool
    size: float
    color: str  # ARGB hex: "FFFF0000" for red
    name: str

class FillStyle(TypedDict, total=False):
    """Cell fill/background styling."""
    pattern: Literal["solid", "gray125", "none"]
    fg_color: str  # ARGB hex
    bg_color: str  # ARGB hex

class BorderSide(TypedDict, total=False):
    """Single border side styling."""
    style: Literal["thin", "medium", "thick", "double", "dotted", "dashed"]
    color: str  # ARGB hex

class BorderStyle(TypedDict, total=False):
    """Cell border styling."""
    left: BorderSide
    right: BorderSide
    top: BorderSide
    bottom: BorderSide

class AlignmentStyle(TypedDict, total=False):
    """Cell alignment options."""
    horizontal: Literal["left", "center", "right", "justify"]
    vertical: Literal["top", "center", "bottom"]
    wrap_text: bool
    text_rotation: int  # 0-180 degrees

class CellStyle(TypedDict, total=False):
    """Complete cell styling."""
    font: FontStyle
    fill: FillStyle
    border: BorderStyle
    alignment: AlignmentStyle
    number_format: NumberFormat

class CellStyleMap(TypedDict):
    """Cell style with position."""
    row: int  # 1-based
    col: int  # 0-based
    font: FontStyle
    fill: FillStyle
    border: BorderStyle
    alignment: AlignmentStyle
    number_format: NumberFormat

class DataValidationList(TypedDict):
    """Dropdown list validation."""
    start_row: int  # 1-based
    start_col: int  # 0-based
    end_row: int
    end_col: int
    type: Literal["list"]
    items: list[str]
    show_dropdown: bool
    error_title: str
    error_message: str

class DataValidationNumber(TypedDict):
    """Number range validation."""
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
    """Text length validation."""
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

class ConditionalFormatCellValue(TypedDict):
    """Cell value conditional formatting."""
    start_row: int  # 1-based
    start_col: int  # 0-based
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
    """Color scale conditional formatting."""
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    rule_type: Literal["color_scale"]
    min_color: str  # ARGB hex
    max_color: str
    mid_color: str  # optional
    priority: int

class ConditionalFormatDataBar(TypedDict):
    """Data bar conditional formatting."""
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    rule_type: Literal["data_bar"]
    color: str  # ARGB hex
    show_value: bool
    priority: int

class ConditionalFormatTop10(TypedDict):
    """Top/Bottom N conditional formatting."""
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

class ExcelTable(TypedDict, total=False):
    """Excel table definition."""
    name: str  # Required
    start_row: int  # Required, 1-based
    start_col: int  # Required, 0-based
    end_row: int  # Required
    end_col: int  # Required
    display_name: str
    style: str  # "TableStyleMedium2", "TableStyleLight1", etc.
    show_first_column: bool
    show_last_column: bool
    show_row_stripes: bool
    show_column_stripes: bool

def write_sheet_arrow(
    arrow_data: Any,
    filename: str,
    sheet_name: Optional[str] = None,
    auto_filter: bool = False,
    freeze_rows: int = 0,
    freeze_cols: int = 0,
    auto_width: bool = False,
    styled_headers: bool = False,
    column_widths: Optional[dict[str, float]] = None,
    column_formats: Optional[dict[str, NumberFormat]] = None,
    merge_cells: Optional[list[tuple[int, int, int, int]]] = None,
    data_validations: Optional[list[DataValidation]] = None,
    hyperlinks: Optional[list[tuple[int, int, str, Optional[str]]]] = None,
    row_heights: Optional[dict[int, float]] = None,
    cell_styles: Optional[list[CellStyleMap]] = None,
    formulas: Optional[list[tuple[int, int, str, Optional[str]]]] = None,
    conditional_formats: Optional[list[ConditionalFormat]] = None,
    tables: Optional[list[ExcelTable]] = None,
) -> None:
    """
    Write Arrow data to Excel with advanced formatting.

    Args:
        arrow_data: PyArrow Table or RecordBatch (from polars: df.to_arrow())
        filename: Output Excel file path
        sheet_name: Sheet name (default: "Sheet1")
        auto_filter: Enable autofilter on header row
        freeze_rows: Number of rows to freeze (e.g., 1 for header)
        freeze_cols: Number of columns to freeze
        auto_width: Auto-calculate column widths based on content
        styled_headers: Apply bold + gray background to headers
        column_widths: Manual widths by column name, e.g., {"Name": 20.0, "Age": 10.0}
        column_formats: Number formats by column, e.g., {"Price": "currency", "Date": "date"}
        merge_cells: List of (start_row, start_col, end_row, end_col), e.g., [(1, 0, 1, 2)]
        data_validations: Validation rules
        hyperlinks: List of (row, col, url, display_text), e.g., [(2, 0, "https://example.com", "Click")]
        row_heights: Custom heights by row number, e.g., {1: 30.0, 5: 25.0}
        cell_styles: Custom styles with position
        formulas: List of (row, col, formula, cached_value), e.g., [(2, 3, "=A2+B2", "100")]
        conditional_formats: Conditional formatting rules
        tables: Excel table definitions

    Examples:
        Basic:
        >>> import polars as pl
        >>> df = pl.DataFrame({"Name": ["Alice"], "Age": [25]})
        >>> write_sheet_arrow(df.to_arrow(), "out.xlsx")

        Formatted:
        >>> write_sheet_arrow(
        ...     df.to_arrow(),
        ...     "out.xlsx",
        ...     sheet_name="Report",
        ...     auto_filter=True,
        ...     freeze_rows=1,
        ...     styled_headers=True,
        ...     column_widths={"Name": 15.0, "Age": 8.0},
        ...     column_formats={"Age": "integer"}
        ... )

        Hyperlinks (1-based rows, 0-based cols):
        >>> write_sheet_arrow(
        ...     df.to_arrow(),
        ...     "out.xlsx",
        ...     hyperlinks=[(2, 0, "https://example.com", "Website")]
        ... )

        Data validation dropdown:
        >>> write_sheet_arrow(
        ...     df.to_arrow(),
        ...     "out.xlsx",
        ...     data_validations=[{
        ...         "start_row": 2, "start_col": 0,
        ...         "end_row": 10, "end_col": 0,
        ...         "type": "list",
        ...         "items": ["Yes", "No", "Maybe"],
        ...         "show_dropdown": True,
        ...         "error_title": "Invalid",
        ...         "error_message": "Pick from list"
        ...     }]
        ... )

        Cell styles:
        >>> write_sheet_arrow(
        ...     df.to_arrow(),
        ...     "out.xlsx",
        ...     cell_styles=[{
        ...         "row": 2, "col": 0,
        ...         "font": {"bold": True, "color": "FFFF0000", "size": 14},
        ...         "fill": {"pattern": "solid", "fg_color": "FFFFFF00"},
        ...         "alignment": {"horizontal": "center", "vertical": "center"}
        ...     }]
        ... )

        Formulas (1-based rows, 0-based cols):
        >>> write_sheet_arrow(
        ...     df.to_arrow(),
        ...     "out.xlsx",
        ...     formulas=[
        ...         (2, 2, "=A2+B2", None),
        ...         (3, 2, "=SUM(A2:B2)", "100")
        ...     ]
        ... )

        Conditional formatting:
        >>> write_sheet_arrow(
        ...     df.to_arrow(),
        ...     "out.xlsx",
        ...     conditional_formats=[{
        ...         "start_row": 2, "start_col": 0,
        ...         "end_row": 10, "end_col": 0,
        ...         "rule_type": "cell_value",
        ...         "operator": "greater_than",
        ...         "value": "100",
        ...         "priority": 1,
        ...         "style": {"font": {"bold": True, "color": "FFFF0000"}}
        ...     }]
        ... )

        Excel tables:
        >>> write_sheet_arrow(
        ...     df.to_arrow(),
        ...     "out.xlsx",
        ...     tables=[{
        ...         "name": "Table1",
        ...         "start_row": 1, "start_col": 0,
        ...         "end_row": 10, "end_col": 2,
        ...         "style": "TableStyleMedium2",
        ...         "show_row_stripes": True
        ...     }]
        ... )
    """
    ...

def write_sheets_arrow(
    arrow_sheets: list[tuple[Any, str]],
    filename: str,
    num_threads: int,
) -> None:
    """
    Write multiple Arrow tables to Excel sheets in parallel.

    Args:
        arrow_sheets: List of (arrow_data, sheet_name) tuples
        filename: Output Excel file path
        num_threads: Number of parallel threads for XML generation

    Example:
        >>> import polars as pl
        >>> df1 = pl.DataFrame({"A": [1, 2]})
        >>> df2 = pl.DataFrame({"B": [3, 4]})
        >>> write_sheets_arrow(
        ...     [(df1.to_arrow(), "Sheet1"), (df2.to_arrow(), "Sheet2")],
        ...     "output.xlsx",
        ...     num_threads=4
        ... )
    """
    ...

def write_sheet(
    columns: dict[str, list[Any]],
    filename: str,
    sheet_name: Optional[str] = None,
) -> None:
    """
    Write dict-based data to Excel (legacy API, slower than Arrow).

    Args:
        columns: Dictionary mapping column names to lists of values
        filename: Output Excel file path
        sheet_name: Sheet name (default: "Sheet1")

    Example:
        >>> write_sheet(
        ...     {"Name": ["Alice", "Bob"], "Age": [25, 30]},
        ...     "output.xlsx",
        ...     sheet_name="People"
        ... )
    """
    ...

def write_sheets(
    sheets_data: list[dict[str, Any]],
    filename: str,
    num_threads: int,
) -> None:
    """
    Write multiple dict-based sheets to Excel (legacy API).

    Args:
        sheets_data: List of dicts with "name" and "columns" keys
        filename: Output Excel file path
        num_threads: Number of parallel threads

    Example:
        >>> write_sheets(
        ...     [
        ...         {"name": "Sheet1", "columns": {"A": [1, 2]}},
        ...         {"name": "Sheet2", "columns": {"B": [3, 4]}}
        ...     ],
        ...     "output.xlsx",
        ...     num_threads=2
        ... )
    """
    ...