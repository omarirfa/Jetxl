"""Shared pieces of the benchmark: the test data, each library's native input,
each library's writer, and the read-back check.

Fairness rules (see README.md):
  * every library writes the same values: a header row plus N rows x 10 columns
  * every library gets its input in the form its users would already have it,
    built before the timer starts (that cost is recorded separately as convert_s)
  * every library uses its fastest documented way of writing a plain sheet
  * every file is read back and checked cell by cell; a writer that drops or
    changes data fails instead of looking fast
"""
from __future__ import annotations

import datetime as dt
import math

import numpy as np

SEED = 42

# column name -> kind. Three string columns with different repetition, three
# integer, two float, one date-only and one date-time column.
COLUMNS = {
    "customer": "str", "product": "str", "region": "str",
    "order_id": "int", "quantity": "int", "store": "int",
    "price": "float", "discount": "float",
    "order_date": "date", "shipped_at": "datetime",
}
HEADER = list(COLUMNS)


def make_source(rows: int) -> dict[str, np.ndarray]:
    """Deterministic source columns as numpy arrays (the neutral starting point)."""
    rng = np.random.default_rng(SEED)
    pools = {
        "customer": np.array([f"Customer {i:05d}" for i in range(10_000)]),
        "product": np.array([f"SKU-{i:04d}-{chr(65 + i % 26)}" for i in range(1_000)]),
        "region": np.array(["North", "South", "East", "West", "Central"]),
    }
    base = np.datetime64("2020-01-01T00:00:00", "us")
    days = rng.integers(0, 2_000, rows).astype("timedelta64[D]").astype("timedelta64[us]")
    secs = rng.integers(0, 86_400, rows).astype("timedelta64[s]").astype("timedelta64[us]")
    return {
        "customer": pools["customer"][rng.integers(0, 10_000, rows)],
        "product": pools["product"][rng.integers(0, 1_000, rows)],
        "region": pools["region"][rng.integers(0, 5, rows)],
        "order_id": np.arange(1, rows + 1, dtype=np.int64),
        "quantity": rng.integers(1, 500, rows, dtype=np.int64),
        "store": rng.integers(100, 999, rows, dtype=np.int64),
        "price": np.round(rng.uniform(0.5, 5_000, rows), 2),
        "discount": rng.random(rows) * 0.3,
        "order_date": base + days,
        "shipped_at": base + days + secs,
    }


def to_python_columns(src: dict[str, np.ndarray]) -> dict[str, list]:
    """Plain Python lists; datetime64[us] becomes datetime.datetime."""
    return {k: v.tolist() for k, v in src.items()}


def to_rows(src: dict[str, np.ndarray]) -> list[list]:
    """Header row followed by one list per data row (row-oriented writers)."""
    cols = [v.tolist() for v in src.values()]
    return [HEADER] + [list(r) for r in zip(*cols, strict=True)]


# ---------------------------------------------------------------------------
# Writers: name -> (packages for the version list, build_input, write)
# Imports happen inside, so each process only loads what that library needs.
# A writer either writes `path` and returns None, or returns the workbook as
# bytes (in-memory output). Entries in BYTES_OUTPUT are kept out of the README
# chart and tables, which compare writing to a file.
# ---------------------------------------------------------------------------

def _arrow_input(src):
    import pyarrow as pa
    return pa.table({k: pa.array(v) for k, v in src.items()})


def _polars_input(src):
    import polars as pl
    return pl.DataFrame(src)


def _pandas_input(src):
    import pandas as pd
    return pd.DataFrame(src)


def _w_jetxl_arrow(table, path):
    import jetxl
    jetxl.write_sheet_arrow(table, path)


def _w_jetxl_arrow_bytes(table, path):
    # in-memory output: returns the workbook instead of writing `path`; the
    # harness saves it afterwards, outside the timer, for the read-back check
    import jetxl
    return jetxl.write_sheet_arrow_to_bytes(table)


def _w_jetxl_dict(cols, path):
    import jetxl
    jetxl.write_sheet(cols, path)


def _w_rustpy(df, path):
    from rustpy_xlsxwriter import write_worksheet
    write_worksheet(df, path)


def _w_xlsxwriter(rows, path):
    import xlsxwriter
    # constant_memory streams rows to disk, so rows must be written in order
    wb = xlsxwriter.Workbook(path, {"constant_memory": True,
                                    "default_date_format": "yyyy-mm-dd hh:mm:ss"})
    ws = wb.add_worksheet()
    for i, row in enumerate(rows):
        ws.write_row(i, 0, row)
    wb.close()


def _w_openpyxl(rows, path):
    from openpyxl import Workbook
    wb = Workbook(write_only=True)
    ws = wb.create_sheet()
    for row in rows:
        ws.append(row)
    wb.save(path)


def _w_pyexcelerate(rows, path):
    from pyexcelerate import Format, Style, Workbook
    wb = Workbook()
    ws = wb.new_sheet("Sheet1", data=rows)
    # the bulk data= path writes datetimes as bare serial numbers; a column
    # style is pyexcelerate's documented way to make them display as dates
    date_style = Style(format=Format("yyyy-mm-dd hh:mm:ss"))
    for i, name in enumerate(HEADER, start=1):
        if COLUMNS[name] in ("date", "datetime"):
            ws.set_col_style(i, date_style)
    wb.save(path)


def _w_polars(df, path):
    # polars always writes an Excel table object; turn off the parts it allows
    df.write_excel(path, autofilter=False)


def _w_pandas_openpyxl(df, path):
    df.to_excel(path, index=False, engine="openpyxl")


def _w_pandas_xlsxwriter(df, path):
    df.to_excel(path, index=False, engine="xlsxwriter")


WRITERS = {
    "jetxl (arrow)": (["jetxl", "pyarrow"], _arrow_input, _w_jetxl_arrow),
    "jetxl (arrow, bytes)": (["jetxl", "pyarrow"], _arrow_input, _w_jetxl_arrow_bytes),
    "jetxl (dict)": (["jetxl"], to_python_columns, _w_jetxl_dict),
    "rustpy_xlsxwriter": (["rustpy-xlsxwriter", "polars"], _polars_input, _w_rustpy),
    "xlsxwriter": (["xlsxwriter"], to_rows, _w_xlsxwriter),
    "polars.write_excel": (["polars", "xlsxwriter"], _polars_input, _w_polars),
    "pyexcelerate": (["pyexcelerate"], to_rows, _w_pyexcelerate),
    "openpyxl": (["openpyxl"], to_rows, _w_openpyxl),
    "pandas+xlsxwriter": (["pandas", "xlsxwriter"], _pandas_input, _w_pandas_xlsxwriter),
    "pandas+openpyxl": (["pandas", "openpyxl"], _pandas_input, _w_pandas_openpyxl),
}
BYTES_OUTPUT = {"jetxl (arrow, bytes)"}


# ---------------------------------------------------------------------------
# Read-back check
# ---------------------------------------------------------------------------

def _as_datetime(v):
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime(v.year, v.month, v.day)
    return None


def check_file(path: str, src: dict[str, np.ndarray]) -> dict:
    """Read the whole file back and compare every cell with the source data."""
    from python_calamine import CalamineWorkbook

    rows = CalamineWorkbook.from_path(path).get_sheet_by_index(0).to_python()
    expected = to_python_columns(src)
    n = len(next(iter(expected.values())))
    result = {"ok": False, "cells": 0, "error": None, "max_float_rel_err": 0.0}

    if not rows or list(rows[0][:len(HEADER)]) != HEADER:
        result["error"] = f"header mismatch: {rows[0] if rows else None}"
        return result
    if len(rows) != n + 1:
        result["error"] = f"expected {n} data rows, found {len(rows) - 1}"
        return result

    max_rel = 0.0
    for c, name in enumerate(HEADER):
        kind = COLUMNS[name]
        exp = expected[name]
        for r in range(n):
            got = rows[r + 1][c] if c < len(rows[r + 1]) else None
            want = exp[r]
            if kind == "str":
                good = got == want
            elif kind == "int":
                good = isinstance(got, (int, float)) and float(got) == float(want)
            elif kind == "float":
                good = isinstance(got, (int, float))
                if good:
                    rel = abs(got - want) / max(abs(want), 1e-300)
                    max_rel = max(max_rel, rel)
                    good = rel <= 1e-12 or math.isclose(got, want, abs_tol=1e-12)
            else:  # date / datetime must come back as dates, not bare serial numbers
                g = _as_datetime(got)
                good = g is not None and abs((g - want).total_seconds()) <= 1
            if not good:
                result["error"] = f"row {r + 2}, column {name!r}: expected {want!r}, got {got!r}"
                result["cells"] = c * n + r
                return result
    result.update(ok=True, cells=n * len(HEADER), max_float_rel_err=max_rel)
    return result
