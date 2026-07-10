#!/usr/bin/env python3
"""
jetxl property-based fuzzing harness
====================================

Generates thousands of randomized-but-valid workbook configurations and asserts
three properties on every one, across all write paths:

  P1  NO PANIC / NO CRASH  — the call returns bytes (or writes a file) and never
                             aborts the process or raises an unexpected error.
  P2  LOADABLE             — openpyxl opens the result with warnings-as-errors
                             (so a silently-discarded rule counts as a failure).
  P3  SCHEMA-CONFORMANT    — every worksheet/styles/table/chart/drawing part
                             validates against the official ECMA-376 XSDs, with
                             the Part-3 vendor-extension wildcard honored.

The generators cover: random Arrow schemas (many dtypes incl. nulls, unicode,
extreme values, categoricals), and random feature configs (cell_styles with
random colors/fonts/borders/alignment/rotation, merges, conditional formats,
data validations, tables, charts, images, hyperlinks, formulas, freeze panes,
column/row sizing, auto-filter). Coordinates are deliberately allowed to go
out of range sometimes, to probe clamping/validation rather than only happy
paths.

Determinism: a fixed seed (overridable with JETXL_FUZZ_SEED) makes any failure
reproducible. On failure the offending config is printed as a repro.

Usage:
    python fuzz_jetxl.py [N] [-v]         # N iterations (default 400)
    JETXL_FUZZ_SEED=123 python fuzz_jetxl.py 2000
"""
from __future__ import annotations

import io
import os
import random
import sys
import tempfile
import traceback
import warnings
import zipfile

warnings.filterwarnings("ignore")

import pyarrow as pa
import openpyxl

import jetxl

N = 400
for a in sys.argv[1:]:
    if a.isdigit():
        N = int(a)
VERBOSE = "-v" in sys.argv or "--verbose" in sys.argv
SEED = int(os.environ.get("JETXL_FUZZ_SEED", "20260705"))

# ---- optional schema validation ----------------------------------------
SCHEMA_DIR = os.environ.get("JETXL_ECMA_SCHEMAS", os.path.join(os.path.dirname(__file__), "ecma", "schemas_transitional"))
OPC_DIR = os.environ.get("JETXL_ECMA_OPC", os.path.join(os.path.dirname(__file__), "ecma", "schemas_opc"))
_schemas = None


def load_schemas():
    global _schemas
    if _schemas is not None:
        return _schemas
    try:
        import xmlschema
        if not os.path.isfile(os.path.join(SCHEMA_DIR, "sml.xsd")):
            _schemas = {}
            return _schemas
        _schemas = {
            "sml": xmlschema.XMLSchema(os.path.join(SCHEMA_DIR, "sml.xsd")),
            "chart": xmlschema.XMLSchema(os.path.join(SCHEMA_DIR, "dml-chart.xsd")),
            "drawing": xmlschema.XMLSchema(os.path.join(SCHEMA_DIR, "dml-spreadsheetDrawing.xsd")),
        }
    except Exception:
        _schemas = {}
    return _schemas


def schema_for(part):
    s = load_schemas()
    if not s:
        return None
    import re
    if re.match(r"xl/(worksheets/sheet|workbook|styles|sharedStrings|tables/table)", part):
        return s["sml"]
    if "charts/chart" in part and part.endswith(".xml"):
        return s["chart"]
    if "drawings/drawing" in part and part.endswith(".xml") and "rels" not in part:
        return s["drawing"]
    return None


def schema_violations(b):
    """Return list of (part, reason) for REAL violations.

    Two classes of validator false-positive are filtered out, because they are
    spec-conformant despite the strict transitional XSD flagging them:

      1. Content inside an extLst/ext wildcard — ECMA-376 Part 3 permits vendor
         extensions there via processContents="lax".
      2. The W3C-reserved xml: namespace attributes (xml:space, xml:lang) — these
         are allowed on ANY element by the XML 1.0 spec independent of the XSD.
         Excel/LibreOffice/openpyxl all emit <t xml:space="preserve"> on inline
         strings; the SpreadsheetML <t> (ST_Xstring) just doesn't model it. jetxl
         needs xml:space to preserve leading/trailing whitespace, so this is a
         schema-completeness gap, not a jetxl defect.
    """
    s = load_schemas()
    if not s:
        return []  # schemas unavailable -> P3 skipped
    out = []
    z = zipfile.ZipFile(io.BytesIO(b))
    for part in z.namelist():
        sch = schema_for(part)
        if sch is None:
            continue
        try:
            for e in sch.iter_errors(z.read(part).decode("utf-8")):
                path = (e.path or "").lower()
                reason = str(e.reason)
                if "ext" in path:
                    continue  # vendor-extension wildcard
                if "1998/namespace" in reason.lower() or "xml:space" in reason.lower() or "xml:lang" in reason.lower():
                    continue  # W3C-reserved xml: attribute
                out.append((part, reason[:80]))
                break
        except Exception as ex:
            out.append((part, f"validator error: {ex!r}"[:80]))
    return out


# ==========================================================================
# Generators
# ==========================================================================
PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d4948445200000001000000010806000000"
    "1f15c4890000000d49444154789c6260f8cf000000ffff03000006000557"
    "bfabd40000000049454e44ae426082")

UNICODE_SAMPLES = ["café", "日本語", "Ω≈ç√", "emoji😀", "a\tb", " lead", "trail ",
                   "", "   ", "quote\"here", "amp&<>", "\x01ctrl", "null\x00byte",
                   "A" * 300, "0007", "-42", "3.14"]

COLOR_SAMPLES = ["FF0000", "FFFF0000", "#00FF00", "#FF112233"[:7], "00ff00",
                 "red", "F00", "zzz", "", "12345678", "GGGGGG", "638EC6"]


def rng_string_array(rng, n):
    return pa.array([rng.choice(UNICODE_SAMPLES) if rng.random() > 0.15 else None
                     for _ in range(n)])


def rng_column(rng, n):
    """Return (arrow_array, kind) for a random dtype."""
    import datetime as dt
    kind = rng.choice([
        "int8", "int16", "int32", "int64", "uint8", "uint16", "uint32", "uint64",
        "float32", "float64", "bool", "string", "large_string",
        "date32", "timestamp_us", "timestamp_ns", "categorical", "all_null",
    ])
    def maybe_null(vals, ty):
        vals = [v if rng.random() > 0.1 else None for v in vals]
        return pa.array(vals, ty)
    if kind == "int8":
        return maybe_null([rng.randint(-128, 127) for _ in range(n)], pa.int8()), kind
    if kind == "int16":
        return maybe_null([rng.randint(-32768, 32767) for _ in range(n)], pa.int16()), kind
    if kind == "int32":
        return maybe_null([rng.randint(-2**31, 2**31 - 1) for _ in range(n)], pa.int32()), kind
    if kind == "int64":
        return maybe_null([rng.choice([0, -1, 2**62, -2**62, rng.randint(-10**9, 10**9)]) for _ in range(n)], pa.int64()), kind
    if kind == "uint8":
        return maybe_null([rng.randint(0, 255) for _ in range(n)], pa.uint8()), kind
    if kind == "uint16":
        return maybe_null([rng.randint(0, 65535) for _ in range(n)], pa.uint16()), kind
    if kind == "uint32":
        return maybe_null([rng.randint(0, 2**32 - 1) for _ in range(n)], pa.uint32()), kind
    if kind == "uint64":
        # deliberately include values above i64::MAX (regression guard for bug 14)
        return maybe_null([rng.choice([0, 2**63, 2**64 - 1, rng.randint(0, 2**64 - 1)]) for _ in range(n)], pa.uint64()), kind
    if kind == "float32":
        return maybe_null([rng.uniform(-1e6, 1e6) for _ in range(n)], pa.float32()), kind
    if kind == "float64":
        return maybe_null([rng.choice([0.0, -0.0, 1e308, -1e308, 1e-308, float("nan"), float("inf"), rng.uniform(-1e6, 1e6)]) for _ in range(n)], pa.float64()), kind
    if kind == "bool":
        return maybe_null([rng.random() > 0.5 for _ in range(n)], pa.bool_()), kind
    if kind == "string":
        return rng_string_array(rng, n), kind
    if kind == "large_string":
        return pa.array([rng.choice(UNICODE_SAMPLES) for _ in range(n)], pa.large_string()), kind
    if kind == "date32":
        return maybe_null([dt.date(rng.randint(1900, 2100), rng.randint(1, 12), rng.randint(1, 28)) for _ in range(n)], pa.date32()), kind
    if kind == "timestamp_us":
        return maybe_null([dt.datetime(rng.randint(1900, 2100), rng.randint(1, 12), rng.randint(1, 28), rng.randint(0, 23), rng.randint(0, 59)) for _ in range(n)], pa.timestamp("us")), kind
    if kind == "timestamp_ns":
        return maybe_null([dt.datetime(rng.randint(1970, 2100), rng.randint(1, 12), rng.randint(1, 28)) for _ in range(n)], pa.timestamp("ns")), kind
    if kind == "categorical":
        cats = [rng.choice(UNICODE_SAMPLES[:6]) for _ in range(n)]
        return pa.array(cats).dictionary_encode(), kind
    if kind == "all_null":
        return pa.array([None] * n, rng.choice([pa.int64(), pa.float64(), pa.string()])), kind
    return pa.array(list(range(n)), pa.int64()), "int64"


def rng_table(rng):
    ncols = rng.randint(1, 6)
    nrows = rng.choice([1, 1, 2, 5, 10, 25, 50])
    cols = {}
    for c in range(ncols):
        arr, _ = rng_column(rng, nrows)
        name = rng.choice(["col", "值", "a b", "x_1", "Field"]) + str(c)
        cols[name] = arr
    return pa.table(cols), ncols, nrows


def rng_color(rng):
    return rng.choice(COLOR_SAMPLES)


def rng_config(rng, ncols, nrows):
    """Random feature config. Coordinates sometimes exceed bounds on purpose."""
    cfg = {}
    last_row = nrows + 1  # 1-based incl header
    def rrow():  # sometimes out of range
        return rng.choice([rng.randint(1, last_row), rng.randint(1, last_row + 5)])
    def rcol():
        return rng.choice([rng.randint(0, ncols - 1), rng.randint(0, ncols + 3)])

    if rng.random() < 0.5:
        styles = []
        for _ in range(rng.randint(1, 4)):
            st = {"row": rrow(), "col": rcol()}
            if rng.random() < 0.7:
                st["font"] = {}
                if rng.random() < 0.5: st["font"]["bold"] = rng.random() < 0.5
                if rng.random() < 0.5: st["font"]["italic"] = rng.random() < 0.5
                if rng.random() < 0.5: st["font"]["size"] = rng.choice([0, 8, 11, 409, 1000])
                if rng.random() < 0.6: st["font"]["color"] = rng_color(rng)
                if rng.random() < 0.3: st["font"]["name"] = rng.choice(["Arial", "", "Calibri"])
            if rng.random() < 0.5:
                st["fill"] = {"pattern": "solid", "fg_color": rng_color(rng)}
            if rng.random() < 0.4:
                side = rng.choice(["left", "right", "top", "bottom"])
                st["border"] = {side: {"style": rng.choice(["thin", "thick", "medium", "dashed"])}}
            if rng.random() < 0.4:
                st["alignment"] = {}
                if rng.random() < 0.5: st["alignment"]["horizontal"] = rng.choice(["left", "center", "right"])
                if rng.random() < 0.5: st["alignment"]["text_rotation"] = rng.choice([0, 45, 90, -45, -90, 180, 200, 255])
            styles.append(st)
        cfg["cell_styles"] = styles

    if rng.random() < 0.3:
        cfg["merge_cells"] = [(rrow(), rcol(), rrow(), rcol()) for _ in range(rng.randint(1, 3))]

    if rng.random() < 0.35:
        cfs = []
        for _ in range(rng.randint(1, 3)):
            rt = rng.choice(["cell_value", "color_scale", "data_bar", "top10"])
            cf = {"start_row": rrow(), "start_col": rcol(), "end_row": rrow(), "end_col": rcol(), "rule_type": rt}
            if rt == "cell_value":
                cf["operator"] = rng.choice(["greater_than", "less_than", "equal", "between"])
                cf["value"] = str(rng.randint(0, 100))
                if cf["operator"] == "between": cf["value2"] = str(rng.randint(100, 200))
            elif rt == "color_scale":
                cf["min_color"] = rng_color(rng); cf["max_color"] = rng_color(rng)
                if rng.random() < 0.5: cf["mid_color"] = rng_color(rng)
            elif rt == "data_bar":
                cf["color"] = rng_color(rng)
            elif rt == "top10":
                cf["rank"] = rng.randint(1, 20); cf["bottom"] = rng.random() < 0.5
            cfs.append(cf)
        cfg["conditional_formats"] = cfs

    if rng.random() < 0.3:
        dvs = []
        for _ in range(rng.randint(1, 2)):
            vt = rng.choice(["list", "whole_number", "decimal"])
            dv = {"start_row": rrow(), "start_col": rcol(), "end_row": rrow(), "end_col": rcol(), "type": vt}
            if vt == "list":
                dv["items"] = [rng.choice(UNICODE_SAMPLES[:5]) for _ in range(rng.randint(1, 5))]
            else:
                dv["min"] = rng.randint(0, 50); dv["max"] = rng.randint(0, 50)
            dvs.append(dv)
        cfg["data_validations"] = dvs

    if rng.random() < 0.3:
        tables = []
        for i in range(rng.randint(1, 2)):
            tables.append({"name": rng.choice(["T", "", "My Table", "值", "D"]),
                           "start_row": 0, "start_col": rng.randint(0, max(0, ncols - 1)),
                           "end_row": rng.choice([nrows, nrows + 3]),
                           "end_col": rng.choice([ncols - 1, ncols + 2]),
                           "style": rng.choice(["TableStyleMedium9", "NotReal", "TableStyleLight1"])})
        cfg["tables"] = tables

    if rng.random() < 0.25:
        charts = []
        for _ in range(rng.randint(1, 2)):
            ct = rng.choice(["column", "bar", "line", "pie", "scatter", "area"])
            ch = {"chart_type": ct, "data_range": (0, min(1, ncols - 1), nrows, min(1, ncols - 1)),
                  "category_col": 0}
            if rng.random() < 0.6: ch["title"] = rng.choice(UNICODE_SAMPLES[:5])
            if rng.random() < 0.4: ch["title_color"] = rng_color(rng)
            if rng.random() < 0.4: ch["title_font_size"] = rng.choice([0, 900, 1400, 999999])
            if rng.random() < 0.4: ch["legend_position"] = rng.choice(["top", "bottom", "left", "right", "nowhere"])
            if rng.random() < 0.3: ch["y_axis_title"] = "Y"; ch["axis_title_color"] = rng_color(rng)
            if rng.random() < 0.3: ch["show_data_labels"] = True
            charts.append(ch)
        cfg["charts"] = charts

    if rng.random() < 0.2:
        cfg["images"] = [{"data": list(PNG), "extension": "png",
                          "from_col": rng.randint(0, ncols), "from_row": rng.randint(0, last_row),
                          "to_col": rng.randint(0, ncols + 3), "to_row": rng.randint(0, last_row + 3)}]

    if rng.random() < 0.3:
        cfg["hyperlinks"] = [(rrow(), rcol(), rng.choice(["https://x.com", "mailto:a@b.com", "#Sheet1!A1", ""]), "lnk")
                             for _ in range(rng.randint(1, 2))]

    if rng.random() < 0.2:
        cfg["formulas"] = [(rrow(), rcol(), rng.choice(["=A1+B1", "A1+B1", "=SUM(A1:A5)", "=IF(A1>0,1,0)"]), "0")
                           for _ in range(rng.randint(1, 2))]

    if rng.random() < 0.3:
        cfg["auto_filter"] = True
    if rng.random() < 0.3:
        cfg["freeze_rows"] = rng.choice([0, 1, 2, 100])
    if rng.random() < 0.2:
        cfg["freeze_cols"] = rng.choice([0, 1, 2])
    if rng.random() < 0.2:
        cfg["zoom_scale"] = rng.choice([10, 50, 100, 200, 400])
    return cfg


# ==========================================================================
def run_paths(data, cfg, rng):
    """Yield (path_label, bytes) for a randomly chosen subset of the 4 paths."""
    results = []
    # single bytes (always)
    results.append(("single_bytes", jetxl.write_sheet_arrow_to_bytes(data, **cfg)))
    if rng.random() < 0.5:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
            p = tf.name
        try:
            jetxl.write_sheet_arrow(data, p, **cfg)
            results.append(("single_file", open(p, "rb").read()))
        finally:
            os.unlink(p)
    if rng.random() < 0.6:
        sheet = {"data": data, "name": "S1"}; sheet.update(cfg)
        results.append(("multi_bytes", jetxl.write_sheets_arrow_to_bytes([sheet], 1)))
    if rng.random() < 0.4:
        sheet = {"data": data, "name": "S1"}; sheet.update(cfg)
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
            p = tf.name
        try:
            jetxl.write_sheets_arrow([sheet], p, 1)
            results.append(("multi_file", open(p, "rb").read()))
        finally:
            os.unlink(p)
    return results


def main():
    print("=" * 74)
    print(f"jetxl property fuzzing — {N} iterations, seed={SEED}")
    have_schema = bool(load_schemas())
    print(f"schema validation (P3): {'ENABLED' if have_schema else 'unavailable (skipped)'}")
    print("=" * 74)

    master = random.Random(SEED)
    p1 = p2 = p3 = 0
    failures = []
    checked_paths = 0

    for i in range(N):
        seed_i = master.randint(0, 2**31)
        rng = random.Random(seed_i)
        try:
            data, ncols, nrows = rng_table(rng)
            cfg = rng_config(rng, ncols, nrows)
        except Exception as e:
            # generator bug, not a jetxl bug
            print(f"  [gen error iter {i} seed {seed_i}] {e!r}")
            continue

        # P1: no panic/crash
        try:
            outs = run_paths(data, cfg, rng)
        except Exception as e:
            failures.append(("P1_CRASH", i, seed_i, repr(e)[:120], cfg))
            continue
        p1 += 1

        for label, b in outs:
            checked_paths += 1
            # P2: loadable (warnings-as-errors)
            try:
                with warnings.catch_warnings():
                    warnings.simplefilter("error", UserWarning)
                    openpyxl.load_workbook(io.BytesIO(b))
                p2 += 1
            except Exception as e:
                failures.append(("P2_LOAD", i, seed_i, f"[{label}] {e!r}"[:120], cfg))
                continue
            # P3: schema conformance
            if have_schema:
                v = schema_violations(b)
                if v:
                    failures.append(("P3_SCHEMA", i, seed_i, f"[{label}] {v[0]}"[:120], cfg))
                else:
                    p3 += 1

        if VERBOSE and i % 50 == 0:
            print(f"  ... {i}/{N} (crashes={sum(1 for f in failures if f[0]=='P1_CRASH')}, "
                  f"load={sum(1 for f in failures if f[0]=='P2_LOAD')}, "
                  f"schema={sum(1 for f in failures if f[0]=='P3_SCHEMA')})")

    print("\n" + "=" * 74)
    print(f"P1 no-crash:   {p1}/{N} configs")
    print(f"P2 loadable:   {p2}/{checked_paths} path-writes")
    if have_schema:
        print(f"P3 schema-ok:  {p3}/{checked_paths} path-writes")
    print(f"total failures: {len(failures)}")
    if failures:
        print("\n--- failures (first 15) ---")
        for kind, i, seed_i, msg, cfg in failures[:15]:
            print(f"\n[{kind}] iter={i} seed={seed_i}")
            print(f"  {msg}")
            print(f"  config keys: {sorted(cfg.keys())}")
        # write a repro file for the first failure
        if failures:
            with open("/tmp/fuzz_repro.txt", "w") as f:
                for kind, i, seed_i, msg, cfg in failures:
                    f.write(f"{kind}\titer={i}\tseed={seed_i}\t{msg}\n\t{cfg}\n\n")
            print("\n(full repro list written to /tmp/fuzz_repro.txt)")
    print("=" * 74)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
