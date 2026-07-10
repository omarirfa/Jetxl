#!/usr/bin/env python3
"""
jetxl benchmark harness
=======================

Measures jetxl write performance across the axes you care about:

  1. Arrow API vs Dict (legacy) API          -> write_sheet_arrow vs write_sheet
  2. Single-sheet vs Multi-sheet             -> write_sheet* vs write_sheets*
  3. File vs Bytes (disk I/O vs in-memory)   -> write_*_arrow vs write_*_arrow_to_bytes
  4. Data type (dtype) comparison            -> int / float / str / bool / date / mixed
  5. Thread scaling for multi-sheet          -> num_threads = 1,2,4,8,...

This version exercises ALL SIX public functions, not just the single-sheet
bytes one:

    write_sheet                    (dict single, file)
    write_sheets                   (dict multi, file)
    write_sheet_arrow              (arrow single, file)
    write_sheets_arrow             (arrow multi, file)
    write_sheet_arrow_to_bytes     (arrow single, bytes)
    write_sheets_arrow_to_bytes    (arrow multi, bytes)   <-- previously untested

The file-vs-bytes axis is applied SYMMETRICALLY: both single- and multi-sheet
arrow paths are timed to file and to bytes, so you can see the true cost of disk
I/O at every scale and sheet count.

Design goals:
  - Statistically honest: warmup run + N timed reps, report MEDIAN (robust to
    outliers) plus min and stdev. Throughput in rows/s and MB/s.
  - Fair: the SAME logical data is fed to every API. Conversion cost
    (DataFrame -> Arrow, or building the dict) is measured SEPARATELY from the
    write, and also reported combined, so you can see both "pure write" and
    "end-to-end" numbers.
  - Robust: gracefully skips anything not installed (polars/pyarrow/pandas), and
    skips APIs jetxl doesn't expose in your build.
  - Reproducible: fixed RNG seed, prints environment + library versions.

Usage:
    python jetxl_benchmark.py                 # sensible defaults
    python jetxl_benchmark.py --rows 10000 100000 1000000
    python jetxl_benchmark.py --reps 7 --sheets 8 --threads 1 2 4 8
    python jetxl_benchmark.py --arrow-backend polars   # force a backend
    python jetxl_benchmark.py --csv results.csv        # also dump raw CSV
    python jetxl_benchmark.py --quick                  # fast smoke run
"""

from __future__ import annotations
import argparse
import gc
import os
import platform
import statistics
import sys
import tempfile
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta

# ----------------------------------------------------------------------------
# Optional imports — we detect what's available and adapt.
# ----------------------------------------------------------------------------
def _try(name):
    try:
        return __import__(name)
    except Exception:
        return None

np = _try("numpy")
pl = _try("polars")
pa = _try("pyarrow")
pd = _try("pandas")

try:
    import jetxl
except Exception as e:  # pragma: no cover
    print("ERROR: could not import jetxl:", e)
    print("This benchmark must run in an environment where jetxl is installed.")
    print("  pip install jetxl   (and ideally polars or pyarrow for the Arrow path)")
    sys.exit(1)

if np is None:
    print("ERROR: numpy is required to synthesize benchmark data (pip install numpy).")
    sys.exit(1)


# ----------------------------------------------------------------------------
# Data generation. We build column dicts of native Python-friendly arrays, then
# provide adapters to (a) the dict API and (b) an Arrow table via whichever
# backend is available. The SAME underlying values are used everywhere.
# ----------------------------------------------------------------------------
DTYPES = ["int", "float", "str", "bool", "date", "mixed"]

def make_columns(n_rows: int, dtype: str, seed: int = 42) -> dict:
    """Return {col_name: numpy array or python list} for the requested dtype.

    'mixed' mirrors a realistic business sheet: a couple of id/label string
    columns, some floats, an int, a bool, and a date — this is the case most
    users actually hit, so it's the most representative single number.
    """
    rng = np.random.default_rng(seed)

    if dtype == "int":
        return {f"int_{i}": rng.integers(-1_000_000, 1_000_000, n_rows, dtype=np.int64)
                for i in range(5)}

    if dtype == "float":
        return {f"float_{i}": rng.standard_normal(n_rows) * 1e6
                for i in range(5)}

    if dtype == "bool":
        return {f"bool_{i}": rng.integers(0, 2, n_rows, dtype=bool)
                for i in range(5)}

    if dtype == "str":
        # Realistic-ish short strings with repetition (categorical-like), which
        # is the common real case and also stresses the inline-string path.
        pool = np.array([f"Item-{k:04d}" for k in range(2000)])
        return {f"str_{i}": pool[rng.integers(0, len(pool), n_rows)]
                for i in range(5)}

    if dtype == "date":
        base = np.datetime64("2020-01-01")
        return {f"date_{i}": base + rng.integers(0, 2000, n_rows).astype("timedelta64[D]")
                for i in range(5)}

    if dtype == "mixed":
        pool = np.array([f"Region-{k}" for k in range(50)])
        names = np.array([f"Person {k}" for k in range(500)])
        base = np.datetime64("2020-01-01")
        return {
            "Region": pool[rng.integers(0, len(pool), n_rows)],
            "SalesPerson": names[rng.integers(0, len(names), n_rows)],
            "Quota": rng.integers(0, 300_000, n_rows, dtype=np.int64),
            "Sales": rng.standard_normal(n_rows) * 50_000 + 100_000,
            "Active": rng.integers(0, 2, n_rows, dtype=bool),
            "AsOf": base + rng.integers(0, 2000, n_rows).astype("timedelta64[D]"),
        }

    raise ValueError(f"unknown dtype {dtype}")


# --- dict-API adapter: convert numpy arrays -> plain python lists ------------
def columns_to_dict(cols: dict) -> dict:
    """The legacy dict API wants {name: list-of-python-scalars}. Converting numpy
    -> python objects is part of that API's real cost, so we time it in the
    'convert' phase, not the write phase."""
    out = {}
    for k, v in cols.items():
        if hasattr(v, "dtype") and str(v.dtype).startswith("datetime64"):
            # dict API accepts python datetime for Date cells
            out[k] = [datetime(1970, 1, 1) + timedelta(days=int((x - np.datetime64("1970-01-01")) / np.timedelta64(1, "D")))
                      for x in v]
        else:
            out[k] = v.tolist()
    return out


# --- Arrow-API adapter: build an Arrow table via best available backend ------
def _pick_arrow_backend(forced: str | None):
    order = [forced] if forced else ["polars", "pyarrow", "pandas"]
    for b in order:
        if b == "polars" and pl is not None:
            return "polars"
        if b == "pyarrow" and pa is not None:
            return "pyarrow"
        if b == "pandas" and pd is not None and pa is not None:
            return "pandas"
    return None

def columns_to_arrow(cols: dict, backend: str):
    """Build an Arrow table (RecordBatch/Table) the way a real user would.
    Returns the arrow object jetxl.write_sheet_arrow accepts."""
    if backend == "polars":
        return pl.DataFrame({k: v for k, v in cols.items()}).to_arrow()
    if backend == "pyarrow":
        arrays = {}
        for k, v in cols.items():
            arrays[k] = pa.array(v)
        return pa.table(arrays)
    if backend == "pandas":
        return pd.DataFrame({k: v for k, v in cols.items()}).to_arrow() \
            if hasattr(pd.DataFrame, "to_arrow") else pa.Table.from_pandas(pd.DataFrame(cols))
    raise ValueError(backend)


# ----------------------------------------------------------------------------
# Timing core
# ----------------------------------------------------------------------------
@dataclass
class Timing:
    label: str
    rows: int
    reps: list = field(default_factory=list)          # per-rep write seconds
    convert_s: float = 0.0                             # one-time conversion cost
    out_bytes: int = 0
    ok: bool = True
    note: str = ""

    @property
    def median(self): return statistics.median(self.reps) if self.reps else float("nan")
    @property
    def best(self):   return min(self.reps) if self.reps else float("nan")
    @property
    def stdev(self):  return statistics.pstdev(self.reps) if len(self.reps) > 1 else 0.0
    @property
    def rows_per_s(self): return self.rows / self.median if self.median else 0.0
    @property
    def mb_per_s(self):
        return (self.out_bytes / 1e6) / self.median if self.median and self.out_bytes else 0.0


def time_it(fn, reps: int, warmup: int = 1) -> list:
    """Run fn() warmup times (discarded), then reps times, returning seconds each.
    gc is disabled during a rep for lower variance, then re-enabled + collected."""
    for _ in range(warmup):
        fn()
    out = []
    for _ in range(reps):
        gc.collect(); gc.disable()
        t0 = time.perf_counter()
        fn()
        dt = time.perf_counter() - t0
        gc.enable()
        out.append(dt)
    return out


# ----------------------------------------------------------------------------
# The actual write closures for each API variant
# ----------------------------------------------------------------------------
def bench_single(cols: dict, rows: int, dtype: str, reps: int, backend: str | None,
                 tmpdir: str) -> list[Timing]:
    """Single-sheet writes: arrow->file, arrow->bytes, dict->file."""
    results = []

    # ---- Arrow single (build the table ONCE, share across file & bytes) ----
    b = _pick_arrow_backend(backend)
    if b is not None and hasattr(jetxl, "write_sheet_arrow"):
        t0 = time.perf_counter()
        arrow = columns_to_arrow(cols, b)
        conv = time.perf_counter() - t0

        # arrow single -> FILE
        path = os.path.join(tmpdir, "arrow_single.xlsx")
        def run(): jetxl.write_sheet_arrow(arrow, path)
        try:
            reps_s = time_it(run, reps)
            size = os.path.getsize(path)
            results.append(Timing(f"arrow-single-file[{b}]", rows, reps_s, conv, size))
        except Exception as e:
            results.append(Timing(f"arrow-single-file[{b}]", rows, ok=False, note=str(e)))

        # arrow single -> BYTES (no disk I/O)
        if hasattr(jetxl, "write_sheet_arrow_to_bytes"):
            def run_b(): jetxl.write_sheet_arrow_to_bytes(arrow)
            try:
                reps_s = time_it(run_b, reps)
                size = len(jetxl.write_sheet_arrow_to_bytes(arrow))
                results.append(Timing(f"arrow-single-bytes[{b}]", rows, reps_s, conv, size))
            except Exception as e:
                results.append(Timing(f"arrow-single-bytes[{b}]", rows, ok=False, note=str(e)))
    else:
        results.append(Timing("arrow-single-file", rows, ok=False,
                              note="no arrow backend (install polars or pyarrow)"))

    # ---- Dict single -> FILE ----
    if hasattr(jetxl, "write_sheet"):
        t0 = time.perf_counter()
        d = columns_to_dict(cols)
        conv = time.perf_counter() - t0
        path = os.path.join(tmpdir, "dict_single.xlsx")
        def run_d(): jetxl.write_sheet(d, path)
        try:
            reps_s = time_it(run_d, reps)
            size = os.path.getsize(path)
            results.append(Timing("dict-single-file", rows, reps_s, conv, size))
        except Exception as e:
            results.append(Timing("dict-single-file", rows, ok=False, note=str(e)))
    else:
        results.append(Timing("dict-single-file", rows, ok=False, note="jetxl.write_sheet missing"))

    return results


def bench_multi(cols: dict, rows_total: int, n_sheets: int, dtype: str, reps: int,
                backend: str | None, threads_list: list[int], tmpdir: str) -> list[Timing]:
    """Multi-sheet writes across each thread count, for:
        arrow -> file   (write_sheets_arrow)
        arrow -> bytes  (write_sheets_arrow_to_bytes)   <-- now covered
        dict  -> file   (write_sheets)

    rows_total is kept constant so single vs multi is comparable on total work
    (same number of cells overall)."""
    results = []
    per = max(1, rows_total // n_sheets)

    # Build the per-sheet column slices ONCE (shared logical data)
    sheet_cols = []
    for s in range(n_sheets):
        sl = {k: v[:per] for k, v in cols.items()}
        sheet_cols.append(sl)

    b = _pick_arrow_backend(backend)

    # ---- Arrow multi: build the sheet list ONCE, reuse for file & bytes ----
    if b is not None and (hasattr(jetxl, "write_sheets_arrow")
                          or hasattr(jetxl, "write_sheets_arrow_to_bytes")):
        t0 = time.perf_counter()
        arrow_sheets = [{"data": columns_to_arrow(sc, b), "name": f"S{i+1}"}
                        for i, sc in enumerate(sheet_cols)]
        conv = time.perf_counter() - t0

        # arrow multi -> FILE, per thread count
        if hasattr(jetxl, "write_sheets_arrow"):
            for th in threads_list:
                path = os.path.join(tmpdir, f"arrow_multi_{th}.xlsx")
                def run(th=th, path=path): jetxl.write_sheets_arrow(arrow_sheets, path, th)
                try:
                    reps_s = time_it(run, reps)
                    size = os.path.getsize(path)
                    results.append(Timing(f"arrow-multi-file[{b}]-t{th}", per * n_sheets, reps_s, conv, size))
                except Exception as e:
                    results.append(Timing(f"arrow-multi-file[{b}]-t{th}", per * n_sheets, ok=False, note=str(e)))

        # arrow multi -> BYTES, per thread count (previously never benchmarked)
        if hasattr(jetxl, "write_sheets_arrow_to_bytes"):
            for th in threads_list:
                def run_b(th=th): jetxl.write_sheets_arrow_to_bytes(arrow_sheets, th)
                try:
                    reps_s = time_it(run_b, reps)
                    size = len(jetxl.write_sheets_arrow_to_bytes(arrow_sheets, th))
                    results.append(Timing(f"arrow-multi-bytes[{b}]-t{th}", per * n_sheets, reps_s, conv, size))
                except Exception as e:
                    results.append(Timing(f"arrow-multi-bytes[{b}]-t{th}", per * n_sheets, ok=False, note=str(e)))
    else:
        results.append(Timing("arrow-multi-file", rows_total, ok=False,
                              note="no arrow backend"))

    # ---- Dict multi -> FILE, per thread count ----
    if hasattr(jetxl, "write_sheets"):
        t0 = time.perf_counter()
        dict_sheets = [{"name": f"S{i+1}", "columns": columns_to_dict(sc)}
                       for i, sc in enumerate(sheet_cols)]
        conv = time.perf_counter() - t0
        for th in threads_list:
            path = os.path.join(tmpdir, f"dict_multi_{th}.xlsx")
            def run_d(th=th, path=path): jetxl.write_sheets(dict_sheets, path, th)
            try:
                reps_s = time_it(run_d, reps)
                size = os.path.getsize(path)
                results.append(Timing(f"dict-multi-file-t{th}", per * n_sheets, reps_s, conv, size))
            except Exception as e:
                results.append(Timing(f"dict-multi-file-t{th}", per * n_sheets, ok=False, note=str(e)))
    else:
        results.append(Timing("dict-multi-file", rows_total, ok=False, note="jetxl.write_sheets missing"))

    return results


# ----------------------------------------------------------------------------
# Reporting  (Rich tables, with a plain-text fallback if rich isn't installed)
# ----------------------------------------------------------------------------
try:
    from rich.console import Console
    from rich.table import Table as RichTable
    from rich.text import Text
    from rich import box
    _RICH = True
    _console = Console()
except Exception:
    _RICH = False
    _console = None


def fmt_rows(n):
    """Human row count: 10000 -> '10K', 100000 -> '100K', 1000000 -> '1M'."""
    if n >= 1_000_000:
        v = n / 1_000_000
        return (f"{v:.0f}M" if v == int(v) else f"{v:.1f}M")
    if n >= 1_000:
        v = n / 1_000
        return (f"{v:.0f}K" if v == int(v) else f"{v:.1f}K")
    return str(n)


def fmt_time(x):
    """Full time formatter: always seconds, plus ms in parens under a second,
    and m:ss for >= 60s.
      0.0736 -> '0.074 s (73.6 ms)'
      2.145  -> '2.145 s'
      95.3   -> '1m 35.3s (95.300 s)'
    """
    if x != x:  # nan
        return "—"
    if x < 1:
        return f"{x:.3f} s ({x*1000:.1f} ms)"
    if x < 60:
        return f"{x:.3f} s"
    m = int(x // 60)
    s = x - m * 60
    return f"{m}m {s:04.1f}s ({x:.3f} s)"


def fmt_sec(x):
    """Compact seconds, with m:ss for >= 60s. Used in the wide comparison tables."""
    if x != x:
        return "—"
    if x < 60:
        return f"{x:.3f} s"
    m = int(x // 60)
    s = x - m * 60
    return f"{m}m{s:04.1f}s"


# Existing report code calls fmt_ms; alias it to the seconds formatter so every
# timing now reads in seconds (with m:ss over a minute) instead of milliseconds.
fmt_ms = fmt_sec


def fmt_rate(x):
    """Throughput as rows/s in K/M units."""
    if x != x or x == 0:
        return "—"
    if x >= 1e6:
        return f"{x/1e6:.2f}M/s"
    if x >= 1e3:
        return f"{x/1e3:.0f}K/s"
    return f"{x:.0f}/s"


def get(rowmap, prefix, th=None):
    """First ok Timing whose label starts with prefix (and matches thread suffix)."""
    for lbl, r in rowmap.items():
        if not r.ok:
            continue
        if th is not None:
            if lbl.startswith(prefix) and lbl.endswith(f"-t{th}"):
                return r
        elif lbl.startswith(prefix):
            return r
    return None


def best_multi(mm, prefix):
    """Fastest ok multi-sheet Timing across all thread counts for a given prefix."""
    best = None
    for lbl, r in mm.items():
        if r.ok and lbl.startswith(prefix) and (best is None or r.median < best.median):
            best = r
    return best


# ---- Rich rendering core --------------------------------------------------
def _render(title, subtitle, row_label_header, col_headers, data_rows,
            highlight="max", caption=None):
    """data_rows: list of (row_label, [cell_str,...], [raw_or_None,...]).
    highlight: 'max' -> best (fastest) cell is the largest raw; 'min' -> smallest;
               None -> no highlight. Cells with raw=None are never highlighted."""
    if _RICH:
        table = RichTable(
            title=f"[bold]{title}[/bold]" + (f"\n[dim]{subtitle}[/dim]" if subtitle else ""),
            box=box.ROUNDED, header_style="bold cyan", title_justify="left",
            caption=(f"[dim]{caption}[/dim]" if caption else None), caption_justify="left",
            expand=False, pad_edge=False,
        )
        table.add_column(row_label_header, justify="right", style="bold white", no_wrap=True)
        for h in col_headers:
            table.add_column(h, justify="right", no_wrap=True)
        for label, cells, raws in data_rows:
            best = None
            if highlight:
                cand = [(i, v) for i, v in enumerate(raws)
                        if isinstance(v, (int, float)) and v == v]
                if cand:
                    best = (max if highlight == "max" else min)(cand, key=lambda t: t[1])[0]
            styled = []
            for i, c in enumerate(cells):
                if i == best:
                    styled.append(Text(c, style="bold green"))
                elif c == "—":
                    styled.append(Text(c, style="dim"))
                else:
                    styled.append(Text(c))
            table.add_row(label, *styled)
        _console.print(table)
        _console.print()
    else:
        # plain-text fallback
        print(f"\n{title}")
        if subtitle:
            print(f"  {subtitle}")
        allcols = [row_label_header] + col_headers
        widths = [len(h) for h in allcols]
        for label, cells, _ in data_rows:
            widths[0] = max(widths[0], len(label))
            for i, c in enumerate(cells):
                widths[i+1] = max(widths[i+1], len(c))
        def fmt_line(vals):
            return "  ".join(v.rjust(widths[i]) for i, v in enumerate(vals))
        print(fmt_line(allcols))
        print("-" * (sum(widths) + 2 * len(widths)))
        for label, cells, raws in data_rows:
            best = None
            if highlight:
                cand = [(i, v) for i, v in enumerate(raws)
                        if isinstance(v, (int, float)) and v == v]
                if cand:
                    best = (max if highlight == "max" else min)(cand, key=lambda t: t[1])[0]
            marked = [(("*" + c) if i == best else c) for i, c in enumerate(cells)]
            row_vals = [label] + [m.rjust(widths[i+1]) for i, m in enumerate(marked)]
            print("  ".join(v.rjust(widths[i]) if i == 0 else v
                            for i, v in enumerate(row_vals)))
        if caption:
            print(f"  {caption}")


def banner(msg):
    if _RICH:
        _console.rule(f"[bold]{msg}[/bold]")
    else:
        print("\n" + "=" * 72 + f"\n{msg}\n" + "=" * 72)


# ---- The comparison reports -----------------------------------------------
def report_dtype(dtype_results, dtype_rows):
    rows = []
    for dt, rowmap in dtype_results.items():
        a = get(rowmap, "arrow-single-file")
        d = get(rowmap, "dict-single-file")
        ar = a.rows_per_s if a else float("nan")
        dr = d.rows_per_s if d else float("nan")
        sp = (d.median / a.median) if (a and d and a.median) else float("nan")
        rows.append((dt,
                     [fmt_rate(ar), fmt_rate(dr),
                      fmt_ms(a.median if a else float("nan")),
                      fmt_ms(d.median if d else float("nan")),
                      (f"{sp:.2f}×" if sp == sp else "—")],
                     [ar, dr, None, None, None]))
    _render("Dtype comparison",
            f"single sheet · {fmt_rows(dtype_rows)} rows · throughput",
            "dtype",
            ["arrow rows/s", "dict rows/s", "arrow time", "dict time", "arrow ÷ dict"],
            rows, highlight="max",
            caption="green = fastest in row · 'arrow ÷ dict' = how many × faster arrow is")


def report_scaling(single_by_n, multi_by_n, sheets):
    """Now includes multi-BYTES alongside multi-FILE."""
    cols = ["arrow single file", "arrow single bytes", "dict single file",
            "arrow multi file", "arrow multi bytes", "dict multi file"]

    # time view (fastest = smallest time -> highlight min)
    trows = []
    for n in sorted(single_by_n):
        sm = single_by_n[n]; mm = multi_by_n.get(n, {})
        a_sf = get(sm, "arrow-single-file"); a_sb = get(sm, "arrow-single-bytes")
        d_s = get(sm, "dict-single-file")
        a_mf = best_multi(mm, "arrow-multi-file"); a_mb = best_multi(mm, "arrow-multi-bytes")
        d_m = best_multi(mm, "dict-multi-file")
        med = lambda x: x.median if x else float("nan")
        trows.append((fmt_rows(n),
                      [fmt_ms(med(a_sf)), fmt_ms(med(a_sb)), fmt_ms(med(d_s)),
                       fmt_ms(med(a_mf)), fmt_ms(med(a_mb)), fmt_ms(med(d_m))],
                      [med(a_sf), med(a_sb), med(d_s), med(a_mf), med(a_mb), med(d_m)]))
    _render("Scaling — write time", "dtype=mixed · lower is better",
            "rows", cols, trows, highlight="min",
            caption=f"multi = {sheets} sheets, best thread count · green = fastest in row")

    # throughput view (fastest = largest rate -> highlight max)
    rrows = []
    for n in sorted(single_by_n):
        sm = single_by_n[n]; mm = multi_by_n.get(n, {})
        a_sf = get(sm, "arrow-single-file"); a_sb = get(sm, "arrow-single-bytes")
        d_s = get(sm, "dict-single-file")
        a_mf = best_multi(mm, "arrow-multi-file"); a_mb = best_multi(mm, "arrow-multi-bytes")
        d_m = best_multi(mm, "dict-multi-file")
        rate = lambda x: x.rows_per_s if x else float("nan")
        rrows.append((fmt_rows(n),
                      [fmt_rate(rate(a_sf)), fmt_rate(rate(a_sb)), fmt_rate(rate(d_s)),
                       fmt_rate(rate(a_mf)), fmt_rate(rate(a_mb)), fmt_rate(rate(d_m))],
                      [rate(a_sf), rate(a_sb), rate(d_s), rate(a_mf), rate(a_mb), rate(d_m)]))
    _render("Scaling — throughput", "dtype=mixed · higher is better",
            "rows", cols, rrows, highlight="max",
            caption="green = fastest in row")


def report_file_vs_bytes(single_by_n, multi_by_n, sheets):
    """Dedicated file-vs-bytes view for BOTH single and multi arrow paths."""
    rows = []
    for n in sorted(single_by_n):
        sm = single_by_n[n]; mm = multi_by_n.get(n, {})
        a_sf = get(sm, "arrow-single-file"); a_sb = get(sm, "arrow-single-bytes")
        a_mf = best_multi(mm, "arrow-multi-file"); a_mb = best_multi(mm, "arrow-multi-bytes")
        med = lambda x: x.median if x else float("nan")
        s_speedup = (med(a_sf) / med(a_sb)) if (a_sf and a_sb and med(a_sb)) else float("nan")
        m_speedup = (med(a_mf) / med(a_mb)) if (a_mf and a_mb and med(a_mb)) else float("nan")
        rows.append((fmt_rows(n),
                     [fmt_ms(med(a_sf)), fmt_ms(med(a_sb)),
                      (f"{s_speedup:.2f}×" if s_speedup == s_speedup else "—"),
                      fmt_ms(med(a_mf)), fmt_ms(med(a_mb)),
                      (f"{m_speedup:.2f}×" if m_speedup == m_speedup else "—")],
                     [None]*6))
    _render("File vs Bytes — arrow paths", "dtype=mixed · lower time is better",
            "rows",
            ["single file", "single bytes", "single f÷b",
             "multi file", "multi bytes", "multi f÷b"],
            rows, highlight=None,
            caption="f÷b = file_time ÷ bytes_time (>1 means bytes is faster, i.e. "
                    "disk I/O overhead) · multi uses best thread count")


def report_arrow_vs_dict(single_by_n):
    rows = []
    for n in sorted(single_by_n):
        sm = single_by_n[n]
        a = get(sm, "arrow-single-file"); d = get(sm, "dict-single-file")
        if not (a and d):
            rows.append((fmt_rows(n), ["—", "—", "—", "—", "—", "—"], [None]*6))
            continue
        pure = d.median / a.median if a.median else float("nan")
        ae = a.median + a.convert_s; de = d.median + d.convert_s
        e2e = de / ae if ae else float("nan")
        rows.append((fmt_rows(n),
                     [fmt_ms(a.median), fmt_ms(d.median), f"{pure:.2f}×",
                      fmt_ms(ae), fmt_ms(de), f"{e2e:.2f}×"],
                     [None]*6))
    _render("Arrow vs Dict — speedup", "single sheet · dtype=mixed",
            "rows",
            ["arrow write", "dict write", "pure ×", "arrow e2e", "dict e2e", "e2e ×"],
            rows, highlight=None,
            caption="pure = write only · e2e = write + building the Arrow table / dict")


def report_threads(multi_by_n, threads, sheets):
    """Thread scaling now covers arrow-file, arrow-bytes AND dict-file."""
    for api, prefix in [("arrow file", "arrow-multi-file"),
                        ("arrow bytes", "arrow-multi-bytes"),
                        ("dict file", "dict-multi-file")]:
        have = any(get(mm, prefix, th) for mm in multi_by_n.values() for th in threads)
        if not have:
            continue
        rows = []
        for n in sorted(multi_by_n):
            mm = multi_by_n[n]
            base = get(mm, prefix, threads[0])
            cells = []; raws = []
            top_eff = float("nan")
            for th in threads:
                r = get(mm, prefix, th)
                if r and base and r.median:
                    sp = base.median / r.median
                    cells.append(f"{sp:.2f}×"); raws.append(sp)
                    if th > 1:
                        top_eff = sp / th * 100
                else:
                    cells.append("—"); raws.append(None)
            cells.append(f"{top_eff:.0f}%" if top_eff == top_eff else "—")
            raws.append(None)
            rows.append((fmt_rows(n), cells, raws))
        _render(f"Thread scaling — {api} multi-sheet",
                f"{sheets} sheets · speedup vs 1 thread",
                "rows",
                [f"{th} thr" for th in threads] + [f"eff@{max(threads)}t"],
                rows, highlight="max",
                caption=f"green = best speedup in row · eff@{max(threads)}t = "
                        f"speedup ÷ threads at the highest thread count")


def report_detailed_timings(single_by_n, multi_by_n, threads, sheets):
    """Granular per-variant timing breakdown: median / best / worst / jitter
    (stdev) / throughput, one table per variant across all row counts. Now also
    covers the multi-BYTES variants."""

    def worst(r):
        return max(r.reps) if r and r.reps else float("nan")

    # single-sheet variants
    variants = [
        ("arrow single sheet (file)",  single_by_n, "arrow-single-file",  None),
        ("arrow single sheet (bytes)", single_by_n, "arrow-single-bytes", None),
        ("dict single sheet (file)",   single_by_n, "dict-single-file",   None),
    ]
    # multi at each thread count so you can see the timing per thread setting
    for th in threads:
        variants.append((f"arrow multi FILE ({sheets} sheets, {th} thr)", multi_by_n,
                         "arrow-multi-file", th))
    for th in threads:
        variants.append((f"arrow multi BYTES ({sheets} sheets, {th} thr)", multi_by_n,
                         "arrow-multi-bytes", th))
    for th in threads:
        variants.append((f"dict multi FILE ({sheets} sheets, {th} thr)", multi_by_n,
                         "dict-multi-file", th))

    printed_header = False
    for vlabel, source, prefix, th in variants:
        rows = []
        any_ok = False
        for n in sorted(source):
            r = get(source[n], prefix, th)
            if not r:
                rows.append((fmt_rows(n), ["—", "—", "—", "—", "—", "—"], [None]*6))
                continue
            any_ok = True
            rows.append((
                fmt_rows(n),
                [fmt_time(r.median), fmt_time(r.best), fmt_time(worst(r)),
                 (f"±{r.stdev*1000:.1f} ms" if r.stdev < 1 else f"±{r.stdev:.3f} s"),
                 fmt_rate(r.rows_per_s),
                 fmt_time(r.convert_s)],
                [None]*6,
            ))
        if not any_ok:
            continue
        if not printed_header:
            banner("Detailed timings (per variant, all row counts)")
            printed_header = True
        _render(vlabel, "median = headline · jitter = run-to-run stdev · all times in seconds",
                "rows",
                ["median", "best", "worst", "jitter", "rows/s", "convert"],
                rows, highlight=None,
                caption="convert = one-time cost to build the Arrow table / python-list dict")


def dump_csv(path, all_rows):
    import csv
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["scenario", "dtype", "variant", "rows", "median_s", "best_s",
                    "stdev_s", "convert_s", "rows_per_s", "mb_per_s", "out_bytes", "ok", "note"])
        for scen, dtype, r in all_rows:
            w.writerow([scen, dtype, r.label, r.rows, f"{r.median:.6f}", f"{r.best:.6f}",
                        f"{r.stdev:.6f}", f"{r.convert_s:.6f}", f"{r.rows_per_s:.1f}",
                        f"{r.mb_per_s:.3f}", r.out_bytes, r.ok, r.note])
    msg = f"Raw per-variant results written to {path}"
    _console.print(f"[green]{msg}[/green]") if _RICH else print("\n" + msg)


def print_skips(all_rows):
    skipped = {(r.label.split("[")[0].split("-t")[0], r.note)
               for _, _, r in all_rows if not r.ok and r.note}
    if not skipped:
        return
    if _RICH:
        _console.print("\n[yellow]Skipped variants (not available here):[/yellow]")
        for lbl, note in sorted(skipped):
            _console.print(f"  [dim]•[/dim] [bold]{lbl}[/bold]: [dim]{note}[/dim]")
    else:
        print("\nSkipped variants (not available here):")
        for lbl, note in sorted(skipped):
            print(f"  - {lbl}: {note}")


# ----------------------------------------------------------------------------
# Main — runs ALL axes by default and prints side-by-side Rich tables
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="jetxl benchmark harness (side-by-side tables)")
    ap.add_argument("--rows", type=int, nargs="+",
                    default=[10_000, 20_000, 50_000, 100_000, 500_000, 1_000_000],
                    help="row counts to test (default 10k/20k/50k/100k/500k/1M)")
    ap.add_argument("--reps", type=int, default=10, help="timed reps per case (median reported)")
    ap.add_argument("--sheets", type=int, default=8, help="sheet count for multi-sheet test")
    ap.add_argument("--threads", type=int, nargs="+", default=[1, 2, 4, 8],
                    help="thread counts for multi-sheet test")
    ap.add_argument("--dtypes", nargs="+", default=DTYPES, choices=DTYPES,
                    help="which dtypes to test")
    ap.add_argument("--arrow-backend", choices=["polars", "pyarrow", "pandas"], default=None,
                    help="force a specific Arrow backend (default: auto-detect)")
    ap.add_argument("--csv", default=None, help="also write raw per-variant results to this CSV")
    ap.add_argument("--quick", action="store_true",
                    help="fast smoke run: 10k/50k/100k rows, 3 reps, 4 sheets, threads 1,4")
    args = ap.parse_args()

    if args.quick:
        args.rows = [10_000, 50_000, 100_000]
        args.reps = 3
        args.sheets = 4
        args.threads = [1, 4]

    backend = _pick_arrow_backend(args.arrow_backend)

    # ---- environment banner ----
    if _RICH:
        from rich.panel import Panel
        env = (
            f"[bold]python[/bold] {platform.python_version()} ({platform.machine()})   "
            f"[bold]cpus[/bold] {os.cpu_count()}   "
            f"[bold]jetxl[/bold] {getattr(jetxl,'__version__','unknown')}\n"
            f"[bold]numpy[/bold] {getattr(np,'__version__','—')}   "
            f"[bold]polars[/bold] {getattr(pl,'__version__','—') if pl else '—'}   "
            f"[bold]pyarrow[/bold] {getattr(pa,'__version__','—') if pa else '—'}   "
            f"[bold]pandas[/bold] {getattr(pd,'__version__','—') if pd else '—'}\n"
            f"[bold]arrow backend[/bold] {backend or '[yellow]NONE (arrow variants skipped)[/yellow]'}\n"
            f"[bold]config[/bold] rows={[fmt_rows(n) for n in args.rows]} reps={args.reps} "
            f"sheets={args.sheets} threads={args.threads}\n"
            f"[bold]dtypes[/bold] {', '.join(args.dtypes)}"
        )
        _console.print(Panel(env, title="[bold]jetxl benchmark[/bold]",
                             subtitle="all 6 write functions · side-by-side", expand=False))
    else:
        print("=" * 72)
        print("jetxl benchmark  —  side-by-side comparison (all 6 functions)")
        print("=" * 72)
        print(f"python {platform.python_version()} | cpus={os.cpu_count()} | "
              f"jetxl={getattr(jetxl,'__version__','unknown')}")
        print(f"arrow backend = {backend or 'NONE (arrow variants skipped)'}")
        print(f"config: rows={[fmt_rows(n) for n in args.rows]} reps={args.reps} "
              f"sheets={args.sheets} threads={args.threads} dtypes={args.dtypes}")

    all_rows = []
    dtype_results = {}
    single_by_n = {}
    multi_by_n = {}

    total_steps = len(args.dtypes) + 2 * len(args.rows)

    def run_all(progress=None, task=None):
        with tempfile.TemporaryDirectory() as tmp:
            dtype_rows = args.rows[len(args.rows)//2]
            for dt in args.dtypes:
                cols = make_columns(dtype_rows, dt)
                res = bench_single(cols, dtype_rows, dt, args.reps, args.arrow_backend, tmp)
                dtype_results[dt] = {r.label: r for r in res}
                all_rows.extend(("dtype", dt, r) for r in res)
                if progress: progress.advance(task)
            for n in args.rows:
                cols = make_columns(n, "mixed")
                single = bench_single(cols, n, "mixed", args.reps, args.arrow_backend, tmp)
                single_by_n[n] = {r.label: r for r in single}
                all_rows.extend(("single", "mixed", r) for r in single)
                if progress: progress.advance(task)
                multi = bench_multi(cols, n, args.sheets, "mixed", args.reps,
                                    args.arrow_backend, args.threads, tmp)
                multi_by_n[n] = {r.label: r for r in multi}
                all_rows.extend(("multi", "mixed", r) for r in multi)
                if progress: progress.advance(task)
        return dtype_rows

    if _RICH:
        from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeElapsedColumn
        with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"),
                      BarColumn(), TextColumn("{task.completed}/{task.total}"),
                      TimeElapsedColumn(), console=_console, transient=True) as prog:
            task = prog.add_task("running benchmarks", total=total_steps)
            dtype_rows = run_all(prog, task)
    else:
        print("\nrunning all scenarios...", flush=True)
        dtype_rows = run_all()

    # ---- side-by-side tables ----
    report_dtype(dtype_results, dtype_rows)
    report_scaling(single_by_n, multi_by_n, args.sheets)
    report_file_vs_bytes(single_by_n, multi_by_n, args.sheets)
    report_arrow_vs_dict(single_by_n)
    report_threads(multi_by_n, args.threads, args.sheets)
    report_detailed_timings(single_by_n, multi_by_n, args.threads, args.sheets)
    print_skips(all_rows)

    if args.csv:
        dump_csv(args.csv, all_rows)

    legend = (
        "median of --reps runs (warmup discarded, GC off per rep) · "
        "rows/s higher = better, time lower = better · "
        "all 6 functions covered: single/multi × file/bytes (arrow) + dict single/multi · "
        "Arrow-vs-Dict shown pure-write AND end-to-end · "
        "multi-sheet holds TOTAL rows constant, thread cols show parallel gain"
    )
    if _RICH:
        _console.print(f"\n[dim]{legend}[/dim]")
    else:
        print("\n" + legend)


if __name__ == "__main__":
    main()
