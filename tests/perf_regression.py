#!/usr/bin/env python3
"""
Performance regression check: fixed jetxl vs pristine baseline.

Builds representative workloads and times both modules over many iterations,
reporting median ns/op and the percentage delta. A positive delta means the
fixed build is SLOWER. We flag anything slower than the noise threshold on the
hot paths (single-sheet write, single-sheet bytes, multi-sheet bytes).

The changes under test:
  - styles.rs: emit real borderId + applyBorder (few extra bytes per styled xf)
  - writer.rs multi-sheet bytes: added a SERIAL style pre-pass + real per-sheet
    style/dxf maps (was an empty map). This is the main thing to watch: it adds
    work proportional to the number of custom styles, but must not slow the
    common no-styles multi-sheet write.
  - lib.rs multi-sheet bytes pyfunction: now parses the full per-sheet config.

Run:
    python perf_regression.py
    python perf_regression.py --iters 40
"""
import argparse
import statistics
import time
import sys

import pyarrow as pa

import jetxl
try:
    import jetxl_baseline
    HAVE_BASELINE = True
except Exception as e:
    print(f"(no baseline module available: {e}) -- timing fixed build only")
    HAVE_BASELINE = False


def big_table(n):
    cats = ["North", "South", "East", "West", "Central"]
    return pa.table({
        "id":     pa.array(list(range(n)), pa.int64()),
        "region": pa.array([cats[i % 5] for i in range(n)]),
        "note":   pa.array([f"unique note {i}" for i in range(n)]),
        "amount": pa.array([i * 1.5 for i in range(n)], pa.float64()),
        "flag":   pa.array([i % 2 == 0 for i in range(n)]),
    })


def bench(fn, iters, warmup=3):
    for _ in range(warmup):
        fn()
    samples = []
    for _ in range(iters):
        t0 = time.perf_counter_ns()
        fn()
        samples.append(time.perf_counter_ns() - t0)
    return statistics.median(samples), min(samples)


def bench_pair(fn_fixed, fn_base, iters):
    """Interleave the two builds so shared drift (disk, CPU freq) cancels out;
    repeat the whole pairing 3 times and take the best (lowest-noise) median
    for each side. Returns (median_base_ms, median_fixed_ms)."""
    best_base = best_fixed = None
    for _ in range(3):
        mb, _ = bench(fn_base, iters)
        mf, _ = bench(fn_fixed, iters)
        best_base = mb if best_base is None else min(best_base, mb)
        best_fixed = mf if best_fixed is None else min(best_fixed, mf)
    return best_base / 1e6, best_fixed / 1e6


def workloads():
    import datetime as _dt
    t = big_table(100_000)
    t_small = big_table(5_000)
    n_dates = 100_000
    base = _dt.date(2000, 1, 1)
    t_dates = pa.table({
        "d": pa.array([base + _dt.timedelta(days=i % 9000) for i in range(n_dates)]),
        "v": pa.array([float(i) for i in range(n_dates)], pa.float64()),
    })
    multi = [{"data": big_table(20_000), "name": f"S{i}"} for i in range(5)]
    multi_styled = [{
        "data": big_table(20_000), "name": f"S{i}",
        "column_formats": {"amount": "currency"},
    } for i in range(5)]

    return {
        "single_file_100k": (
            lambda m: m.write_sheet_arrow(t, "/tmp/perf_fixed.xlsx", auto_filter=True,
                                          column_formats={"amount": "currency"}),
        ),
        "single_bytes_100k": (
            lambda m: m.write_sheet_arrow_to_bytes(t, auto_filter=True,
                                                   column_formats={"amount": "currency"}),
        ),
        "single_bytes_styled_5k": (
            lambda m: m.write_sheet_arrow_to_bytes(
                t_small,
                cell_styles=[{"row": r, "col": 0,
                              "border": {"bottom": {"style": "thin"}},
                              "font": {"bold": True}} for r in range(2, 52)]),
        ),
        "multi_bytes_5x20k": (
            lambda m: m.write_sheets_arrow_to_bytes(multi, 4),
        ),
        "multi_bytes_styled_5x20k": (
            lambda m: m.write_sheets_arrow_to_bytes(multi_styled, 4),
        ),
        "single_bytes_tables_5k": (
            lambda m: m.write_sheet_arrow_to_bytes(
                t_small,
                tables=[{"name": f"Table {i}", "display_name": f"My Data {i}",
                         "start_row": 0, "start_col": 0, "end_row": 5000, "end_col": 4}
                        for i in range(1)]),
        ),
        "single_bytes_dates_100k": (
            lambda m: m.write_sheet_arrow_to_bytes(t_dates, column_formats={"d": "date"}),
        ),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--iters", type=int, default=25)
    ap.add_argument("--threshold", type=float, default=5.0,
                    help="percent slowdown that counts as a regression")
    args = ap.parse_args()

    print("=" * 74)
    print(f"perf regression: fixed vs baseline | iters={args.iters} | threshold=+{args.threshold}%")
    print("=" * 74)
    print(f"{'workload':<28}{'baseline ms':>13}{'fixed ms':>12}{'delta':>10}")
    print("-" * 74)

    wl = workloads()
    regressions = []
    for name, (fn,) in wl.items():
        if HAVE_BASELINE:
            med_base, med_fixed = bench_pair(lambda: fn(jetxl), lambda: fn(jetxl_baseline), args.iters)
            delta = (med_fixed - med_base) / med_base * 100.0
            flag = "  <== REGRESSION" if delta > args.threshold else ""
            if delta > args.threshold:
                regressions.append((name, delta))
            print(f"{name:<28}{med_base:>13.2f}{med_fixed:>12.2f}{delta:>+9.1f}%{flag}")
        else:
            med_fixed, _ = bench(lambda: fn(jetxl), args.iters)
            print(f"{name:<28}{'--':>13}{med_fixed/1e6:>12.2f}{'--':>10}")

    print("-" * 74)
    if HAVE_BASELINE:
        if regressions:
            print("RESULT: PERFORMANCE REGRESSION DETECTED")
            for n, d in regressions:
                print(f"  {n}: +{d:.1f}%")
            return 1
        print("RESULT: no regression beyond threshold on any hot path")
    else:
        print("RESULT: baseline unavailable; reported fixed-build timings only")
    print("=" * 74)
    return 0


if __name__ == "__main__":
    sys.exit(main())
