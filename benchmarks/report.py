"""Turn benchmarks/results/latest.json into the README chart and tables.

    python benchmarks/report.py                    # chart + benchmarks/results/summary.md
    python benchmarks/report.py --update-readme    # also rewrite the README section
                                                   # between the benchmark markers

Runs that crashed or failed the read-back check are left out and listed.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import math
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
START, END = "<!-- benchmark:start -->", "<!-- benchmark:end -->"

# light / dark bar colour per library
COLORS = {
    "jetxl (arrow)": ("#2ecc71", "#3ee686"),
    "jetxl (dict)": ("#27ae60", "#34c576"),
    "rustpy_xlsxwriter": ("#3498db", "#4aa8ec"),
    "pyexcelerate": ("#00bcd4", "#26d9ef"),
    "xlsxwriter": ("#e74c3c", "#f26658"),
    "polars.write_excel": ("#f39c12", "#ffb733"),
    "pandas+xlsxwriter": ("#ff5722", "#ff7a4d"),
    "openpyxl": ("#9b59b6", "#b673cf"),
    "pandas+openpyxl": ("#e91e63", "#ff4d84"),
}
# legend spacing per label (swatch + text + gap); unknown labels are estimated
LEGEND_ADV = {
    "jetxl (arrow)": 106, "jetxl (dict)": 100, "rustpy_xlsxwriter": 130,
    "pyexcelerate": 100, "xlsxwriter": 88, "pandas+xlsxwriter": 130, "openpyxl": 76,
}
# README version list: display name -> package
VERSION_NAMES = {
    "Jetxl": "jetxl", "Polars": "polars", "Pyarrow": "pyarrow", "Pandas": "pandas",
    "Pyexcelerate": "pyexcelerate", "Rustpy-xlsxwriter": "rustpy-xlsxwriter",
    "Openpyxl": "openpyxl", "Xlsxwriter": "xlsxwriter",
}

SVG_STYLE = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 760 420" width="100%" font-family="-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif" role="img" aria-label="Excel write performance: execution time">
<style>
  .ink   { fill: #768390; }
  .inkw  { stroke: #768390; }
  .grid  { stroke: #b6bec6; }
  .axis  { stroke: #768390; }
  .sub   { fill: #768390; }
  .lite  { display: inline; }
  .dark  { display: none; }
  @media (prefers-color-scheme: light), (prefers-color-scheme: no-preference) {
    .ink  { fill: #24292f; }
    .inkw { stroke: #24292f; }
    .grid { stroke: #d0d7de; }
    .axis { stroke: #57606a; }
    .sub  { fill: #57606a; }
    .lite { display: inline; }
    .dark { display: none; }
  }
  @media (prefers-color-scheme: dark) {
    .ink  { fill: #e6edf3; }
    .inkw { stroke: #e6edf3; }
    .grid { stroke: #30363d; }
    .axis { stroke: #8b949e; }
    .sub  { fill: #adbac7; }
    .lite { display: none; }
    .dark { display: inline; }
  }
  @media (prefers-reduced-motion: reduce) {
    * { animation: none !important; }
    .anim-init { opacity: 1 !important; transform: none !important;
                 stroke-dashoffset: 0 !important; }
  }

  @keyframes gU { from { transform: scaleY(0);} to { transform: scaleY(1);} }
  .bar { transform-box: fill-box; transform-origin: bottom;
         animation: gU .7s cubic-bezier(.2,.8,.2,1) both; }
</style>"""

PLOT_W, BASE_Y = 680, 314


# ---------------------------------------------------------------------------
# data
# ---------------------------------------------------------------------------

def load(path: Path):
    """Successful runs split into file output (README chart and tables) and
    in-memory output (its own table), plus a list of failed runs."""
    data = json.loads(path.read_text(encoding="utf-8"))
    good, in_memory, bad = {}, {}, []
    for r in data["results"]:
        if "median_s" in r and (r.get("check") is None or r["check"]["ok"]):
            target = in_memory if r.get("output") == "bytes" else good
            target.setdefault(r["rows"], {})[r["library"]] = r
        elif "skipped" not in r:
            reason = r.get("error") or r["check"]["error"]
            bad.append(f"{r['library']} at {r['rows']:,} rows: {reason}")
    return data["meta"], good, in_memory, bad


def fmt_rows(n: int) -> str:
    return f"{n // 1_000_000}M" if n >= 1_000_000 else f"{n // 1000}K"


def fmt_time(s: float) -> str:
    if s < 0.1:
        return f"{s:.3f}s"
    return f"{s:.2f}s"


def fmt_mem(mb: float) -> str:
    return f"{mb:.0f} MB" if mb < 1024 else f"{mb / 1024:.2f} GB"


# ---------------------------------------------------------------------------
# chart (same layout as the original README chart)
# ---------------------------------------------------------------------------

def svg_chart(good: dict) -> str:
    sizes = sorted(good)
    big = sizes[-1]
    libs = sorted(good[big], key=lambda lib: good[big][lib]["median_s"])
    times = {(n, lib): good[n][lib]["median_s"] for n in sizes for lib in libs if lib in good[n]}
    fastest = math.log10(min(times.values()))
    # reach at least slightly below the decade gridline nearest the fastest bar,
    # so the chart always has a labelled line under the shortest bars
    lo = min(fastest - 0.15, round(fastest) - 0.15)
    hi = math.log10(max(times.values())) + 0.10
    scale = BASE_Y / (hi - lo)

    def y(t):
        return BASE_Y - (math.log10(t) - lo) * scale

    def tick(t):
        if t < 1:
            return f"{t * 1000:.0f}ms"
        return f"{t:.1f}s" if t < 100 else f"{t:.0f}s"

    out = [SVG_STYLE,
           '<text x="380.0" y="26" text-anchor="middle" class="ink" font-size="16" '
           'font-weight="600">Excel write performance: execution time</text>',
           '<g transform="translate(60,46)">']
    for p in range(math.ceil(lo), math.floor(hi) + 1):
        gy = y(10.0 ** p)
        out.append(f'<line x1="0" y1="{gy:.1f}" x2="{PLOT_W}" y2="{gy:.1f}" class="grid" stroke-width="1"/>')
        out.append(f'<text x="-6" y="{gy + 3:.1f}" text-anchor="end" class="sub" '
                   f'font-size="10">{tick(10.0 ** p)}</text>')
    out.append(f'<line x1="0" y1="{BASE_Y}" x2="{PLOT_W}" y2="{BASE_Y}" class="axis" stroke-width="1"/>')

    group_w = PLOT_W / len(sizes)
    step = group_w / (len(libs) + 1)
    k = 0
    for g, n in enumerate(sizes):
        for i, lib in enumerate(libs):
            if (n, lib) not in times:
                continue
            x = g * group_w + (i + 0.5) * step
            top = y(times[n, lib])
            for cls, fill in zip(("lite", "dark"), COLORS.get(lib, ("#888888", "#aaaaaa")), strict=True):
                out.append(f'<rect class="bar {cls} anim-init" x="{x:.1f}" y="{top:.1f}" '
                           f'width="{step - 2:.1f}" height="{BASE_Y - top:.1f}" rx="3" fill="{fill}" '
                           f'style="animation-delay:{k * 0.04:.2f}s"/>')
            k += 1
        out.append(f'<text x="{(g + 0.5) * group_w:.1f}" y="332.0" text-anchor="middle" class="sub" '
                   f'font-size="11">{fmt_rows(n)}</text>')
    out.append("</g>")

    lx, ly = 60, 388
    for lib in libs:
        adv = LEGEND_ADV.get(lib, 14 + round(len(lib) * 6.2) + 12)
        if lx + adv > 740:
            lx, ly = 60, ly + 16
        light, dark = COLORS.get(lib, ("#888888", "#aaaaaa"))
        out.append(f'<rect class="lite" x="{lx}" y="{ly}" width="10" height="10" rx="2" fill="{light}"/>'
                   f'<rect class="dark" x="{lx}" y="{ly}" width="10" height="10" rx="2" fill="{dark}"/>'
                   f'<text x="{lx + 14}" y="{ly + 9}" class="sub" font-size="10">{lib}</text>')
        lx += adv
    out.append("</svg>")
    return "\n".join(out) + "\n"


# ---------------------------------------------------------------------------
# markdown
# ---------------------------------------------------------------------------

def environment_note(meta: dict) -> str:
    date = dt.datetime.fromisoformat(meta["date"]).date().isoformat()
    check = ", every output file read back and checked cell by cell" if meta.get("checked") else ""
    return (f"*Benchmark environment: Python {meta['python']}, {meta['cpu']}, {meta['logical_cpus']} logical CPUs, "
            f"{meta['ram_gb']} GB RAM, {meta['os']}. Run on {date} with "
            f"[`benchmarks/bench.py`](benchmarks/): each library and size in its own fresh process, "
            f"median of {meta['reps']} timed writes after one warmup{check}.*")


def version_list(meta: dict) -> str:
    lines = ["At the time of the test, the following library versions were used:"]
    for name, pkg in VERSION_NAMES.items():
        if meta["versions"].get(pkg):
            lines.append(f"- {name}: {meta['versions'][pkg]}")
    return "\n".join(lines)


def summary_table(good: dict) -> str:
    big = max(good)
    rows = sorted(good[big].values(), key=lambda r: r["median_s"])
    base = rows[0]["median_s"]
    out = [f"| Library | {fmt_rows(big)} Rows | Speedup | Throughput | Write Memory |",
           "|---------|---------|---------|------------|--------------|"]
    for i, r in enumerate(rows):
        ratio = r["median_s"] / base
        speed = "1.0x" if i == 0 else (f"{ratio:.1f}x slower" if ratio < 10 else f"**{ratio:.0f}x slower**")
        cells = [r["library"], f"{r['median_s']:.2f}s", speed,
                 f"{big / r['median_s'] / 1000:.0f}K rows/s", fmt_mem(r["write_mem_mb"])]
        if i == 0:
            cells = [f"**{c}**" for c in cells]
        out.append("| " + " | ".join(cells) + " |")
    return "\n".join(out)


def sizes_table(good: dict) -> str:
    sizes = sorted(good)
    big = sizes[-1]
    libs = sorted(good[big], key=lambda lib: good[big][lib]["median_s"])
    out = ["| Library | " + " | ".join(f"{fmt_rows(n)} rows" for n in sizes) + " |",
           "|---" * (len(sizes) + 1) + "|"]
    for lib in libs:
        out.append(f"| {lib} | " + " | ".join(
            fmt_time(good[n][lib]["median_s"]) if lib in good[n] else "—" for n in sizes) + " |")
    return "\n".join(out)


def in_memory_table(good: dict, in_memory: dict) -> str:
    """Each in-memory entry next to its file-writing counterpart, per size."""
    out = ["| Rows | Library | To file | To bytes | Write memory (file / bytes) |",
           "|---|---|---|---|---|"]
    for n in sorted(in_memory):
        for lib, b in sorted(in_memory[n].items()):
            base = lib.replace(", bytes)", ")")
            f = good.get(n, {}).get(base)
            file_time = fmt_time(f["median_s"]) if f else "—"
            file_mem = fmt_mem(f["write_mem_mb"]) if f else "—"
            out.append(f"| {fmt_rows(n)} | {base} | {file_time} | {fmt_time(b['median_s'])} | "
                       f"{file_mem} / {fmt_mem(b['write_mem_mb'])} |")
    return "\n".join(out)


def readme_sizes_table(good: dict) -> str:
    """Seconds per size, fastest library in bold (the README's own table format)."""
    sizes = sorted(good)
    big = sizes[-1]
    libs = sorted(good[big], key=lambda lib: good[big][lib]["median_s"])
    out = ["| Library | " + " | ".join(f"{fmt_rows(n)} rows" for n in sizes) + " |",
           "|---------|" + "|".join("-" * 10 for _ in sizes) + "|"]
    for i, lib in enumerate(libs):
        cells = [lib] + [fmt_time(good[n][lib]["median_s"]).rstrip("s") if lib in good[n] else "—"
                         for n in sizes]
        if i == 0:
            cells = [f"**{c}**" for c in cells]
        out.append("| " + " | ".join(cells) + " |")
    return "\n".join(out)


def readme_in_memory_section(meta: dict, good: dict, in_memory: dict) -> str | None:
    """README table: write_sheet_arrow vs write_sheet_arrow_to_bytes, both from
    the same run so the two columns compare directly."""
    rows = [n for n in sorted(in_memory)
            if "jetxl (arrow, bytes)" in in_memory[n] and "jetxl (arrow)" in good.get(n, {})]
    if not rows:
        return None
    date = dt.datetime.fromisoformat(meta["date"]).date().isoformat()
    out = ["### In-memory output", "",
           "`write_sheet_arrow_to_bytes` returns the workbook as `bytes` instead of writing a file, for web "
           "responses, cloud functions and object storage. Both columns come from one run on "
           f"{date}, with the same data and method as above, so they compare directly with each other. "
           "Timings vary a little between runs, so the file times can differ slightly from the tables above.",
           "",
           "| Rows | To file (`write_sheet_arrow`) | To bytes (`write_sheet_arrow_to_bytes`) "
           "| Write memory (file / bytes) |",
           "|------|------|------|------|"]
    for n in rows:
        f, b = good[n]["jetxl (arrow)"], in_memory[n]["jetxl (arrow, bytes)"]
        out.append(f"| {fmt_rows(n)} | {fmt_time(f['median_s'])} | {fmt_time(b['median_s'])} | "
                   f"{fmt_mem(f['write_mem_mb'])} / {fmt_mem(b['write_mem_mb'])} |")
    return "\n".join(out)


def readme_block(meta: dict, good: dict, in_memory_section: str | None = None) -> str:
    return "\n\n".join(part for part in [
        environment_note(meta),
        version_list(meta),
        "### Library comparison summary\n" + summary_table(good),
        "*Write Memory is the extra memory used while writing: peak resident memory during the write "
        "minus resident memory just before it. The data itself is already in memory and not counted.*",
        "The chart below shows the execution time comparison with popular libraries for a single sheet "
        "file against Jetxl. Performance will vary on different hardware.\n"
        "![exec_time_comparison](benchmark/execution_time.svg)",
        "### Execution Time (seconds)\n\n" + readme_sizes_table(good),
        in_memory_section,
    ] if part)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--results", type=Path, default=HERE / "results" / "latest.json")
    ap.add_argument("--svg", type=Path, default=ROOT / "benchmark" / "execution_time.svg")
    ap.add_argument("--no-svg", action="store_true", help="don't write the chart")
    ap.add_argument("--in-memory", type=Path, default=HERE / "results" / "in_memory.json",
                    help="jetxl file-vs-bytes results for the README's in-memory table (skipped if missing)")
    ap.add_argument("--summary", type=Path, default=HERE / "results" / "summary.md")
    ap.add_argument("--update-readme", action="store_true")
    ap.add_argument("--readme", type=Path, default=ROOT / "README.md")
    args = ap.parse_args()

    meta, good, in_memory, bad = load(args.results)
    if not good:
        raise SystemExit("no successful runs in " + str(args.results))

    if not args.no_svg:
        args.svg.write_text(svg_chart(good), encoding="utf-8", newline="\n")
    mem_section = None
    if args.in_memory.exists() and args.in_memory.resolve() != args.results.resolve():
        mem_meta, mem_good, mem_bytes, _ = load(args.in_memory)
        mem_section = readme_in_memory_section(mem_meta, mem_good, mem_bytes)
    elif in_memory:
        mem_section = readme_in_memory_section(meta, good, in_memory)
    block = readme_block(meta, good, mem_section)
    summary = block + "\n\n### All sizes (median write time)\n" + sizes_table(good) + "\n"
    if in_memory and not mem_section:
        summary += ("\n### In-memory output (median write time)\n" + in_memory_table(good, in_memory) + "\n")
    if bad:
        summary += "\n### Left out\n" + "\n".join(f"- {b}" for b in bad) + "\n"
    args.summary.write_text(summary, encoding="utf-8", newline="\n")
    print(summary)
    if not args.no_svg:
        print(f"wrote {args.svg}")
    print(f"wrote {args.summary}")

    if args.update_readme:
        text = args.readme.read_bytes().decode("utf-8")  # keep the file's own line endings
        if START not in text or END not in text:
            raise SystemExit(f"{args.readme} has no {START} ... {END} section to replace")
        nl = "\r\n" if "\r\n" in text else "\n"
        head, rest = text.split(START, 1)
        _, tail = rest.split(END, 1)
        new = f"{head}{START}{nl}{block.replace(chr(10), nl)}{nl}{END}{tail}"
        args.readme.write_bytes(new.encode("utf-8"))
        print(f"updated {args.readme}")


if __name__ == "__main__":
    main()
