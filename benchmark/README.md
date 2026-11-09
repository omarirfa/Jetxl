# Benchmarks

Compares how long jetxl and other Python libraries take to write the same data to an
`.xlsx` file, and how much extra memory each one uses while writing. The chart and
table in the main README are generated from `results/latest.json`.

## Run it

```bash
python -m venv .venv-bench
.venv-bench/bin/pip install -r benchmark/requirements.txt   # Windows: .venv-bench\Scripts\pip
.venv-bench/bin/python benchmark/bench.py          # 10K, 100K and 1M rows, 5 timed reps
.venv-bench/bin/python benchmark/report.py         # chart + benchmark/results/summary.md
```

`report.py --update-readme` also rewrites the section of the main README between
`<!-- benchmark:start -->` and `<!-- benchmark:end -->`, and the chart at
`benchmark/execution_time.svg`.

Other options: `--quick` (10K and 100K rows, 3 reps), `--sizes`, `--reps`,
`--libs "jetxl (arrow)" xlsxwriter`, `--no-check`. A full run takes about an hour,
mostly the slowest libraries at 1M rows.

## What is measured

The data is 10 columns, the same values for every library, generated with a fixed seed:
three text columns (10,000, 1,000 and 5 distinct values), three integer columns, two
float columns, a date column and a date-time column, plus a header row.

Every library and size runs in its own fresh Python process, and the order of the
libraries rotates between sizes. Inside that process:

1. The library's input is built from the source data before any timing starts, in the
   form its users would already have it. The cost is recorded as `convert_s` but is not
   part of the write time.
2. One untimed warmup write. Memory is sampled every few milliseconds during it (by a
   separate thread, so it does not slow the timed writes): **write memory** is
   the peak resident memory during the write minus the resident memory just before it.
   The data is already in memory and is not counted.
3. `--reps` timed writes to a temporary folder. The **median** is the reported time.
4. The file is read back with `python-calamine` and every cell is compared with the
   source data: text exactly, numbers to 12 significant digits, and dates must come back
   as dates rather than bare numbers. A run that fails is left out of the chart and
   listed in `results/summary.md`.

## How each library is used

Each library writes a plain sheet using its fastest documented approach:

| Library | Input | How it writes |
|---|---|---|
| jetxl (arrow) | PyArrow table | `write_sheet_arrow` |
| jetxl (arrow, bytes) | PyArrow table | `write_sheet_arrow_to_bytes` (in-memory output, kept out of the README chart; see below) |
| jetxl (dict) | dict of Python lists | `write_sheet` |
| rustpy_xlsxwriter | Polars DataFrame | `write_worksheet` |
| xlsxwriter | list of rows | `constant_memory` mode, one `write_row` per row, a default date format |
| polars.write_excel | Polars DataFrame | `write_excel(autofilter=False)` |
| pyexcelerate | list of rows | `new_sheet(data=...)`, with a date style on the two date columns |
| openpyxl | list of rows | `write_only` mode, one `append` per row |
| pandas+xlsxwriter | pandas DataFrame | `to_excel(engine="xlsxwriter")` |
| pandas+openpyxl | pandas DataFrame | `to_excel(engine="openpyxl")` |

Things that are not perfectly equal, and why:

- `polars.write_excel` always wraps the data in an Excel table; polars has no option to
  turn that off.
- pandas uses openpyxl's normal mode, not `write_only`, because that is what
  `to_excel` does.
- pyexcelerate's bulk `data=` path writes dates as plain numbers, so the benchmark adds
  its documented column style to make them dates (about 10% slower than without).

## In-memory output

`jetxl (arrow, bytes)` times `write_sheet_arrow_to_bytes`. The returned bytes are saved
to disk after the timer stops, so the same read-back check applies. It is reported in
its own table next to `write_sheet_arrow`, not in the README chart, which compares
writing to a file. There is no dict version because jetxl has no dict bytes function.
To time only the two jetxl output paths:

```bash
python benchmark/bench.py --libs "jetxl (arrow)" "jetxl (arrow, bytes)" --out benchmark/results/in_memory.json
python benchmark/report.py --results benchmark/results/in_memory.json --no-svg --summary benchmark/results/in_memory_summary.md
```

## Files

- `bench.py`: runs the benchmark and writes `results/latest.json`
- `libs.py`: the data, each library's input and writer, and the read-back check
- `report.py`: builds the README chart and tables from the results
- `requirements.txt`: the exact versions used for the committed results
- `execution_time.svg`: the chart shown in the main README, written by `report.py`
- `results/`: raw results (`latest.json`, `in_memory.json`) and the generated summaries
