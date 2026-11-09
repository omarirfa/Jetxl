*Benchmark environment: Python 3.13.1, AMD Ryzen 9 7900X 12-Core Processor, 24 logical CPUs, 64 GB RAM, Windows 10. Run on 2026-11-09 with [`benchmarks/bench.py`](benchmarks/): each library and size in its own fresh process, median of 5 timed writes after one warmup, every output file read back and checked cell by cell.*

At the time of the test, the following library versions were used:
- Jetxl: 0.3.3
- Polars: 2.1.0
- Pyarrow: 26.0.1
- Pandas: 3.1.1
- Pyexcelerate: 0.13.0
- Rustpy-xlsxwriter: 0.9.1
- Openpyxl: 3.1.5
- Xlsxwriter: 3.3.1

### Library comparison summary
| Library | 1M Rows | Speedup | Throughput | Write Memory |
|---------|---------|---------|------------|--------------|
| **jetxl (arrow)** | **1.18s** | **1.0x** | **851K rows/s** | **939 MB** |
| jetxl (dict) | 2.50s | 2.1x slower | 401K rows/s | 1021 MB |
| rustpy_xlsxwriter | 6.31s | 5.4x slower | 159K rows/s | 2 MB |
| pyexcelerate | 29.53s | **25x slower** | 34K rows/s | 175 MB |
| xlsxwriter | 35.89s | **31x slower** | 28K rows/s | 7 MB |
| polars.write_excel | 47.28s | **40x slower** | 21K rows/s | 2.53 GB |
| openpyxl | 63.51s | **54x slower** | 16K rows/s | 8 MB |
| pandas+xlsxwriter | 77.33s | **66x slower** | 13K rows/s | 1.20 GB |
| pandas+openpyxl | 112.36s | **96x slower** | 9K rows/s | 3.93 GB |

*Write Memory is the extra memory used while writing: peak resident memory during the write minus resident memory just before it. The data itself is already in memory and not counted.*

The chart below shows the execution time comparison with popular libraries for a single sheet file against Jetxl. Performance will vary on different hardware.
![exec_time_comparison](benchmark/execution_time.svg)

### Execution Time (seconds)

| Library | 10K rows | 100K rows | 1M rows |
|---------|----------|----------|----------|
| **jetxl (arrow)** | **0.016** | **0.099** | **1.18** |
| jetxl (dict) | 0.028 | 0.22 | 2.50 |
| rustpy_xlsxwriter | 0.051 | 0.58 | 6.31 |
| pyexcelerate | 0.31 | 3.30 | 29.53 |
| xlsxwriter | 0.34 | 4.00 | 35.89 |
| polars.write_excel | 0.44 | 4.71 | 47.28 |
| openpyxl | 0.70 | 6.26 | 63.51 |
| pandas+xlsxwriter | 0.76 | 7.01 | 77.33 |
| pandas+openpyxl | 1.15 | 11.30 | 112.36 |

### All sizes (median write time)
| Library | 10K rows | 100K rows | 1M rows |
|---|---|---|---|
| jetxl (arrow) | 0.016s | 0.099s | 1.18s |
| jetxl (dict) | 0.028s | 0.22s | 2.50s |
| rustpy_xlsxwriter | 0.051s | 0.58s | 6.31s |
| pyexcelerate | 0.31s | 3.30s | 29.53s |
| xlsxwriter | 0.34s | 4.00s | 35.89s |
| polars.write_excel | 0.44s | 4.71s | 47.28s |
| openpyxl | 0.70s | 6.26s | 63.51s |
| pandas+xlsxwriter | 0.76s | 7.01s | 77.33s |
| pandas+openpyxl | 1.15s | 11.30s | 112.36s |
