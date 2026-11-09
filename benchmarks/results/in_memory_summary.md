*Benchmark environment: Python 3.13.1, AMD Ryzen 9 7900X 12-Core Processor, 24 logical CPUs, 64 GB RAM, Windows 10. Run on 2026-11-09 with [`benchmarks/bench.py`](benchmarks/): each library and size in its own fresh process, median of 5 timed writes after one warmup, every output file read back and checked cell by cell.*

At the time of the test, the following library versions were used:
- Jetxl: 0.3.3
- Pyarrow: 26.0.1

### Library comparison summary
| Library | 1M Rows | Speedup | Throughput | Write Memory |
|---------|---------|---------|------------|--------------|
| **jetxl (arrow)** | **1.23s** | **1.0x** | **813K rows/s** | **941 MB** |

*Write Memory is the extra memory used while writing: peak resident memory during the write minus resident memory just before it. The data itself is already in memory and not counted.*

The chart below shows the execution time comparison with popular libraries for a single sheet file against Jetxl. Performance will vary on different hardware.
![exec_time_comparison](benchmark/execution_time.svg)

### Execution Time (seconds)

| Library | 10K rows | 100K rows | 1M rows |
|---------|----------|----------|----------|
| **jetxl (arrow)** | **0.015** | **0.10** | **1.23** |

### All sizes (median write time)
| Library | 10K rows | 100K rows | 1M rows |
|---|---|---|---|
| jetxl (arrow) | 0.015s | 0.10s | 1.23s |

### In-memory output (median write time)
| Rows | Library | To file | To bytes | Write memory (file / bytes) |
|---|---|---|---|---|
| 10K | jetxl (arrow) | 0.015s | 0.014s | 9 MB / 11 MB |
| 100K | jetxl (arrow) | 0.10s | 0.10s | 89 MB / 115 MB |
| 1M | jetxl (arrow) | 1.23s | 1.13s | 941 MB / 940 MB |
