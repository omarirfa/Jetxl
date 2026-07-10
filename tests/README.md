# jetxl test suite

All tests require the built `jetxl` wheel installed, plus:

```bash
pip install pyarrow polars pandas openpyxl xmlschema
```

## Functional suites

| File | What it covers |
|------|----------------|
| `test_function_matrix.py` | every write function × DataFrame type (polars/pandas/pyarrow) × path, categoricals, direct PyCapsule input, column-name resolution |
| `test_edge_robustness.py` | corruption / edge-case / panic guards (uint64 overflow, reversed ranges, invalid colors, NaN/Inf, pre-1900 dates, merged-cell hyperlinks, …) |
| `test_kitchen_sink.py` | broad feature coverage in combination |
| `test_max_limits.py` | Excel grid-limit enforcement |
| `test_reference_parity.py` | cross-path output parity |
| `test_v1_stability.py` | stability / API surface |
| `test_readme_conformance.py` | verifies documented examples behave as described |
| `test_ecma_conformance.py` | validates output against the official ECMA-376 XSDs |
| `test_parallel_rows.py` | parallel single-sheet row generation: determinism + byte-for-byte parity with the serial path |
| `test_suite.py`, `test_features.py`, `test_jetxl.py` | additional/legacy coverage |
| `silent_drop_audit.py` | asserts every feature option takes effect on every path (no silent drops) |

## Fuzzers (property-based / differential)

| File | Strategy |
|------|----------|
| `fuzz_jetxl.py` | single-sheet: generates randomized valid workbooks and asserts no-crash / loadable-in-openpyxl / ECMA-schema-conformant |
| `fuzz_jetxl_expanded.py` | composite Arrow types (struct/list/map/binary/decimal/…), multi-sheet, and dict paths |
| `fuzz_diff.py` | differential value-fidelity: write → read back → assert every cell value round-trips |

Set `JETXL_FUZZ_SEED=<n>` for reproducible runs. Pass an iteration count as the
first arg, e.g. `python fuzz_diff.py 400`.

## Performance

`perf_regression.py` runs an interleaved best-of-N comparison and expects a
`jetxl_baseline` wheel (a pristine/previous build installed under that import
name) to compare the current build against:

```bash
python perf_regression.py --iters 20
```

`jetxl_benchmark.py` is a standalone throughput benchmark (no baseline needed).

## ECMA-376 schemas

`test_ecma_conformance.py` and the schema-validation step of the fuzzers need the
official ECMA-376 XSDs. By default they look under `tests/ecma/schemas_transitional`
(and `tests/ecma/schemas_opc`); override with `JETXL_ECMA_SCHEMAS` /
`JETXL_ECMA_OPC`. If the schemas or `xmlschema` aren't present, these tests skip
the schema check gracefully (the fuzzers still assert no-crash + loadability).

## Exercising the parallel single-sheet path

The large-single-sheet row parallelism activates only above ~50k rows with more
than one worker thread. On any machine you can force it with:

```bash
RAYON_NUM_THREADS=4 python tests/test_parallel_rows.py
```
