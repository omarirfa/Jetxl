# jetxl test suite

All tests require the built `jetxl` wheel installed, plus:

```bash
pip install pyarrow polars pandas openpyxl xmlschema
```

Each file is a standalone script; run it with `python <file>` from this directory.

## Functional suites

| File | What it covers |
|------|----------------|
| `test_suite.py` | core coverage of the public API |
| `test_function_matrix.py` | every write function × DataFrame type (polars/pandas/pyarrow) × path, categoricals, direct PyCapsule input, column-name resolution |
| `test_kitchen_sink.py` | broad feature coverage in combination |
| `test_max_limits.py` | Excel grid-limit enforcement |
| `test_reference_parity.py` | cross-path output parity |
| `test_ecma_conformance.py` | validates output against the official ECMA-376 XSDs |

## Fuzzers (property-based)

| File | Strategy |
|------|----------|
| `fuzz_jetxl.py` | single-sheet: generates randomized valid workbooks and asserts no-crash / loadable-in-openpyxl / ECMA-schema-conformant |
| `fuzz_jetxl_expanded.py` | composite Arrow types (struct/list/map/binary/decimal/…), multi-sheet, and dict paths |

Set `JETXL_FUZZ_SEED=<n>` for reproducible runs. Pass an iteration count as the
first arg, e.g. `python fuzz_jetxl.py 400`.

## ECMA-376 schemas

`test_ecma_conformance.py` and the schema-validation step of the fuzzers need the
official ECMA-376 XSDs. By default they look under `tests/ecma/schemas_transitional`
(and `tests/ecma/schemas_opc`); override with `JETXL_ECMA_SCHEMAS` /
`JETXL_ECMA_OPC`. If the schemas or `xmlschema` aren't present, these tests skip
the schema check gracefully (the fuzzers still assert no-crash + loadability).

## Exercising the parallel single-sheet path

The large-single-sheet row parallelism activates only above ~50k rows with more
than one worker thread. On a machine with few cores you can force it by setting
`RAYON_NUM_THREADS=4` before running a script that writes a large sheet.
