#!/usr/bin/env python3
"""
num_threads argument — behavior and safety guard (single-sheet writers).

Verifies the contract:
  * default and num_threads="auto" print `Using num_threads as N` (auto only).
  * an explicit integer prints nothing.
  * every num_threads value produces byte-identical worksheet output (thread
    count is a performance knob, never a correctness one).
  * invalid values raise a clean error.
  * both single-sheet functions (file + bytes) honor it.
"""
from __future__ import annotations

import contextlib
import io
import sys
import tempfile
import os
import zipfile

import pyarrow as pa

import jetxl

_p = _f = 0


def check(desc, cond):
    global _p, _f
    if cond:
        _p += 1
    else:
        _f += 1
        print(f"  FAIL {desc}")


def capture_stdout(fn):
    """Capture stdout at the OS fd level, since Rust's println! writes to fd 1
    directly (Python's contextlib.redirect_stdout only swaps sys.stdout and would
    miss it)."""
    import os
    r, w = os.pipe()
    old = os.dup(1)
    os.dup2(w, 1)
    try:
        result = fn()
    finally:
        os.dup2(old, 1)
        os.close(w)
        os.close(old)
    out = b""
    while True:
        try:
            chunk = os.read(r, 65536)
        except OSError:
            break
        if not chunk:
            break
        out += chunk
        if len(chunk) < 65536:
            break
    os.close(r)
    return result, out.decode("utf-8", "replace")


def ws(b):
    return zipfile.ZipFile(io.BytesIO(b)).read("xl/worksheets/sheet1.xml")


def make_table(n=100_000):
    return pa.table({
        "i": pa.array(range(n), pa.int64()),
        "f": pa.array([x * 1.5 for x in range(n)], pa.float64()),
        "s": pa.array([f"v{x}" for x in range(n)]),
        "b": pa.array([x % 2 == 0 for x in range(n)]),
    })


def test_print_behavior():
    t = make_table(1000)
    # default -> prints auto
    _, out = capture_stdout(lambda: jetxl.write_sheet_arrow_to_bytes(t))
    check("default prints 'Using num_threads as'", "Using num_threads as" in out)
    # explicit "auto" -> prints
    _, out = capture_stdout(lambda: jetxl.write_sheet_arrow_to_bytes(t, num_threads="auto"))
    check("num_threads='auto' prints", "Using num_threads as" in out)
    # explicit int -> no print
    _, out = capture_stdout(lambda: jetxl.write_sheet_arrow_to_bytes(t, num_threads=4))
    check("num_threads=4 does not print", "Using num_threads" not in out)
    _, out = capture_stdout(lambda: jetxl.write_sheet_arrow_to_bytes(t, num_threads=1))
    check("num_threads=1 does not print", "Using num_threads" not in out)


def test_byte_identity_across_threads():
    t = make_table(120_000)  # above the 50k parallel threshold

    def silent_bytes(**kw):
        b, _ = capture_stdout(lambda: jetxl.write_sheet_arrow_to_bytes(t, **kw))
        return ws(b)

    ref = silent_bytes(num_threads=1)
    for nt in [None, "auto", 2, 4, 8]:
        kw = {} if nt is None else {"num_threads": nt}
        check(f"num_threads={nt} worksheet identical to serial", silent_bytes(**kw) == ref)


def test_file_path():
    t = make_table(60_000)
    # file path with explicit int -> no print, valid file
    p = tempfile.mktemp(suffix=".xlsx")
    try:
        _, out = capture_stdout(lambda: jetxl.write_sheet_arrow(t, p, num_threads=2))
        check("file path num_threads=2 no print", "Using num_threads" not in out)
        check("file path produced a file", os.path.getsize(p) > 0)
    finally:
        if os.path.exists(p):
            os.unlink(p)
    # file path default -> prints
    p = tempfile.mktemp(suffix=".xlsx")
    try:
        _, out = capture_stdout(lambda: jetxl.write_sheet_arrow(t, p))
        check("file path default prints", "Using num_threads as" in out)
    finally:
        if os.path.exists(p):
            os.unlink(p)


def test_invalid_values():
    t = make_table(100)
    for bad in ["fast", "many", 3.5, [4]]:
        raised = False
        try:
            with contextlib.redirect_stdout(io.StringIO()):
                jetxl.write_sheet_arrow_to_bytes(t, num_threads=bad)
        except Exception:
            raised = True
        check(f"num_threads={bad!r} raises", raised)
    # bool is rejected (not silently treated as 0/1)
    raised = False
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            jetxl.write_sheet_arrow_to_bytes(t, num_threads=True)
    except Exception:
        raised = True
    check("num_threads=True raises (bool rejected)", raised)


def main():
    print("=" * 66)
    print("num_threads argument — behavior + byte-identity guard")
    print("=" * 66)
    test_print_behavior()
    test_byte_identity_across_threads()
    test_file_path()
    test_invalid_values()
    print("-" * 66)
    print(f"TOTAL: {_p} passed, {_f} failed")
    print("=" * 66)
    return 1 if _f else 0


if __name__ == "__main__":
    sys.exit(main())
