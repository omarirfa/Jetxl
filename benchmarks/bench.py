"""Excel-writer benchmark: jetxl against other Python libraries.

Every (library, size) pair runs in its own fresh Python process, in an order
that rotates between sizes. Inside that process:

  1. build the same source data, then the library's native input (timed
     separately as convert_s, never part of the write time)
  2. one untimed warmup write, during which memory is sampled: the reported
     write memory is the peak resident memory during the write minus the
     resident memory just before it
  3. `--reps` timed writes; the median is the headline number
  4. the file is read back and every cell is compared with the source data

Usage:
    python benchmarks/bench.py                      # 10K, 100K, 1M rows, 5 reps
    python benchmarks/bench.py --quick              # 10K and 100K rows, 3 reps
    python benchmarks/bench.py --sizes 1000000 --libs "jetxl (arrow)" xlsxwriter
    python benchmarks/report.py                     # chart + tables from the results
"""
from __future__ import annotations

import argparse
import datetime as dt
import gc
import json
import math
import os
import platform
import shutil
import statistics
import subprocess
import sys
import tempfile
import threading
import time
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT_OUT = HERE / "results" / "latest.json"


# ---------------------------------------------------------------------------
# Child process: one library, one size
# ---------------------------------------------------------------------------

class _RssSampler(threading.Thread):
    """Polls this process's resident memory and keeps the highest value seen."""

    def __init__(self, interval: float = 0.002):
        import psutil
        super().__init__(daemon=True)
        self._proc = psutil.Process()
        self._interval = interval
        self._done = False
        self.peak = 0
        self.samples = 0

    def rss(self) -> int:
        return self._proc.memory_info().rss

    def run(self):
        # time.sleep, not Event.wait: on Windows only sleep uses the high-resolution
        # timer (Python 3.11+); Event.wait is limited to ~15 ms steps
        while not self._done:
            self.peak = max(self.peak, self.rss())
            self.samples += 1
            time.sleep(self._interval)

    def stop(self):
        self._done = True
        self.join()
        self.peak = max(self.peak, self.rss())


def _process_peak_mb() -> float:
    """Peak resident memory of the whole process so far, as the OS reports it."""
    if sys.platform == "win32":
        import psutil
        return psutil.Process().memory_info().peak_wset / 2**20
    import resource
    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return peak / 2**20 if sys.platform == "darwin" else peak / 1024


def run_child(lib: str, rows: int, reps: int, check: bool) -> dict:
    import libs

    packages, build_input, write = libs.WRITERS[lib]
    try:
        versions = {p: version(p) for p in packages}
    except PackageNotFoundError as e:
        return {"library": lib, "rows": rows, "skipped": f"{e.name} is not installed"}

    src = libs.make_source(rows)
    t0 = time.perf_counter()
    data = build_input(src)
    convert_s = time.perf_counter() - t0

    tmp = tempfile.mkdtemp(prefix="jetxl-bench-")
    path = os.path.join(tmp, "out.xlsx")
    try:
        # warmup write, also used for the memory measurement
        gc.collect()
        sampler = _RssSampler()
        before = sampler.rss()
        sampler.start()
        out = write(data, path)  # in-memory writers return bytes, still held here
        sampler.stop()
        write_mem_mb = max(0, sampler.peak - before) / 2**20

        times = []
        for _ in range(reps):
            out = None
            if os.path.exists(path):
                os.remove(path)
            gc.collect()
            t0 = time.perf_counter()
            out = write(data, path)
            times.append(time.perf_counter() - t0)
        if out is not None:
            # in-memory output: save it outside the timer for the size and the check
            with open(path, "wb") as f:
                f.write(out)
            out = None

        record = {
            "library": lib,
            "rows": rows,
            "output": "bytes" if lib in libs.BYTES_OUTPUT else "file",
            "times_s": times,
            "median_s": statistics.median(times),
            "min_s": min(times),
            "max_s": max(times),
            "convert_s": convert_s,
            "write_mem_mb": write_mem_mb,
            "process_peak_mb": _process_peak_mb(),
            "file_bytes": os.path.getsize(path),
            "versions": versions,
        }
        del data
        gc.collect()
        record["check"] = libs.check_file(path, src) if check else None
        return record
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ---------------------------------------------------------------------------
# Parent process
# ---------------------------------------------------------------------------

def _cpu_name() -> str:
    try:
        if sys.platform == "win32":
            import winreg
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                 r"HARDWARE\DESCRIPTION\System\CentralProcessor\0")
            return winreg.QueryValueEx(key, "ProcessorNameString")[0].strip()
        if sys.platform == "darwin":
            return subprocess.check_output(["sysctl", "-n", "machdep.cpu.brand_string"], text=True).strip()
        for line in Path("/proc/cpuinfo").read_text().splitlines():
            if line.startswith("model name"):
                return line.split(":", 1)[1].strip()
    except Exception:
        pass
    return platform.processor() or "unknown"


def environment(args) -> dict:
    import psutil
    pkgs = sorted({p for lib in args.libs for p in WRITER_PACKAGES[lib]} | {"numpy", "python-calamine"})
    versions = {}
    for p in pkgs:
        try:
            versions[p] = version(p)
        except PackageNotFoundError:
            versions[p] = None
    return {
        "date": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "os": f"{platform.system()} {platform.release()}",
        "os_detail": platform.platform(),
        "cpu": _cpu_name(),
        "logical_cpus": os.cpu_count(),
        # installed RAM: the OS reports slightly less than the modules' size
        "ram_gb": math.ceil(psutil.virtual_memory().total / 2**30),
        "python": platform.python_version(),
        "versions": versions,
        "sizes": args.sizes,
        "reps": args.reps,
        "warmup": 1,
        "columns": 10,
        "checked": not args.no_check,
    }


def _save(out: Path, meta: dict, results: list):
    out.parent.mkdir(parents=True, exist_ok=True)
    tmp = out.with_suffix(".tmp")
    tmp.write_text(json.dumps({"meta": meta, "results": results}, indent=2), encoding="utf-8")
    tmp.replace(out)


def _fmt(sec: float) -> str:
    return f"{sec * 1000:.1f}ms" if sec < 1 else f"{sec:.2f}s"


def main():
    import libs

    global WRITER_PACKAGES
    WRITER_PACKAGES = {name: w[0] for name, w in libs.WRITERS.items()}

    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--sizes", type=int, nargs="+", default=[10_000, 100_000, 1_000_000])
    ap.add_argument("--reps", type=int, default=5, help="timed writes per library and size (median is reported)")
    ap.add_argument("--libs", nargs="+", default=list(libs.WRITERS), choices=list(libs.WRITERS))
    ap.add_argument("--out", type=Path, default=DEFAULT_OUT)
    ap.add_argument("--no-check", action="store_true", help="skip reading the files back")
    ap.add_argument("--timeout", type=int, default=3600, help="seconds allowed per library and size")
    ap.add_argument("--quick", action="store_true", help="10K and 100K rows, 3 reps")
    args = ap.parse_args()
    if args.quick:
        args.sizes, args.reps = [10_000, 100_000], 3

    meta = environment(args)
    print(f"{meta['cpu']} | {meta['logical_cpus']} CPUs | {meta['ram_gb']} GB | Python {meta['python']}")
    print(f"sizes {args.sizes} | {args.reps} timed reps + 1 warmup | read-back check: {meta['checked']}\n")

    results = []
    for si, rows in enumerate(args.sizes):
        order = args.libs[si % len(args.libs):] + args.libs[:si % len(args.libs)]
        for lib in order:
            cmd = [sys.executable, str(HERE / "bench.py"), "--child", lib, str(rows), str(args.reps)]
            if args.no_check:
                cmd.append("--no-check")
            try:
                p = subprocess.run(cmd, cwd=HERE, capture_output=True, text=True, timeout=args.timeout)
                if p.returncode != 0:
                    err = p.stderr.strip()
                    raise RuntimeError(err.splitlines()[-1] if err else f"exit {p.returncode}")
                rec = json.loads(p.stdout.strip().splitlines()[-1])
            except Exception as e:
                rec = {"library": lib, "rows": rows, "error": f"{type(e).__name__}: {e}"}
            results.append(rec)
            _save(args.out, meta, results)

            if "median_s" in rec:
                chk = rec["check"]
                status = "check skipped" if chk is None else ("ok" if chk["ok"] else f"FAILED CHECK: {chk['error']}")
                print(f"{rows:>9,}  {lib:<20} {_fmt(rec['median_s']):>9}  "
                      f"write mem {rec['write_mem_mb']:7.0f} MB  {status}", flush=True)
            else:
                print(f"{rows:>9,}  {lib:<20} {rec.get('skipped') or rec.get('error')}", flush=True)

    print(f"\nwrote {args.out}")
    failed = [r for r in results if r.get("error") or (r.get("check") and not r["check"]["ok"])]
    if failed:
        print(f"{len(failed)} run(s) failed or wrote wrong data; report.py leaves them out of the chart.")


WRITER_PACKAGES: dict = {}

if __name__ == "__main__":
    if len(sys.argv) >= 5 and sys.argv[1] == "--child":
        out = run_child(sys.argv[2], int(sys.argv[3]), int(sys.argv[4]), "--no-check" not in sys.argv)
        print(json.dumps(out))
    else:
        main()
