"""Benchmark Python vs Rust formula engines.

Runs full model calculation multiple times with each engine and reports timing.

Usage: python scripts/benchmark_engine.py <model_id> [--runs N]
"""
import asyncio
import os
import sys
import time
import statistics

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from backend.db import init_db, get_db
import backend.formula_engine as fe


async def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("model_id")
    parser.add_argument("--runs", type=int, default=5)
    parser.add_argument("--engine", choices=["both", "python", "rust"], default="both")
    args = parser.parse_args()

    await init_db()
    db = get_db()
    model_id = args.model_id
    n_runs = args.runs

    print(f"Benchmarking model {model_id}, {n_runs} runs per engine\n")

    engine = args.engine
    py_median = py_cells = rs_median = rs_cells = 0

    if engine in ("both", "python"):
        print("=== Python Engine ===")
        fe._USE_RUST = False
        py_times = []
        for i in range(n_runs):
            t0 = time.perf_counter()
            result = await fe.calculate_model(db, model_id)
            t1 = time.perf_counter()
            elapsed = t1 - t0
            py_times.append(elapsed)
            py_cells = sum(len(v) for v in result.values())
            print(f"  Run {i+1}: {elapsed:.3f}s ({py_cells} cells)")

        py_median = statistics.median(py_times)
        print(f"  Median: {py_median:.3f}s, Min: {min(py_times):.3f}s, Max: {max(py_times):.3f}s")

    if engine in ("both", "rust"):
        print("\n=== Rust Engine ===")
        fe._USE_RUST = True
        if fe._rust_engine is None:
            print("ERROR: pebble_calc not installed")
            sys.exit(1)
        rs_times = []
        for i in range(n_runs):
            t0 = time.perf_counter()
            result = await fe.calculate_model(db, model_id)
            t1 = time.perf_counter()
            elapsed = t1 - t0
            rs_times.append(elapsed)
            rs_cells = sum(len(v) for v in result.values())
            print(f"  Run {i+1}: {elapsed:.3f}s ({rs_cells} cells)")

        rs_median = statistics.median(rs_times)
        print(f"  Median: {rs_median:.3f}s, Min: {min(rs_times):.3f}s, Max: {max(rs_times):.3f}s")

    # --- Summary ---
    if engine == "both":
        print("\n=== Summary ===")
        print(f"  Python: {py_median:.3f}s median, {py_cells} cells")
        print(f"  Rust:   {rs_median:.3f}s median, {rs_cells} cells")
        speedup = py_median / rs_median if rs_median > 0 else float('inf')
        print(f"  Speedup: {speedup:.2f}x")


if __name__ == "__main__":
    asyncio.run(main())
