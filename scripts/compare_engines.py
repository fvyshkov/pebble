"""Direct comparison: Python vs Rust formula engine on same model.

Runs both engines against the same DB, compares results cell-by-cell.
Usage: cd /path/to/pebble-rust-calc && python scripts/compare_engines.py <model_id>
"""
import asyncio
import os
import sys
import time

# Ensure we can import backend
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from backend.db import get_db, init_db
import backend.formula_engine as fe


def vals_equal(a: str, b: str, tol=1e-6) -> bool:
    if a == b:
        return True
    try:
        fa, fb = float(a), float(b)
        if fa == fb:
            return True
        if fa == 0 and fb == 0:
            return True
        denom = max(abs(fa), abs(fb))
        if denom > 1e-15:
            return abs(fa - fb) / denom < tol
    except (ValueError, TypeError):
        pass
    return False


async def main():
    if len(sys.argv) < 2:
        print("Usage: python scripts/compare_engines.py <model_id>")
        sys.exit(1)

    model_id = sys.argv[1]
    await init_db()
    db = get_db()

    # --- Python engine ---
    print("Running Python engine...")
    fe._USE_RUST = False
    t0 = time.perf_counter()
    py_result = await fe.calculate_model(db, model_id)
    t1 = time.perf_counter()
    py_cells = sum(len(v) for v in py_result.values())
    print(f"  Python: {t1 - t0:.3f}s, {py_cells} cells changed")

    # --- Rust V2 engine ---
    print("Running Rust V2 engine...")
    fe._USE_RUST = True
    fe._ENGINE_MODE = "rust_v2"
    if fe._rust_engine is None:
        print("ERROR: pebble_calc not installed")
        sys.exit(1)
    t0 = time.perf_counter()
    rs_result = await fe.calculate_model(db, model_id)
    t1 = time.perf_counter()
    rs_cells = sum(len(v) for v in rs_result.values())
    print(f"  Rust V2: {t1 - t0:.3f}s, {rs_cells} cells changed")

    # --- Compare ---
    print("\n--- Comparison ---")
    all_keys = set()
    for sid, changes in py_result.items():
        for ck in changes:
            all_keys.add((sid, ck))
    for sid, changes in rs_result.items():
        for ck in changes:
            all_keys.add((sid, ck))

    mismatches = []
    py_only = 0
    rs_only = 0
    matched = 0

    for sid, ck in sorted(all_keys):
        py_val = py_result.get(sid, {}).get(ck)
        rs_val = rs_result.get(sid, {}).get(ck)

        if py_val is None and rs_val is not None:
            rs_only += 1
            mismatches.append((sid, ck, "RUST_ONLY", "", rs_val))
        elif rs_val is None and py_val is not None:
            py_only += 1
            mismatches.append((sid, ck, "PY_ONLY", py_val, ""))
        elif not vals_equal(py_val, rs_val):
            mismatches.append((sid, ck, "DIFF", py_val, rs_val))
        else:
            matched += 1

    print(f"Matched:    {matched}")
    print(f"Python-only: {py_only}")
    print(f"Rust-only:   {rs_only}")
    diff_count = len(mismatches) - py_only - rs_only
    print(f"Value diff:  {diff_count}")

    if mismatches:
        # Get sheet names for readable output
        sheets = await db.execute_fetchall(
            "SELECT id, name FROM sheets WHERE model_id = ?", (model_id,))
        sname = {s["id"]: s["name"] for s in sheets}

        # Analyze the magnitudes of differences
        rel_diffs = []
        for sid, ck, kind, pv, rv in mismatches:
            if kind == "DIFF":
                try:
                    fp, fr = float(pv), float(rv)
                    denom = max(abs(fp), abs(fr), 1e-15)
                    rel_diffs.append(abs(fp - fr) / denom)
                except (ValueError, TypeError):
                    pass

        if rel_diffs:
            rel_diffs.sort()
            print(f"\nRelative differences (DIFF only):")
            print(f"  min: {min(rel_diffs):.2e}")
            print(f"  max: {max(rel_diffs):.2e}")
            print(f"  median: {rel_diffs[len(rel_diffs)//2]:.2e}")

        print(f"\nFirst 50 mismatches:")
        for sid, ck, kind, pv, rv in mismatches[:50]:
            sheet_label = sname.get(sid, sid[:8])
            print(f"  {kind:10s} | {sheet_label[:30]:30s} | {ck[:50]:50s} | py={pv} rs={rv}")

    total = len(all_keys)
    total_diff = len(mismatches)
    pct = (total_diff / total * 100) if total else 0
    print(f"\nTotal: {total} cells, {total_diff} differences ({pct:.2f}%)")

    if total_diff == 0:
        print("PERFECT MATCH!")
    elif pct < 1.0:
        print("CLOSE MATCH — investigate remaining differences")
    else:
        print("SIGNIFICANT DIFFERENCES — needs debugging")


if __name__ == "__main__":
    asyncio.run(main())
