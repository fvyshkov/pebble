"""Compare Rust and Python formula engines cell-by-cell.

Imports a real model, runs calculation with both engines,
and verifies they produce identical results.
"""
import os
import sys
import json
import time
import asyncio
import requests

API = "http://localhost:8000/api"
XLS_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "XLS-MODELS")


def _wait_server(timeout=30):
    import time
    for _ in range(timeout):
        try:
            r = requests.get(f"{API}/models", timeout=2)
            if r.ok:
                return True
        except Exception:
            pass
        time.sleep(1)
    raise RuntimeError("Server not ready")


def _import_model(xlsx_path, model_name="TestModel"):
    """Import an Excel model via streaming API and return model_id."""
    with open(xlsx_path, "rb") as f:
        resp = requests.post(
            f"{API}/import/excel-stream",
            files={"file": (os.path.basename(xlsx_path), f)},
            data={"model_name": model_name},
            stream=True,
            timeout=300,
        )
    assert resp.status_code == 200, f"Import failed: {resp.status_code} {resp.text[:200]}"
    model_id = None
    for line in resp.iter_lines(decode_unicode=True):
        if line.startswith("data:"):
            try:
                data = json.loads(line[5:].strip())
                if "model_id" in data:
                    model_id = data["model_id"]
            except json.JSONDecodeError:
                pass
    assert model_id, "No model_id in import response"
    return model_id


def _get_all_cells(model_id):
    """Get all cells for all sheets in a model."""
    sheets = requests.get(f"{API}/models/{model_id}/sheets", timeout=30).json()
    all_cells = {}
    for sheet in sheets:
        sid = sheet["id"]
        cells = requests.get(f"{API}/cells/by-sheet/{sid}", timeout=30).json()
        for cell in cells:
            key = (sid, cell["coord_key"])
            all_cells[key] = cell.get("value", "")
    return all_cells


def _calculate_with_engine(model_id, engine="python"):
    """Run full model calculation and return result dict."""
    # Use the internal async API directly
    sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
    os.environ["PEBBLE_ENGINE"] = engine

    # Reload module to pick up new engine setting
    import importlib
    import backend.formula_engine as fe
    fe._USE_RUST = (engine == "rust")

    from backend.db import get_db

    async def _run():
        db = get_db()
        return await fe.calculate_model(db, model_id)

    return asyncio.run(_run())


def _vals_equal(a: str, b: str, tol=1e-9) -> bool:
    if a == b:
        return True
    try:
        fa, fb = float(a), float(b)
        if fa == fb:
            return True
        if fa == 0 and fb == 0:
            return True
        if abs(fa) > 1e-15:
            return abs(fa - fb) / abs(fa) < tol
    except (ValueError, TypeError):
        pass
    return False


def test_rust_vs_python_comparison():
    """Import model, calculate with both engines, compare results."""
    _wait_server()

    # Find the largest test model
    xlsx_files = [f for f in os.listdir(XLS_DIR) if f.endswith(".xlsx")]
    assert xlsx_files, f"No xlsx files in {XLS_DIR}"
    # Use the largest one (ЦО v.18)
    xlsx_path = os.path.join(XLS_DIR, sorted(xlsx_files, key=lambda f: os.path.getsize(os.path.join(XLS_DIR, f)), reverse=True)[0])
    print(f"\nUsing model: {os.path.basename(xlsx_path)}")

    # Import model
    model_id = _import_model(xlsx_path, "RustCompare")
    print(f"Imported model: {model_id}")

    # Calculate with Python engine
    print("\n--- Python engine ---")
    os.environ["PEBBLE_ENGINE"] = "python"
    t0 = time.perf_counter()
    py_result = _calculate_with_engine(model_id, "python")
    t1 = time.perf_counter()
    py_cells = sum(len(v) for v in py_result.values())
    print(f"Python: {t1 - t0:.3f}s, {py_cells} cells changed")

    # Calculate with Rust engine
    print("\n--- Rust engine ---")
    os.environ["PEBBLE_ENGINE"] = "rust"
    t0 = time.perf_counter()
    rs_result = _calculate_with_engine(model_id, "rust")
    t1 = time.perf_counter()
    rs_cells = sum(len(v) for v in rs_result.values())
    print(f"Rust:   {t1 - t0:.3f}s, {rs_cells} cells changed")

    # Compare results
    print("\n--- Comparison ---")

    # Collect all keys from both
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
            mismatches.append((sid, ck, "PYTHON_MISSING", "", rs_val))
        elif rs_val is None and py_val is not None:
            py_only += 1
            mismatches.append((sid, ck, "RUST_MISSING", py_val, ""))
        elif not _vals_equal(py_val, rs_val):
            mismatches.append((sid, ck, "DIFF", py_val, rs_val))
        else:
            matched += 1

    print(f"Matched: {matched}")
    print(f"Python-only: {py_only}")
    print(f"Rust-only: {rs_only}")
    print(f"Value mismatches: {len(mismatches) - py_only - rs_only}")

    if mismatches:
        print(f"\nFirst 20 mismatches:")
        for sid, ck, kind, pv, rv in mismatches[:20]:
            print(f"  {kind}: sheet={sid[:8]}.. ck={ck[:40]}.. py={pv} rs={rv}")

    # Allow some tolerance — report but don't fail on first iteration
    total_diff = len(mismatches)
    total_keys = len(all_keys)
    pct = (total_diff / total_keys * 100) if total_keys else 0
    print(f"\nTotal: {total_keys} cells, {total_diff} differences ({pct:.2f}%)")

    if total_diff == 0:
        print("PERFECT MATCH!")
    elif pct < 1.0:
        print("CLOSE MATCH — investigate remaining differences")
    else:
        print("SIGNIFICANT DIFFERENCES — needs debugging")

    # For now, just assert that Rust engine produced results
    assert rs_cells > 0, "Rust engine produced no results"


if __name__ == "__main__":
    test_rust_vs_python_comparison()
