"""Re-import MAIN.xlsx, recalc, run accuracy comparison.

Steps:
  1. Delete existing MAIN model (via API)
  2. Import MAIN.xlsx via streaming
  3. Run 3 recalc rounds
  4. Trigger accuracy comparison via tests/compare_excel_exact.py
"""
from __future__ import annotations
import json, os, sqlite3, sys, time, requests
from pathlib import Path

API = os.environ.get("PEBBLE_API", "http://localhost:8000/api")
EXCEL_PATH = Path(__file__).resolve().parent.parent / "XLS-MODELS" / "MAIN.xlsx"
DB_PATH = Path(__file__).resolve().parent.parent / "pebble.db"


def delete_existing():
    db = sqlite3.connect(str(DB_PATH))
    rows = db.execute("SELECT id FROM models WHERE name='MAIN'").fetchall()
    db.close()
    for (mid,) in rows:
        resp = requests.delete(f"{API}/models/{mid}")
        print(f"Delete {mid}: {resp.status_code}")


def import_model():
    print(f"Importing {EXCEL_PATH.name}…")
    with open(EXCEL_PATH, "rb") as f:
        resp = requests.post(
            f"{API}/import/excel-stream",
            files={"file": (EXCEL_PATH.name, f)},
            stream=True,
            timeout=600,
        )
    model_id = None
    for line in resp.iter_lines(decode_unicode=True):
        if not line or not line.startswith("data: "):
            continue
        data = json.loads(line[6:])
        if data.get("type") == "progress":
            step = data.get("step", "")
            detail = data.get("detail", "")
            pct = data.get("percent", "")
            print(f"  [{step}] {detail} {pct}")
        if data.get("done"):
            model_id = data.get("model_id")
            print(f"Import done. model_id={model_id}")
            break
        if data.get("type") == "error":
            print(f"  ERROR: {data}")
    return model_id


def recalc(model_id, rounds=3):
    db = sqlite3.connect(str(DB_PATH))
    sheet_ids = [r[0] for r in db.execute(
        "SELECT id FROM sheets WHERE model_id=? ORDER BY sort_order", (model_id,)
    ).fetchall()]
    db.close()
    print(f"Recalc: {len(sheet_ids)} sheets × {rounds} rounds")
    for r in range(1, rounds + 1):
        for sid in sheet_ids:
            resp = requests.post(f"{API}/cells/calculate/{sid}", timeout=300)
            if resp.status_code != 200:
                print(f"  round{r} sheet {sid}: {resp.status_code} {resp.text[:200]}")
        print(f"  round{r} done")


def main():
    delete_existing()
    mid = import_model()
    if not mid:
        print("Import failed")
        sys.exit(1)
    recalc(mid, rounds=3)
    print(f"\n→ python3 tests/compare_excel_exact.py {EXCEL_PATH} {mid}")


if __name__ == "__main__":
    main()
