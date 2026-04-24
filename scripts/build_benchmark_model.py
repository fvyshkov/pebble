"""Build a large synthetic benchmark model for engine comparison.

Takes an existing model (ЦО), adds two large analytics:
- "Версии" (Versions): 10 flat records
- "Подразделения" (Departments): ~1000 records, hierarchy depth 3, varied branches

Then fills all manual cells with random values.

Usage: python scripts/build_benchmark_model.py [--api URL] [--model-id ID]
"""
import argparse
import json
import random
import time
import requests

DEFAULT_API = "http://localhost:8001/api"


def create_analytic(api, model_id, name, **kwargs):
    body = {"model_id": model_id, "name": name, **kwargs}
    r = requests.post(f"{api}/analytics", json=body, timeout=30)
    assert r.ok, f"create_analytic failed: {r.status_code} {r.text[:200]}"
    return r.json()


def create_records_bulk(api, analytic_id, records):
    """Create records in bulk. Returns list of dicts with id, sort_order, data_json."""
    r = requests.post(f"{api}/analytics/{analytic_id}/records/bulk", json=records, timeout=60)
    assert r.ok, f"create_records_bulk failed: {r.status_code} {r.text[:200]}"
    result = r.json()
    ids = result.get("created", [])
    # Merge back the input data with IDs
    out = []
    for i, rid in enumerate(ids):
        entry = dict(records[i])
        entry["id"] = rid
        out.append(entry)
    return out


def add_analytic_to_all_sheets(api, model_id, analytic_id):
    r = requests.post(f"{api}/chat/bulk_add_analytic",
                       json={"model_id": model_id, "analytic_id": analytic_id}, timeout=120)
    assert r.ok, f"add_analytic failed: {r.status_code} {r.text[:200]}"
    return r.json()


def get_sheets(api, model_id):
    r = requests.get(f"{api}/sheets/by-model/{model_id}", timeout=30)
    assert r.ok, f"get_sheets failed: {r.status_code}"
    return r.json()


def get_cells(api, sheet_id):
    r = requests.get(f"{api}/cells/by-sheet/{sheet_id}", timeout=60)
    assert r.ok, f"get_cells failed: {r.status_code}"
    return r.json()


def update_cells(api, sheet_id, cells, no_recalc=True):
    r = requests.put(f"{api}/cells/by-sheet/{sheet_id}",
                     json={"cells": cells},
                     params={"no_recalc": "true" if no_recalc else "false"},
                     timeout=120)
    assert r.ok, f"update_cells failed: {r.status_code} {r.text[:200]}"
    return r.json()


def build_versions(api, model_id, n_versions=10):
    """Create 'Версии' analytic: N flat records."""
    print(f"Creating 'Версии' analytic ({n_versions} flat records)...")
    analytic = create_analytic(api, model_id, "Версии", code="versions")
    aid = analytic["id"]

    records = []
    for i in range(n_versions):
        records.append({
            "sort_order": i,
            "data_json": {"name": f"Версия {i+1}"}
        })
    created = create_records_bulk(api, aid, records)
    print(f"  Created analytic {aid}, {len(created)} records")
    return aid


def build_departments(api, model_id, n_top=10, n_sub_min=5, n_sub_max=15, leaf_prob=0.6, n_leaf_min=3, n_leaf_max=10):
    """Create 'Подразделения' analytic with hierarchy."""
    print(f"Creating 'Подразделения' analytic (top={n_top})...")
    analytic = create_analytic(api, model_id, "Подразделения", code="departments")
    aid = analytic["id"]

    # Level 1
    top_records = []
    for i in range(n_top):
        top_records.append({
            "sort_order": i,
            "data_json": {"name": f"Департамент {i+1}"}
        })
    top_created = create_records_bulk(api, aid, top_records)
    total = len(top_created)
    print(f"  Level 1: {total} top departments")

    # Level 2
    for top_rec in top_created:
        parent_id = top_rec["id"]
        n_subs = random.randint(n_sub_min, n_sub_max)
        sub_records = []
        for j in range(n_subs):
            sub_records.append({
                "parent_id": parent_id,
                "sort_order": j,
                "data_json": {"name": f"Отдел {top_rec['sort_order']+1}.{j+1}"}
            })
        sub_created = create_records_bulk(api, aid, sub_records)
        total += len(sub_created)

        # Level 3
        for sub_rec in sub_created:
            if random.random() < leaf_prob:
                n_leaves = random.randint(n_leaf_min, n_leaf_max)
                leaf_records = []
                for k in range(n_leaves):
                    leaf_records.append({
                        "parent_id": sub_rec["id"],
                        "sort_order": k,
                        "data_json": {"name": f"Группа {sub_rec['data_json']['name']}.{k+1}"}
                    })
                leaf_created = create_records_bulk(api, aid, leaf_records)
                total += len(leaf_created)

    print(f"  Total: {total} records across 3 levels")
    return aid


def fill_manual_cells(api, model_id):
    """Fill all manual cells with random values."""
    print("Filling manual cells with random values...")
    sheets = get_sheets(api, model_id)
    total_filled = 0

    for sheet in sheets:
        sid = sheet["id"]
        cells = get_cells(api, sid)
        manual_cells = [c for c in cells if c.get("rule") == "manual" and c.get("value", "") != ""]
        if not manual_cells:
            continue

        # Update manual cells with random values
        updates = []
        for cell in manual_cells:
            updates.append({
                "coord_key": cell["coord_key"],
                "value": str(round(random.uniform(0, 10000), 2)),
                "rule": "manual",
            })

        # Send in batches of 500
        for i in range(0, len(updates), 500):
            batch = updates[i:i+500]
            update_cells(api, sid, batch, no_recalc=True)
            total_filled += len(batch)

        print(f"  Sheet '{sheet['name'][:40]}': {len(updates)} cells filled")

    print(f"  Total: {total_filled} manual cells filled")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--api", default=DEFAULT_API)
    parser.add_argument("--model-id", required=True, help="Existing model to augment")
    parser.add_argument("--skip-fill", action="store_true", help="Skip filling manual cells")
    parser.add_argument("--versions", type=int, default=3, help="Number of version records")
    parser.add_argument("--dept-top", type=int, default=5, help="Number of top-level departments")
    parser.add_argument("--dept-sub-min", type=int, default=3, help="Min subdivisions per dept")
    parser.add_argument("--dept-sub-max", type=int, default=5, help="Max subdivisions per dept")
    args = parser.parse_args()

    api = args.api
    model_id = args.model_id

    print(f"Building benchmark model from {model_id}")
    print(f"API: {api}\n")

    t0 = time.time()

    # 1. Create analytics
    versions_aid = build_versions(api, model_id, n_versions=args.versions)
    departments_aid = build_departments(api, model_id, n_top=args.dept_top,
                                         n_sub_min=args.dept_sub_min,
                                         n_sub_max=args.dept_sub_max)

    # 2. Add analytics to all sheets
    print("\nAdding 'Версии' to all sheets...")
    r1 = add_analytic_to_all_sheets(api, model_id, versions_aid)
    print(f"  Added to {r1['added']}/{r1['total_sheets']} sheets, {r1.get('formulas_suggested', 0)} formulas")

    print("Adding 'Подразделения' to all sheets...")
    r2 = add_analytic_to_all_sheets(api, model_id, departments_aid)
    print(f"  Added to {r2['added']}/{r2['total_sheets']} sheets, {r2.get('formulas_suggested', 0)} formulas")

    # 3. Fill manual cells
    if not args.skip_fill:
        fill_manual_cells(api, model_id)

    elapsed = time.time() - t0
    print(f"\nBenchmark model ready in {elapsed:.1f}s")
    print(f"Model ID: {model_id}")


if __name__ == "__main__":
    main()
