"""Attach standard 'Подразделения' (branches_reference.json) + 'Версии'
analytics to a model and recalc.

Usage: python tests/add_branches_versions.py <model_id>
"""
from __future__ import annotations
import json, sys, time
from pathlib import Path
import requests

API = "http://localhost:8000/api"
ROOT = Path(__file__).resolve().parent.parent
BRANCHES_REF = ROOT / "docs" / "branches_reference.json"


def _ok(r: requests.Response, label: str) -> requests.Response:
    if not r.ok:
        raise SystemExit(f"{label}: {r.status_code} {r.text[:200]}")
    return r


def create_branches(model_id: str) -> str:
    aid = _ok(requests.post(f"{API}/analytics",
                            json={"model_id": model_id, "name": "Подразделения"},
                            timeout=30),
              "create branches analytic").json()["id"]
    structure = json.loads(BRANCHES_REF.read_text())["structure"]

    def walk(node: dict, parent_id: str | None) -> None:
        rid = _ok(requests.post(f"{API}/analytics/{aid}/records",
                                json={"data_json": {"name": node["name"]},
                                      "parent_id": parent_id},
                                timeout=30),
                  f"add record {node['name']}").json()["id"]
        for ch in node.get("children", []):
            walk(ch, rid)

    for top in structure:
        walk(top, None)
    print(f"  created Подразделения {aid}")
    return aid


def create_versions(model_id: str) -> str:
    aid = _ok(requests.post(f"{API}/analytics",
                            json={"model_id": model_id, "name": "Версии"},
                            timeout=30),
              "create versions analytic").json()["id"]
    for name in ("Базовый", "Оптимистичный", "Пессимистичный"):
        _ok(requests.post(f"{API}/analytics/{aid}/records",
                          json={"data_json": {"name": name}, "parent_id": None},
                          timeout=30),
            f"add version {name}")
    print(f"  created Версии {aid}")
    return aid


def bind_to_all_sheets(model_id: str, analytic_id: str, label: str) -> None:
    sheets = _ok(requests.get(f"{API}/sheets/by-model/{model_id}", timeout=30), "list sheets").json()
    for s in sheets:
        sas = _ok(requests.get(f"{API}/sheets/{s['id']}/analytics", timeout=30), "list sheet analytics").json()
        next_order = max([sa["sort_order"] for sa in sas], default=-1) + 1
        _ok(requests.post(f"{API}/sheets/{s['id']}/analytics",
                          json={"analytic_id": analytic_id, "sort_order": next_order},
                          timeout=120),
            f"bind {label} to {s['name']}")
    print(f"  bound {label} to {len(sheets)} sheets")


def recalc(model_id: str, rounds: int = 3) -> None:
    sheets = _ok(requests.get(f"{API}/sheets/by-model/{model_id}", timeout=30), "list sheets").json()
    for r in range(1, rounds + 1):
        for s in sheets:
            _ok(requests.post(f"{API}/cells/calculate/{s['id']}", timeout=300),
                f"recalc {s['name']} round{r}")
        print(f"  recalc round {r} done")


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: python tests/add_branches_versions.py <model_id>")
        sys.exit(1)
    model_id = sys.argv[1]
    print(f"Adding analytics to model {model_id}…")
    versions_aid = create_versions(model_id)
    bind_to_all_sheets(model_id, versions_aid, "Версии")
    branches_aid = create_branches(model_id)
    bind_to_all_sheets(model_id, branches_aid, "Подразделения")
    print("Recalc…")
    recalc(model_id)
    print(f"\n→ python3 tests/compare_excel_exact.py XLS-MODELS/MAIN.xlsx {model_id}")


if __name__ == "__main__":
    main()
