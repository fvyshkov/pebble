"""Convert existing cell_data.coord_key from UUID|UUID|UUID form to seq_id form.

Run the new backend at least once first — it backfills analytic_records.seq_id
on startup. This script then translates coord_keys table-side. Idempotent:
already-numeric parts pass through; UUID parts whose record has no seq_id are
left as-is and reported.

Usage:
    python scripts/migrate_coord_key_to_int.py           # dry run
    python scripts/migrate_coord_key_to_int.py --apply   # write
"""
from __future__ import annotations

import argparse
import os
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


def _is_uuid_part(p: str) -> bool:
    return "-" in p and len(p) >= 32


def _convert_one(ck: str, uuid_to_seq: dict[str, int]) -> tuple[str, list[str]]:
    out: list[str] = []
    unknown: list[str] = []
    for p in ck.split("|"):
        if p.isdigit():
            out.append(p)
        elif _is_uuid_part(p):
            sid = uuid_to_seq.get(p)
            if sid is None:
                unknown.append(p)
                out.append(p)
            else:
                out.append(str(sid))
        else:
            out.append(p)
    return "|".join(out), unknown


def migrate_sqlite(db_path: Path, apply: bool) -> None:
    db = sqlite3.connect(str(db_path))
    db.row_factory = sqlite3.Row
    try:
        uuid_to_seq: dict[str, int] = {}
        for r in db.execute(
            "SELECT id, seq_id FROM analytic_records WHERE seq_id IS NOT NULL"
        ):
            uuid_to_seq[r["id"]] = int(r["seq_id"])
        print(f"Loaded {len(uuid_to_seq)} record→seq_id mappings")

        unknown_uuids: set[str] = set()
        for table in ("cell_data", "cell_history"):
            converted = unchanged = 0
            updates: list[tuple[str, str]] = []
            for row in db.execute(f"SELECT id, coord_key FROM {table}"):
                ck = row["coord_key"]
                if not ck:
                    unchanged += 1
                    continue
                new_ck, unknown = _convert_one(ck, uuid_to_seq)
                if unknown:
                    unknown_uuids.update(unknown)
                if new_ck != ck:
                    updates.append((new_ck, row["id"]))
                    converted += 1
                else:
                    unchanged += 1
            print(f"  {table}: {converted} would change, {unchanged} unchanged")
            if apply and updates:
                db.executemany(
                    f"UPDATE {table} SET coord_key = ? WHERE id = ?",
                    updates,
                )
                db.commit()
                print(f"  {table}: wrote {len(updates)} rows")

        if unknown_uuids:
            print(f"\n⚠ {len(unknown_uuids)} UUID parts had no seq_id mapping (left as-is)")
            for u in list(unknown_uuids)[:10]:
                print(f"   {u}")
    finally:
        db.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true",
                    help="actually write changes (default: dry run)")
    ap.add_argument("--db", type=str, default=None,
                    help="path to SQLite DB (default: $PEBBLE_DB or pebble.db)")
    args = ap.parse_args()

    if os.environ.get("DATABASE_URL"):
        print("ERROR: Postgres migration not implemented in this script.")
        print("       Use a psql script that joins cell_data with analytic_records.seq_id.")
        sys.exit(2)

    db_path = Path(args.db or os.environ.get("PEBBLE_DB") or (ROOT / "pebble.db"))
    if not db_path.exists():
        print(f"DB not found: {db_path}")
        sys.exit(2)
    print(f"DB: {db_path}  ({'APPLY' if args.apply else 'dry run'})")
    migrate_sqlite(db_path, apply=args.apply)


if __name__ == "__main__":
    main()
