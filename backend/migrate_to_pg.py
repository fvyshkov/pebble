"""Migrate data from SQLite to PostgreSQL.

Usage:
    DATABASE_URL=postgresql://localhost/pebble python -m backend.migrate_to_pg
"""
import asyncio
import sqlite3
import asyncpg
import os
from datetime import datetime

SQLITE_PATH = os.environ.get("PEBBLE_DB", "pebble.db")
PG_DSN = os.environ.get("DATABASE_URL", "postgresql://localhost/pebble")

# Tables in dependency order (parents before children)
TABLES = [
    "models",
    "analytics",
    "analytic_fields",
    "analytic_records",
    "sheets",
    "sheet_analytics",
    "indicator_formula_rules",
    "cell_data",
    "users",
    "sheet_permissions",
    "cell_history",
    "sheet_view_settings",
    "analytic_record_permissions",
    "translations",
    "dag_cache",
    "translation_cache",
    "llm_cache",
    "import_kb",
    "import_kb_log",
]


async def migrate():
    # Connect to both
    sdb = sqlite3.connect(SQLITE_PATH)
    sdb.row_factory = sqlite3.Row
    pool = await asyncpg.create_pool(PG_DSN, min_size=2, max_size=5)

    # First, run the PG schema (init_db would do this, but let's do it explicitly)
    from backend.db import _SCHEMA_PG
    async with pool.acquire() as conn:
        for stmt in _SCHEMA_PG.split(";"):
            stmt = stmt.strip()
            if stmt:
                try:
                    await conn.execute(stmt)
                except Exception as e:
                    pass  # table exists

    for table in TABLES:
        try:
            rows = sdb.execute(f"SELECT * FROM {table}").fetchall()
        except Exception:
            print(f"  {table}: SKIP (not in SQLite)")
            continue

        if not rows:
            print(f"  {table}: 0 rows")
            continue

        cols = rows[0].keys()
        n = len(cols)
        placeholders = ", ".join(f"${i+1}" for i in range(n))
        col_list = ", ".join(cols)
        insert_sql = f"INSERT INTO {table} ({col_list}) VALUES ({placeholders}) ON CONFLICT DO NOTHING"

        # Detect TIMESTAMPTZ columns by checking PG column types
        ts_cols: set[str] = set()
        async with pool.acquire() as conn:
            pg_cols = await conn.fetch(
                "SELECT column_name, data_type FROM information_schema.columns WHERE table_name = $1",
                table,
            )
            for pc in pg_cols:
                if 'timestamp' in pc['data_type']:
                    ts_cols.add(pc['column_name'])

        # Batch insert
        batch = []
        for r in rows:
            vals = []
            for c in cols:
                v = r[c]
                # Convert datetime strings to Python datetime for TIMESTAMPTZ columns
                if c in ts_cols and isinstance(v, str) and v:
                    try:
                        v = datetime.fromisoformat(v.replace('Z', '+00:00'))
                    except ValueError:
                        try:
                            v = datetime.strptime(v, "%Y-%m-%d %H:%M:%S")
                        except ValueError:
                            pass
                vals.append(v)
            batch.append(tuple(vals))

        async with pool.acquire() as conn:
            # Clear existing data in this table (idempotent migration)
            await conn.execute(f"DELETE FROM {table}")
            # Insert in batches of 5000
            for i in range(0, len(batch), 5000):
                chunk = batch[i:i+5000]
                await conn.executemany(insert_sql, chunk)

        print(f"  {table}: {len(rows)} rows migrated")

    await pool.close()
    sdb.close()
    print("\nMigration complete!")


if __name__ == "__main__":
    asyncio.run(migrate())
