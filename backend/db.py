import aiosqlite
from backend.config import DB_PATH

_db: aiosqlite.Connection | None = None

SCHEMA = """
CREATE TABLE IF NOT EXISTS models (
    id          TEXT PRIMARY KEY,
    name        TEXT NOT NULL DEFAULT '',
    description TEXT NOT NULL DEFAULT '',
    created_at  TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at  TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS analytics (
    id           TEXT PRIMARY KEY,
    model_id     TEXT NOT NULL REFERENCES models(id) ON DELETE CASCADE,
    name         TEXT NOT NULL DEFAULT '',
    code         TEXT NOT NULL DEFAULT '',
    icon         TEXT NOT NULL DEFAULT '',
    is_periods   INTEGER NOT NULL DEFAULT 0,
    data_type    TEXT NOT NULL DEFAULT 'sum',
    period_types TEXT NOT NULL DEFAULT '[]',
    period_start TEXT,
    period_end   TEXT,
    sort_order   INTEGER NOT NULL DEFAULT 0,
    created_at   TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at   TEXT NOT NULL DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS idx_analytics_model ON analytics(model_id);

CREATE TABLE IF NOT EXISTS analytic_fields (
    id           TEXT PRIMARY KEY,
    analytic_id  TEXT NOT NULL REFERENCES analytics(id) ON DELETE CASCADE,
    name         TEXT NOT NULL DEFAULT '',
    code         TEXT NOT NULL DEFAULT '',
    data_type    TEXT NOT NULL DEFAULT 'string',
    sort_order   INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_afields_analytic ON analytic_fields(analytic_id);

CREATE TABLE IF NOT EXISTS analytic_records (
    id           TEXT PRIMARY KEY,
    analytic_id  TEXT NOT NULL REFERENCES analytics(id) ON DELETE CASCADE,
    parent_id    TEXT REFERENCES analytic_records(id) ON DELETE CASCADE,
    sort_order   INTEGER NOT NULL DEFAULT 0,
    data_json    TEXT NOT NULL DEFAULT '{}'
);
CREATE INDEX IF NOT EXISTS idx_arecords_analytic ON analytic_records(analytic_id);
CREATE INDEX IF NOT EXISTS idx_arecords_parent ON analytic_records(parent_id);

CREATE TABLE IF NOT EXISTS sheets (
    id          TEXT PRIMARY KEY,
    model_id    TEXT NOT NULL REFERENCES models(id) ON DELETE CASCADE,
    name        TEXT NOT NULL DEFAULT '',
    created_at  TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at  TEXT NOT NULL DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS idx_sheets_model ON sheets(model_id);

CREATE TABLE IF NOT EXISTS sheet_analytics (
    id              TEXT PRIMARY KEY,
    sheet_id        TEXT NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
    analytic_id     TEXT NOT NULL REFERENCES analytics(id) ON DELETE CASCADE,
    sort_order      INTEGER NOT NULL DEFAULT 0,
    is_fixed        INTEGER NOT NULL DEFAULT 0,
    fixed_record_id TEXT REFERENCES analytic_records(id) ON DELETE SET NULL,
    is_main         INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_sa_sheet ON sheet_analytics(sheet_id);

-- ── Indicator formula rules (per-indicator, per-sheet) ───────
-- kind: 'leaf'          — база для клетки, где все не-главные аналитики листовые
--       'consolidation' — база для клетки, где хотя бы одна не-главная ссылается на не-лист
--       'scoped'        — матчится если все пары scope_json ⊆ coord клетки
-- priority — выше = важнее при мульти-совпадении.
CREATE TABLE IF NOT EXISTS indicator_formula_rules (
    id           TEXT PRIMARY KEY,
    sheet_id     TEXT NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
    indicator_id TEXT NOT NULL REFERENCES analytic_records(id) ON DELETE CASCADE,
    kind         TEXT NOT NULL,
    scope_json   TEXT NOT NULL DEFAULT '{}',
    priority     INTEGER NOT NULL DEFAULT 0,
    formula      TEXT NOT NULL DEFAULT '',
    created_at   TEXT NOT NULL DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS idx_ifr_sheet_indicator
    ON indicator_formula_rules(sheet_id, indicator_id);

CREATE TABLE IF NOT EXISTS cell_data (
    id           TEXT PRIMARY KEY,
    sheet_id     TEXT NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
    coord_key    TEXT NOT NULL,
    value        TEXT,
    data_type    TEXT NOT NULL DEFAULT 'number',
    rule         TEXT NOT NULL DEFAULT 'manual',
    formula      TEXT NOT NULL DEFAULT '',
    UNIQUE(sheet_id, coord_key)
);
CREATE INDEX IF NOT EXISTS idx_cells_sheet ON cell_data(sheet_id);

-- ── Users ───────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS users (
    id          TEXT PRIMARY KEY,
    username    TEXT NOT NULL UNIQUE,
    password    TEXT NOT NULL DEFAULT '',
    created_at  TEXT NOT NULL DEFAULT (datetime('now'))
);

-- ── Sheet Permissions ───────────────────────────────────────
CREATE TABLE IF NOT EXISTS sheet_permissions (
    id          TEXT PRIMARY KEY,
    sheet_id    TEXT NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
    user_id     TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    can_view    INTEGER NOT NULL DEFAULT 1,
    can_edit    INTEGER NOT NULL DEFAULT 1,
    UNIQUE(sheet_id, user_id)
);
CREATE INDEX IF NOT EXISTS idx_sp_sheet ON sheet_permissions(sheet_id);

-- ── Cell History ────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS cell_history (
    id          TEXT PRIMARY KEY,
    sheet_id    TEXT NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
    coord_key   TEXT NOT NULL,
    user_id     TEXT,
    old_value   TEXT,
    new_value   TEXT,
    created_at  TEXT NOT NULL DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS idx_ch_sheet_coord ON cell_history(sheet_id, coord_key);

-- ── Sheet View Settings ─────────────────────────────────────
CREATE TABLE IF NOT EXISTS sheet_view_settings (
    sheet_id    TEXT NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
    user_id     TEXT NOT NULL DEFAULT '',
    settings    TEXT NOT NULL DEFAULT '{}',
    PRIMARY KEY (sheet_id, user_id)
);

-- ── Analytic Record Permissions ─────────────────────────────
CREATE TABLE IF NOT EXISTS analytic_record_permissions (
    id          TEXT PRIMARY KEY,
    user_id     TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    analytic_id TEXT NOT NULL REFERENCES analytics(id) ON DELETE CASCADE,
    record_id   TEXT NOT NULL REFERENCES analytic_records(id) ON DELETE CASCADE,
    can_view    INTEGER DEFAULT 1,
    can_edit    INTEGER DEFAULT 0,
    UNIQUE(user_id, analytic_id, record_id)
);
"""


MIGRATIONS = [
    "ALTER TABLE analytics ADD COLUMN data_type TEXT NOT NULL DEFAULT 'sum'",
    "ALTER TABLE cell_data ADD COLUMN rule TEXT NOT NULL DEFAULT 'manual'",
    "ALTER TABLE cell_data ADD COLUMN formula TEXT NOT NULL DEFAULT ''",
    "ALTER TABLE users ADD COLUMN can_admin INTEGER NOT NULL DEFAULT 0",
    "ALTER TABLE sheets ADD COLUMN excel_code TEXT DEFAULT ''",
    "ALTER TABLE sheets ADD COLUMN sort_order INTEGER DEFAULT 0",
    "ALTER TABLE analytic_records ADD COLUMN excel_row INTEGER",
    "ALTER TABLE sheet_analytics ADD COLUMN is_main INTEGER NOT NULL DEFAULT 0",
    "ALTER TABLE sheet_analytics ADD COLUMN min_period_level TEXT DEFAULT NULL",
]


async def _backfill_is_main(db: aiosqlite.Connection) -> None:
    """Mark one sheet_analytics row per sheet as 'main':
    the lowest sort_order among non-period analytics. If a sheet already
    has any is_main=1, leave it alone.
    """
    # Sheets that need main backfill
    rows = await db.execute_fetchall(
        """SELECT sa.id, sa.sheet_id, sa.sort_order, a.is_periods
           FROM sheet_analytics sa
           JOIN analytics a ON a.id = sa.analytic_id
           WHERE sa.sheet_id NOT IN (
               SELECT DISTINCT sheet_id FROM sheet_analytics WHERE is_main = 1
           )
           ORDER BY sa.sheet_id, sa.sort_order"""
    )
    by_sheet: dict[str, list] = {}
    for r in rows:
        by_sheet.setdefault(r["sheet_id"], []).append(r)
    for sheet_id, lst in by_sheet.items():
        target = next((r for r in lst if not r["is_periods"]), None)
        if target is not None:
            await db.execute(
                "UPDATE sheet_analytics SET is_main = 1 WHERE id = ?",
                (target["id"],),
            )


async def init_db():
    global _db
    _db = await aiosqlite.connect(DB_PATH)
    _db.row_factory = aiosqlite.Row
    await _db.execute("PRAGMA journal_mode=WAL")
    await _db.execute("PRAGMA foreign_keys=ON")
    await _db.executescript(SCHEMA)
    # Run migrations (ignore if column already exists)
    for sql in MIGRATIONS:
        try:
            await _db.execute(sql)
        except Exception:
            pass
    # Migrate sheet_view_settings: add user_id column (composite PK)
    try:
        await _db.execute("ALTER TABLE sheet_view_settings ADD COLUMN user_id TEXT NOT NULL DEFAULT ''")
        # Recreate with composite PK — SQLite can't alter PK, so rebuild
        await _db.executescript("""
            CREATE TABLE IF NOT EXISTS sheet_view_settings_new (
                sheet_id TEXT NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
                user_id  TEXT NOT NULL DEFAULT '',
                settings TEXT NOT NULL DEFAULT '{}',
                PRIMARY KEY (sheet_id, user_id)
            );
            INSERT OR IGNORE INTO sheet_view_settings_new (sheet_id, user_id, settings)
                SELECT sheet_id, COALESCE(user_id, ''), settings FROM sheet_view_settings;
            DROP TABLE sheet_view_settings;
            ALTER TABLE sheet_view_settings_new RENAME TO sheet_view_settings;
        """)
    except Exception:
        pass
    # Ensure default admin user exists
    existing = await _db.execute("SELECT id FROM users WHERE username = 'admin'")
    row = await existing.fetchone()
    if not row:
        import uuid
        admin_id = str(uuid.uuid4())
        await _db.execute(
            "INSERT INTO users (id, username, password, can_admin) VALUES (?, 'admin', 'admin', 1)",
            (admin_id,),
        )
    # Ensure 'admin' user has can_admin flag
    await _db.execute("UPDATE users SET can_admin = 1 WHERE username IN ('Админ', 'admin')")
    # Backfill sheet_analytics.is_main
    await _backfill_is_main(_db)
    await _db.commit()


async def close_db():
    global _db
    if _db:
        await _db.close()
        _db = None


def get_db() -> aiosqlite.Connection:
    if _db is None:
        raise RuntimeError("DB not initialized")
    return _db
