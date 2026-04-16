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
    fixed_record_id TEXT REFERENCES analytic_records(id) ON DELETE SET NULL
);
CREATE INDEX IF NOT EXISTS idx_sa_sheet ON sheet_analytics(sheet_id);

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
    sheet_id    TEXT PRIMARY KEY REFERENCES sheets(id) ON DELETE CASCADE,
    settings    TEXT NOT NULL DEFAULT '{}'
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
]


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
