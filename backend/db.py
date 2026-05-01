"""Database abstraction: SQLite (aiosqlite) or PostgreSQL (asyncpg).

Set DATABASE_URL env var to a postgresql:// DSN to use Postgres.
Otherwise falls back to SQLite at PEBBLE_DB (default: pebble.db).
"""
import asyncio
import re
import aiosqlite
from backend.config import DB_PATH, DATABASE_URL

# ── Active backend ──────────────────────────────────────────────────
_backend: str = ""  # "sqlite" or "pg"
_db: aiosqlite.Connection | None = None  # SQLite write connection
_pg_pool = None  # asyncpg pool

# SQLite read pool (WAL parallel reads)
_read_pool: list[aiosqlite.Connection] = []
_read_pool_available: asyncio.Queue | None = None
_READ_POOL_SIZE = 4


# ── SQL dialect helpers ─────────────────────────────────────────────

def _sqlite_to_pg(sql: str) -> str:
    """Convert SQLite SQL to PostgreSQL-compatible SQL."""
    # Track special INSERT modes before placeholder replacement
    has_or_ignore = bool(re.search(r'INSERT\s+OR\s+IGNORE', sql, re.IGNORECASE))
    has_or_replace = bool(re.search(r'INSERT\s+OR\s+REPLACE', sql, re.IGNORECASE))

    # Remove OR IGNORE / OR REPLACE keywords
    sql = re.sub(r'INSERT\s+OR\s+IGNORE\s+INTO', 'INSERT INTO', sql, flags=re.IGNORECASE)
    sql = re.sub(r'INSERT\s+OR\s+REPLACE\s+INTO', 'INSERT INTO', sql, flags=re.IGNORECASE)

    # Replace ? placeholders with $N
    counter = 0
    def _replace_placeholder(m):
        nonlocal counter
        counter += 1
        return f"${counter}"
    sql = re.sub(r'\?', _replace_placeholder, sql)

    # datetime('now') → now()
    sql = sql.replace("datetime('now')", "now()")

    # json_extract(col, '$.field') → col::json->>'field'
    sql = re.sub(
        r"json_extract\((\w+(?:\.\w+)?),\s*'\$\.(\w+)'\)",
        r"\1::json->>'\2'",
        sql,
    )

    # Append ON CONFLICT DO NOTHING for OR IGNORE queries
    if has_or_ignore and 'ON CONFLICT' not in sql:
        sql = sql.rstrip().rstrip(';') + ' ON CONFLICT DO NOTHING'

    return sql


class _DictRow(dict):
    """Dict subclass that supports both dict['key'] and row['key'] access
    (matches aiosqlite.Row behavior)."""
    def __getitem__(self, key):
        return super().__getitem__(key)


class PgConnection:
    """asyncpg pool wrapper that mimics aiosqlite.Connection interface.

    All read queries go through pool.fetch() which runs in parallel.
    Write queries go through pool.execute().
    """

    def __init__(self, pool):
        self._pool = pool

    async def execute_fetchall(self, sql: str, params: tuple = ()) -> list[dict]:
        pg_sql = _sqlite_to_pg(sql)
        rows = await self._pool.fetch(pg_sql, *params)
        return [_DictRow(dict(r)) for r in rows]

    async def execute(self, sql: str, params: tuple = ()):
        pg_sql = _sqlite_to_pg(sql)
        return await self._pool.execute(pg_sql, *params)

    async def executemany(self, sql: str, params_list):
        pg_sql = _sqlite_to_pg(sql)
        async with self._pool.acquire() as conn:
            # asyncpg executemany requires list of tuples
            await conn.executemany(pg_sql, params_list)

    async def executescript(self, sql: str):
        """Execute raw SQL (multiple statements)."""
        async with self._pool.acquire() as conn:
            await conn.execute(sql)

    async def commit(self):
        pass  # asyncpg auto-commits; no-op

    async def close(self):
        await self._pool.close()


# ── Schema (shared, dialect-aware) ──────────────────────────────────

_SCHEMA_SQLITE = """
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
    updated_at   TEXT NOT NULL DEFAULT (datetime('now')),
    color        TEXT DEFAULT NULL
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
    data_json    TEXT NOT NULL DEFAULT '{}',
    excel_row    INTEGER
);
CREATE INDEX IF NOT EXISTS idx_arecords_analytic ON analytic_records(analytic_id);
CREATE INDEX IF NOT EXISTS idx_arecords_parent ON analytic_records(parent_id);

CREATE TABLE IF NOT EXISTS sheets (
    id          TEXT PRIMARY KEY,
    model_id    TEXT NOT NULL REFERENCES models(id) ON DELETE CASCADE,
    name        TEXT NOT NULL DEFAULT '',
    excel_code  TEXT DEFAULT '',
    sort_order  INTEGER DEFAULT 0,
    locked      INTEGER NOT NULL DEFAULT 0,
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
    is_main         INTEGER NOT NULL DEFAULT 0,
    min_period_level TEXT DEFAULT NULL,
    visible_record_ids TEXT DEFAULT NULL
);
CREATE INDEX IF NOT EXISTS idx_sa_sheet ON sheet_analytics(sheet_id);

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

CREATE TABLE IF NOT EXISTS users (
    id          TEXT PRIMARY KEY,
    username    TEXT NOT NULL UNIQUE,
    password    TEXT NOT NULL DEFAULT '',
    can_admin   INTEGER NOT NULL DEFAULT 0,
    created_at  TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS sheet_permissions (
    id          TEXT PRIMARY KEY,
    sheet_id    TEXT NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
    user_id     TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    can_view    INTEGER NOT NULL DEFAULT 1,
    can_edit    INTEGER NOT NULL DEFAULT 1,
    UNIQUE(sheet_id, user_id)
);
CREATE INDEX IF NOT EXISTS idx_sp_sheet ON sheet_permissions(sheet_id);

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

CREATE TABLE IF NOT EXISTS sheet_view_settings (
    sheet_id    TEXT NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
    user_id     TEXT NOT NULL DEFAULT '',
    settings    TEXT NOT NULL DEFAULT '{}',
    PRIMARY KEY (sheet_id, user_id)
);

CREATE TABLE IF NOT EXISTS analytic_record_permissions (
    id          TEXT PRIMARY KEY,
    user_id     TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    analytic_id TEXT NOT NULL REFERENCES analytics(id) ON DELETE CASCADE,
    record_id   TEXT NOT NULL REFERENCES analytic_records(id) ON DELETE CASCADE,
    can_view    INTEGER DEFAULT 1,
    can_edit    INTEGER DEFAULT 0,
    UNIQUE(user_id, analytic_id, record_id)
);

CREATE TABLE IF NOT EXISTS translations (
    id          TEXT PRIMARY KEY,
    entity_type TEXT NOT NULL,
    entity_id   TEXT NOT NULL,
    field       TEXT NOT NULL DEFAULT 'name',
    lang        TEXT NOT NULL,
    value       TEXT NOT NULL DEFAULT '',
    UNIQUE(entity_type, entity_id, field, lang)
);
CREATE INDEX IF NOT EXISTS idx_translations_entity
    ON translations(entity_type, entity_id);

CREATE TABLE IF NOT EXISTS dag_cache (
    model_id   TEXT PRIMARY KEY,
    dag_blob   BLOB NOT NULL,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS translation_cache (
    source_text TEXT NOT NULL,
    lang        TEXT NOT NULL,
    translated  TEXT NOT NULL DEFAULT '',
    PRIMARY KEY (source_text, lang)
);

CREATE TABLE IF NOT EXISTS llm_cache (
    cache_key   TEXT PRIMARY KEY,
    response    TEXT NOT NULL DEFAULT '{}',
    provider    TEXT NOT NULL DEFAULT '',
    created_at  TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS import_kb (
    id          TEXT PRIMARY KEY,
    pattern_type TEXT NOT NULL,
    pattern_key  TEXT NOT NULL UNIQUE,
    match_rule   TEXT NOT NULL DEFAULT '{}',
    action       TEXT NOT NULL DEFAULT '{}',
    confidence   REAL NOT NULL DEFAULT 1.0,
    source       TEXT NOT NULL DEFAULT 'default',
    created_at   TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at   TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS import_kb_log (
    id           TEXT PRIMARY KEY,
    session_id   TEXT NOT NULL,
    sheet_name   TEXT NOT NULL,
    question     TEXT NOT NULL,
    answer       TEXT NOT NULL,
    pattern_id   TEXT,
    created_at   TEXT NOT NULL DEFAULT (datetime('now'))
);
"""

_SCHEMA_PG = """
CREATE TABLE IF NOT EXISTS models (
    id          TEXT PRIMARY KEY,
    name        TEXT NOT NULL DEFAULT '',
    description TEXT NOT NULL DEFAULT '',
    created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS analytics (
    id           TEXT PRIMARY KEY,
    model_id     TEXT NOT NULL REFERENCES models(id) ON DELETE CASCADE,
    name         TEXT NOT NULL DEFAULT '',
    code         TEXT NOT NULL DEFAULT '',
    icon         TEXT NOT NULL DEFAULT '',
    is_periods   INT NOT NULL DEFAULT 0,
    data_type    TEXT NOT NULL DEFAULT 'sum',
    period_types TEXT NOT NULL DEFAULT '[]',
    period_start TEXT,
    period_end   TEXT,
    sort_order   INT NOT NULL DEFAULT 0,
    created_at   TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at   TIMESTAMPTZ NOT NULL DEFAULT now(),
    color        TEXT DEFAULT NULL
);
CREATE INDEX IF NOT EXISTS idx_analytics_model ON analytics(model_id);

CREATE TABLE IF NOT EXISTS analytic_fields (
    id           TEXT PRIMARY KEY,
    analytic_id  TEXT NOT NULL REFERENCES analytics(id) ON DELETE CASCADE,
    name         TEXT NOT NULL DEFAULT '',
    code         TEXT NOT NULL DEFAULT '',
    data_type    TEXT NOT NULL DEFAULT 'string',
    sort_order   INT NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_afields_analytic ON analytic_fields(analytic_id);

CREATE TABLE IF NOT EXISTS analytic_records (
    id           TEXT PRIMARY KEY,
    analytic_id  TEXT NOT NULL REFERENCES analytics(id) ON DELETE CASCADE,
    parent_id    TEXT REFERENCES analytic_records(id) ON DELETE CASCADE,
    sort_order   INT NOT NULL DEFAULT 0,
    data_json    TEXT NOT NULL DEFAULT '{}',
    excel_row    INT
);
CREATE INDEX IF NOT EXISTS idx_arecords_analytic ON analytic_records(analytic_id);
CREATE INDEX IF NOT EXISTS idx_arecords_parent ON analytic_records(parent_id);

CREATE TABLE IF NOT EXISTS sheets (
    id          TEXT PRIMARY KEY,
    model_id    TEXT NOT NULL REFERENCES models(id) ON DELETE CASCADE,
    name        TEXT NOT NULL DEFAULT '',
    excel_code  TEXT DEFAULT '',
    sort_order  INT DEFAULT 0,
    locked      INT NOT NULL DEFAULT 0,
    created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_sheets_model ON sheets(model_id);

CREATE TABLE IF NOT EXISTS sheet_analytics (
    id              TEXT PRIMARY KEY,
    sheet_id        TEXT NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
    analytic_id     TEXT NOT NULL REFERENCES analytics(id) ON DELETE CASCADE,
    sort_order      INT NOT NULL DEFAULT 0,
    is_fixed        INT NOT NULL DEFAULT 0,
    fixed_record_id TEXT REFERENCES analytic_records(id) ON DELETE SET NULL,
    is_main         INT NOT NULL DEFAULT 0,
    min_period_level TEXT DEFAULT NULL,
    visible_record_ids TEXT DEFAULT NULL
);
CREATE INDEX IF NOT EXISTS idx_sa_sheet ON sheet_analytics(sheet_id);

CREATE TABLE IF NOT EXISTS indicator_formula_rules (
    id           TEXT PRIMARY KEY,
    sheet_id     TEXT NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
    indicator_id TEXT NOT NULL REFERENCES analytic_records(id) ON DELETE CASCADE,
    kind         TEXT NOT NULL,
    scope_json   TEXT NOT NULL DEFAULT '{}',
    priority     INT NOT NULL DEFAULT 0,
    formula      TEXT NOT NULL DEFAULT '',
    created_at   TIMESTAMPTZ NOT NULL DEFAULT now()
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

CREATE TABLE IF NOT EXISTS users (
    id          TEXT PRIMARY KEY,
    username    TEXT NOT NULL UNIQUE,
    password    TEXT NOT NULL DEFAULT '',
    can_admin   INT NOT NULL DEFAULT 0,
    created_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS sheet_permissions (
    id          TEXT PRIMARY KEY,
    sheet_id    TEXT NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
    user_id     TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    can_view    INT NOT NULL DEFAULT 1,
    can_edit    INT NOT NULL DEFAULT 1,
    UNIQUE(sheet_id, user_id)
);
CREATE INDEX IF NOT EXISTS idx_sp_sheet ON sheet_permissions(sheet_id);

CREATE TABLE IF NOT EXISTS cell_history (
    id          TEXT PRIMARY KEY,
    sheet_id    TEXT NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
    coord_key   TEXT NOT NULL,
    user_id     TEXT,
    old_value   TEXT,
    new_value   TEXT,
    created_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_ch_sheet_coord ON cell_history(sheet_id, coord_key);

CREATE TABLE IF NOT EXISTS sheet_view_settings (
    sheet_id    TEXT NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
    user_id     TEXT NOT NULL DEFAULT '',
    settings    TEXT NOT NULL DEFAULT '{}',
    PRIMARY KEY (sheet_id, user_id)
);

CREATE TABLE IF NOT EXISTS analytic_record_permissions (
    id          TEXT PRIMARY KEY,
    user_id     TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    analytic_id TEXT NOT NULL REFERENCES analytics(id) ON DELETE CASCADE,
    record_id   TEXT NOT NULL REFERENCES analytic_records(id) ON DELETE CASCADE,
    can_view    INT DEFAULT 1,
    can_edit    INT DEFAULT 0,
    UNIQUE(user_id, analytic_id, record_id)
);

CREATE TABLE IF NOT EXISTS translations (
    id          TEXT PRIMARY KEY,
    entity_type TEXT NOT NULL,
    entity_id   TEXT NOT NULL,
    field       TEXT NOT NULL DEFAULT 'name',
    lang        TEXT NOT NULL,
    value       TEXT NOT NULL DEFAULT '',
    UNIQUE(entity_type, entity_id, field, lang)
);
CREATE INDEX IF NOT EXISTS idx_translations_entity
    ON translations(entity_type, entity_id);

CREATE TABLE IF NOT EXISTS dag_cache (
    model_id   TEXT PRIMARY KEY,
    dag_blob   BYTEA NOT NULL,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS translation_cache (
    source_text TEXT NOT NULL,
    lang        TEXT NOT NULL,
    translated  TEXT NOT NULL DEFAULT '',
    PRIMARY KEY (source_text, lang)
);

CREATE TABLE IF NOT EXISTS llm_cache (
    cache_key   TEXT PRIMARY KEY,
    response    TEXT NOT NULL DEFAULT '{}',
    provider    TEXT NOT NULL DEFAULT '',
    created_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS import_kb (
    id          TEXT PRIMARY KEY,
    pattern_type TEXT NOT NULL,
    pattern_key  TEXT NOT NULL UNIQUE,
    match_rule   TEXT NOT NULL DEFAULT '{}',
    action       TEXT NOT NULL DEFAULT '{}',
    confidence   REAL NOT NULL DEFAULT 1.0,
    source       TEXT NOT NULL DEFAULT 'default',
    created_at   TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at   TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS import_kb_log (
    id           TEXT PRIMARY KEY,
    session_id   TEXT NOT NULL,
    sheet_name   TEXT NOT NULL,
    question     TEXT NOT NULL,
    answer       TEXT NOT NULL,
    pattern_id   TEXT,
    created_at   TIMESTAMPTZ NOT NULL DEFAULT now()
);
"""


# ── Backfill helper ────────────────────────────────────────────────

# Process-local cache of formula text → id. Populated lazily.
_formula_id_cache: dict[str, int] = {}


async def intern_formula(db, text: str | None) -> int | None:
    """Return the formula_id for *text*, inserting a new formulas row if needed.
    Returns None for empty/null formulas. Cached in-process — safe across
    transactions because formulas table is append-only (UNIQUE on text)."""
    if not text:
        return None
    fid = _formula_id_cache.get(text)
    if fid is not None:
        return fid
    rows = await db.execute_fetchall("SELECT id FROM formulas WHERE text = ?", (text,))
    if rows:
        fid = rows[0]["id"]
    else:
        cur = await db.execute("INSERT OR IGNORE INTO formulas (text) VALUES (?)", (text,))
        fid = getattr(cur, "lastrowid", None)
        if fid is None or fid == 0:
            rows = await db.execute_fetchall("SELECT id FROM formulas WHERE text = ?", (text,))
            fid = rows[0]["id"] if rows else None
    if fid is not None:
        _formula_id_cache[text] = fid
    return fid


async def _backfill_is_main(db) -> None:
    """Mark one sheet_analytics row per sheet as 'main'."""
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


# ── Init / Close ────────────────────────────────────────────────────

async def init_db():
    global _db, _pg_pool, _backend, _read_pool_available

    if DATABASE_URL:
        await _init_pg()
    else:
        await _init_sqlite()


async def _init_pg():
    global _pg_pool, _db, _backend
    import asyncpg
    _backend = "pg"
    _pg_pool = await asyncpg.create_pool(DATABASE_URL, min_size=4, max_size=20)
    _db = PgConnection(_pg_pool)

    # Create schema — execute each statement separately (Postgres doesn't support
    # multi-statement CREATE TABLE IF NOT EXISTS in one call the same way).
    async with _pg_pool.acquire() as conn:
        for stmt in _SCHEMA_PG.split(";"):
            stmt = stmt.strip()
            if stmt:
                try:
                    await conn.execute(stmt)
                except Exception:
                    pass  # table already exists, etc.

    # Ensure admin user
    rows = await _db.execute_fetchall("SELECT id FROM users WHERE username = ?", ("admin",))
    if not rows:
        import uuid
        await _db.execute(
            "INSERT INTO users (id, username, password, can_admin) VALUES (?, 'admin', 'admin', 1)",
            (str(uuid.uuid4()),),
        )
    await _db.execute("UPDATE users SET can_admin = 1 WHERE username IN (?, ?)", ("Админ", "admin"))
    await _backfill_is_main(_db)
    print(f"[db] PostgreSQL connected: {DATABASE_URL}")


async def _init_sqlite():
    global _db, _backend, _read_pool_available
    _backend = "sqlite"
    _db = await aiosqlite.connect(DB_PATH)
    _db.row_factory = aiosqlite.Row
    await _db.execute("PRAGMA journal_mode=WAL")
    await _db.execute("PRAGMA foreign_keys=ON")
    await _db.executescript(_SCHEMA_SQLITE)
    # Run migrations (ignore if column already exists)
    MIGRATIONS = [
        "ALTER TABLE analytics ADD COLUMN data_type TEXT NOT NULL DEFAULT 'sum'",
        "ALTER TABLE cell_data ADD COLUMN rule TEXT NOT NULL DEFAULT 'manual'",
        "ALTER TABLE cell_data ADD COLUMN formula TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE users ADD COLUMN can_admin INTEGER NOT NULL DEFAULT 0",
        "ALTER TABLE sheets ADD COLUMN excel_code TEXT DEFAULT ''",
        "ALTER TABLE sheets ADD COLUMN sort_order INTEGER DEFAULT 0",
        "ALTER TABLE analytic_records ADD COLUMN excel_row INTEGER",
        "ALTER TABLE sheets ADD COLUMN locked INTEGER NOT NULL DEFAULT 0",
        "ALTER TABLE sheet_analytics ADD COLUMN is_main INTEGER NOT NULL DEFAULT 0",
        "ALTER TABLE sheet_analytics ADD COLUMN min_period_level TEXT DEFAULT NULL",
        "ALTER TABLE sheet_analytics ADD COLUMN visible_record_ids TEXT DEFAULT NULL",
        """CREATE TABLE IF NOT EXISTS translation_cache (
            source_text TEXT NOT NULL,
            lang        TEXT NOT NULL,
            translated  TEXT NOT NULL DEFAULT '',
            PRIMARY KEY (source_text, lang)
        )""",
        """CREATE TABLE IF NOT EXISTS llm_cache (
            cache_key   TEXT PRIMARY KEY,
            response    TEXT NOT NULL DEFAULT '{}',
            provider    TEXT NOT NULL DEFAULT '',
            created_at  TEXT NOT NULL DEFAULT (datetime('now'))
        )""",
        "ALTER TABLE analytics ADD COLUMN color TEXT DEFAULT NULL",
        "ALTER TABLE models ADD COLUMN calc_status TEXT NOT NULL DEFAULT 'ready'",
        """CREATE TABLE IF NOT EXISTS import_kb (
            id          TEXT PRIMARY KEY,
            pattern_type TEXT NOT NULL,
            pattern_key  TEXT NOT NULL UNIQUE,
            match_rule   TEXT NOT NULL DEFAULT '{}',
            action       TEXT NOT NULL DEFAULT '{}',
            confidence   REAL NOT NULL DEFAULT 1.0,
            source       TEXT NOT NULL DEFAULT 'default',
            created_at   TEXT NOT NULL DEFAULT (datetime('now')),
            updated_at   TEXT NOT NULL DEFAULT (datetime('now'))
        )""",
        """CREATE TABLE IF NOT EXISTS import_kb_log (
            id           TEXT PRIMARY KEY,
            session_id   TEXT NOT NULL,
            sheet_name   TEXT NOT NULL,
            question     TEXT NOT NULL,
            answer       TEXT NOT NULL,
            pattern_id   TEXT,
            created_at   TEXT NOT NULL DEFAULT (datetime('now'))
        )""",
        # Formula interning: 99.7% of formula rows reuse one of ~2000 formulas.
        # Storing them by integer id slashes cell_data from ~480 MB to ~410 MB.
        """CREATE TABLE IF NOT EXISTS formulas (
            id   INTEGER PRIMARY KEY AUTOINCREMENT,
            text TEXT NOT NULL UNIQUE
        )""",
        "ALTER TABLE cell_data ADD COLUMN formula_id INTEGER",
        "CREATE INDEX IF NOT EXISTS idx_cells_formula_id ON cell_data(formula_id) WHERE formula_id IS NOT NULL",
    ]
    for sql in MIGRATIONS:
        try:
            await _db.execute(sql)
        except Exception:
            pass
    # Migrate sheet_view_settings
    try:
        await _db.execute("ALTER TABLE sheet_view_settings ADD COLUMN user_id TEXT NOT NULL DEFAULT ''")
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
    # Ensure admin user
    existing = await _db.execute("SELECT id FROM users WHERE username = 'admin'")
    row = await existing.fetchone()
    if not row:
        import uuid
        await _db.execute(
            "INSERT INTO users (id, username, password, can_admin) VALUES (?, 'admin', 'admin', 1)",
            (str(uuid.uuid4()),),
        )
    await _db.execute("UPDATE users SET can_admin = 1 WHERE username IN ('Админ', 'admin')")
    await _backfill_is_main(_db)
    await _db.commit()

    # Read-only connection pool for parallel reads
    _read_pool_available = asyncio.Queue(maxsize=_READ_POOL_SIZE)
    for _ in range(_READ_POOL_SIZE):
        rconn = await aiosqlite.connect(f"file:{DB_PATH}?mode=ro", uri=True)
        rconn.row_factory = aiosqlite.Row
        await rconn.execute("PRAGMA journal_mode=WAL")
        _read_pool.append(rconn)
        await _read_pool_available.put(rconn)
    print(f"[db] SQLite connected: {DB_PATH}")


async def close_db():
    global _db, _read_pool_available, _pg_pool
    if _backend == "pg" and _pg_pool:
        await _pg_pool.close()
        _pg_pool = None
    elif _db:
        await _db.close()
    _db = None
    for conn in _read_pool:
        await conn.close()
    _read_pool.clear()
    _read_pool_available = None


def get_db():
    if _db is None:
        raise RuntimeError("DB not initialized")
    return _db


class ReadConn:
    """Async context manager: borrows a read-only connection from the pool.
    For Postgres, just returns the main PgConnection (pool handles parallelism).
    For SQLite, borrows from the read-only pool."""

    __slots__ = ("_conn",)

    def __init__(self):
        self._conn = None

    async def __aenter__(self):
        if _backend == "pg":
            return get_db()  # asyncpg pool handles parallel reads natively
        if _read_pool_available is None:
            return get_db()
        self._conn = await _read_pool_available.get()
        return self._conn

    async def __aexit__(self, *exc):
        if self._conn is not None and _read_pool_available is not None:
            await _read_pool_available.put(self._conn)
            self._conn = None


def get_read_db() -> ReadConn:
    """Get a read-only connection (use as async context manager)."""
    return ReadConn()


def is_postgres() -> bool:
    return _backend == "pg"
