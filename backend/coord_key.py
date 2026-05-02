"""coord_key int interning.

cell_data.coord_key historically stored UUID|UUID|UUID (108 bytes/cell on a
3-analytic sheet, ~70% of the table on big models). This module replaces the
UUIDs with auto-allocated INTEGER seq_ids so coord_keys become "12|34|56"
(typically <16 bytes). Save target: -300 MB on the largest model.

Boundary contract:
- DB columns cell_data.coord_key / cell_history.coord_key store seq_id form
- Internal Python code (formula engine, importer, formula generators) keeps
  working with UUIDs and uses pack()/unpack() at the DB boundary
- API still exposes uuid-form coord_keys to clients (frontend builds keys
  from record.id, so the wire format must stay UUID); the server normalizes
  uuid→seq_id on writes and translates seq_id→uuid on reads

Allocation strategy:
- Each analytic_records row carries a UNIQUE seq_id
- Process-local counter primed from MAX(seq_id) on first use; new uuids
  trigger INSERT (UPDATE existing row to set seq_id)
"""
from __future__ import annotations

import asyncio

# uuid -> seq_id (int)
_uuid_to_seq: dict[str, int] = {}
# seq_id (int) -> uuid
_seq_to_uuid: dict[int, str] = {}
# True once we've primed both maps from DB
_loaded = False
_load_lock = asyncio.Lock()


async def _load_all(db) -> None:
    """Prime the cache from analytic_records.seq_id once per process."""
    global _loaded
    if _loaded:
        return
    async with _load_lock:
        if _loaded:
            return
        rows = await db.execute_fetchall(
            "SELECT id, seq_id FROM analytic_records WHERE seq_id IS NOT NULL"
        )
        for r in rows:
            uid = r["id"]
            sid = int(r["seq_id"])
            _uuid_to_seq[uid] = sid
            _seq_to_uuid[sid] = uid
        _loaded = True


def _next_seq() -> int:
    return (max(_seq_to_uuid) if _seq_to_uuid else 0) + 1


async def intern(db, record_uuid: str) -> int:
    """Return the seq_id for *record_uuid*, allocating one if needed."""
    await _load_all(db)
    sid = _uuid_to_seq.get(record_uuid)
    if sid is not None:
        return sid
    # New record (or pre-existing row without seq_id) — allocate and persist
    sid = _next_seq()
    await db.execute(
        "UPDATE analytic_records SET seq_id = ? WHERE id = ?",
        (sid, record_uuid),
    )
    _uuid_to_seq[record_uuid] = sid
    _seq_to_uuid[sid] = record_uuid
    return sid


async def intern_many(db, uuids: list[str]) -> list[int]:
    return [await intern(db, u) for u in uuids]


def expand(seq_id: int | str) -> str:
    """seq_id -> uuid. Raises KeyError on unknown id."""
    return _seq_to_uuid[int(seq_id)]


def expand_safe(seq_id: int | str) -> str | None:
    try:
        return _seq_to_uuid[int(seq_id)]
    except (KeyError, ValueError, TypeError):
        return None


async def pack(db, uuids: list[str]) -> str:
    """Build a coord_key (seq_id form) from uuid parts."""
    return "|".join(str(s) for s in await intern_many(db, uuids))


def pack_sync(uuids: list[str]) -> str:
    """Sync version: requires every uuid already in cache (e.g. inside the
    formula engine after a full load)."""
    return "|".join(str(_uuid_to_seq[u]) for u in uuids)


def unpack(coord_key: str) -> list[str]:
    """coord_key (seq_id form) -> list of uuids. Unknown parts come through
    as the raw string so callers can decide what to do with them (legacy
    UUID-form coord_keys still parse, just become a no-op pass-through)."""
    out: list[str] = []
    for p in coord_key.split("|"):
        out.append(_seq_to_uuid.get(int(p), p) if p.isdigit() else p)
    return out


def to_uuid_coord_key(coord_key: str) -> str:
    """seq_id-form -> uuid-form coord_key (for code paths that still expect
    uuids, e.g. building Rust engine input)."""
    return "|".join(unpack(coord_key))


def from_uuid_coord_key(coord_key: str) -> str:
    """uuid-form -> seq_id-form (sync: cache must be primed)."""
    return "|".join(str(_uuid_to_seq[p]) for p in coord_key.split("|"))


async def from_uuid_coord_key_intern(db, coord_key: str) -> str:
    """uuid-form -> seq_id-form, interning any unknown uuids on the fly.

    Used when consuming engine output: the engine may produce coord_keys for
    auto-aggregate cells whose record uuids exist in analytic_records but
    don't yet carry a seq_id."""
    out: list[str] = []
    for p in coord_key.split("|"):
        sid = _uuid_to_seq.get(p)
        if sid is None:
            sid = await intern(db, p)
        out.append(str(sid))
    return "|".join(out)


async def normalize(db, coord_key: str, *, read_only: bool = True) -> str:
    """Accept either uuid-form or seq_id-form coord_keys and return seq_id form.

    Used at the API boundary so legacy callers that send "{uuid}|{uuid}|..."
    keep working. Default is read_only — unknown UUIDs map to "-1" so the
    lookup simply returns no row. Pass read_only=False on write paths if a
    new UUID can legitimately appear and should be interned (allocates a
    seq_id via UPDATE on analytic_records).
    """
    await _load_all(db)
    out: list[str] = []
    for p in coord_key.split("|"):
        if p.isdigit():
            out.append(p)
        else:
            sid = _uuid_to_seq.get(p)
            if sid is None and not read_only:
                sid = await intern(db, p)
            out.append(str(sid) if sid is not None else "-1")
    return "|".join(out)


def cache_size() -> int:
    return len(_uuid_to_seq)


def _reset_for_tests() -> None:
    """Drop the cache. Tests use a fresh DB per case and need a clean slate."""
    global _loaded
    _uuid_to_seq.clear()
    _seq_to_uuid.clear()
    _loaded = False
