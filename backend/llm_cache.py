"""Disk-based cache for Anthropic `messages.create` calls.

Why: tests and repeated local runs against the same prompts would otherwise
hit the live API every time — slow, flaky, and costs money. With caching,
identical (model, system, messages, tools) input → same response is replayed
instantly from `.llm_cache/` on disk.

Modes (env `PEBBLE_LLM_MODE`):
    - `live`   (default) — real API call on miss, cache the result.
    - `cache`  — cache-only: miss raises `LLMCacheMiss` so tests can assert
                 no live calls are made. Useful for the "no LLM" test variant
                 once the cache is warm.
    - `off`    — bypass cache entirely (always live call, no store).

Public API:
    cached_messages_create(client, **kwargs) → anthropic.types.Message
"""
from __future__ import annotations

import hashlib
import json
import os
from typing import Any

_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    ".llm_cache",
)


class LLMCacheMiss(RuntimeError):
    """Raised in `cache` mode when a request is not in the cache."""


def _serialize_messages(msgs: Any) -> Any:
    """Anthropic message blocks may be pydantic objects — normalise to dict."""
    if isinstance(msgs, list):
        return [_serialize_messages(m) for m in msgs]
    if hasattr(msgs, "model_dump"):
        return msgs.model_dump()
    if isinstance(msgs, dict):
        return {k: _serialize_messages(v) for k, v in msgs.items()}
    return msgs


def _key_from(kwargs: dict) -> str:
    norm = _serialize_messages(kwargs)
    s = json.dumps(norm, sort_keys=True, ensure_ascii=False, default=str)
    return hashlib.sha256(s.encode()).hexdigest()[:24]


def _mode() -> str:
    return os.environ.get("PEBBLE_LLM_MODE", "live").lower()


def _path_for(key: str) -> str:
    return os.path.join(_DIR, f"msg_{key}.json")


def _load(key: str):
    p = _path_for(key)
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def _store(key: str, value) -> None:
    os.makedirs(_DIR, exist_ok=True)
    with open(_path_for(key), "w", encoding="utf-8") as f:
        json.dump(value, f, ensure_ascii=False, default=str)


def cached_messages_create(client, **kwargs):
    """Drop-in replacement for `client.messages.create(**kwargs)` with disk
    cache. Returns a rehydrated Anthropic Message object (same API as live).
    """
    mode = _mode()
    if mode == "off":
        return client.messages.create(**kwargs)

    key = _key_from(kwargs)
    cached = _load(key)
    if cached is not None:
        # Rehydrate to a Message-like object so callers can keep using
        # `.content[i].text / .type / .name / .input / .id` and `.stop_reason`.
        try:
            import anthropic
            return anthropic.types.Message.model_validate(cached)
        except Exception:
            # Fallback: SimpleNamespace tree (enough for simple callers).
            from types import SimpleNamespace
            def _rehy(v):
                if isinstance(v, dict):
                    return SimpleNamespace(**{k: _rehy(x) for k, x in v.items()})
                if isinstance(v, list):
                    return [_rehy(x) for x in v]
                return v
            return _rehy(cached)

    if mode == "cache":
        raise LLMCacheMiss(
            f"No cached LLM response for key={key} and PEBBLE_LLM_MODE=cache. "
            f"Run once with PEBBLE_LLM_MODE=live to warm the cache."
        )

    # live + miss → call API, cache the result
    resp = client.messages.create(**kwargs)
    try:
        _store(key, resp.model_dump())
    except Exception as e:
        # Never let cache write failure bubble up.
        print(f"[llm_cache] store failed: {e}")
    return resp
