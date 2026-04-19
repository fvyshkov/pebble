"""Tests for backend.llm_cache — disk cache for Anthropic messages.create."""
from __future__ import annotations

import os
import shutil
import tempfile
from types import SimpleNamespace

import pytest

from backend import llm_cache as C


@pytest.fixture(autouse=True)
def _tmp_cache(monkeypatch):
    d = tempfile.mkdtemp(prefix="pebble-llm-cache-")
    monkeypatch.setattr(C, "_DIR", d)
    yield d
    shutil.rmtree(d, ignore_errors=True)


class _FakeClient:
    """Stand-in for anthropic.Anthropic — counts create() calls."""
    def __init__(self, response_dict: dict):
        self.calls = 0
        self._resp = response_dict
        self.messages = self

    def create(self, **kwargs):
        self.calls += 1
        # Return a Message-like SimpleNamespace. The cache will call
        # .model_dump() if available — emulate that attribute.
        resp = SimpleNamespace(**self._resp)
        resp.model_dump = lambda: self._resp
        return resp


_FAKE = {
    "id": "msg_x",
    "type": "message",
    "role": "assistant",
    "model": "claude-sonnet-4-20250514",
    "content": [{"type": "text", "text": "hello"}],
    "stop_reason": "end_turn",
    "stop_sequence": None,
    "usage": {"input_tokens": 1, "output_tokens": 1},
}


def test_cache_stores_on_first_call_and_serves_on_repeat(monkeypatch):
    monkeypatch.setenv("PEBBLE_LLM_MODE", "live")
    client = _FakeClient(_FAKE)

    r1 = C.cached_messages_create(client, model="m1", messages=[{"role": "user", "content": "hi"}])
    r2 = C.cached_messages_create(client, model="m1", messages=[{"role": "user", "content": "hi"}])

    assert client.calls == 1, "Second identical call must be served from cache"
    # Both responses carry the same content text.
    def text_of(r):
        c = r.content
        first = c[0]
        # model-validated Message or SimpleNamespace
        return getattr(first, "text", None) or first["text"]
    assert text_of(r1) == "hello"
    assert text_of(r2) == "hello"


def test_cache_miss_raises_in_cache_mode(monkeypatch):
    monkeypatch.setenv("PEBBLE_LLM_MODE", "cache")
    client = _FakeClient(_FAKE)
    with pytest.raises(C.LLMCacheMiss):
        C.cached_messages_create(client, model="m1", messages=[{"role": "user", "content": "novel"}])
    assert client.calls == 0, "cache mode must not issue a live call on miss"


def test_cache_off_always_calls_live(monkeypatch):
    monkeypatch.setenv("PEBBLE_LLM_MODE", "off")
    client = _FakeClient(_FAKE)
    C.cached_messages_create(client, model="m1", messages=[{"role": "user", "content": "x"}])
    C.cached_messages_create(client, model="m1", messages=[{"role": "user", "content": "x"}])
    assert client.calls == 2, "off mode bypasses cache entirely"


def test_cache_keys_differ_on_different_inputs(monkeypatch):
    monkeypatch.setenv("PEBBLE_LLM_MODE", "live")
    client = _FakeClient(_FAKE)
    C.cached_messages_create(client, model="m1", messages=[{"role": "user", "content": "a"}])
    C.cached_messages_create(client, model="m1", messages=[{"role": "user", "content": "b"}])
    assert client.calls == 2, "distinct inputs must produce distinct cache keys"
