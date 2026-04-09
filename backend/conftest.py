import os
import pytest

# Use in-memory database for tests
os.environ["PEBBLE_DB"] = ":memory:"

from httpx import AsyncClient, ASGITransport
from backend.main import app
from backend.db import init_db, close_db


@pytest.fixture(autouse=True)
async def setup_db():
    await init_db()
    yield
    await close_db()


@pytest.fixture
async def client():
    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as c:
        yield c
