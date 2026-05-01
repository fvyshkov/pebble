import os
from dotenv import load_dotenv

load_dotenv()

DB_PATH = os.environ.get("PEBBLE_DB", "pebble.db")
# PostgreSQL DSN. Set to enable Postgres instead of SQLite.
# Example: postgresql://localhost/pebble
DATABASE_URL = os.environ.get("DATABASE_URL", "")
PORT = int(os.environ.get("PEBBLE_PORT", "8000"))
