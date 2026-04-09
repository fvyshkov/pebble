import os

DB_PATH = os.environ.get("PEBBLE_DB", "pebble.db")
PORT = int(os.environ.get("PEBBLE_PORT", "8000"))
