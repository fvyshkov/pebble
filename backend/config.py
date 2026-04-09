import os
from dotenv import load_dotenv

load_dotenv()

DB_PATH = os.environ.get("PEBBLE_DB", "pebble.db")
PORT = int(os.environ.get("PEBBLE_PORT", "8000"))
