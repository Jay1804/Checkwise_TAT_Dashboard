"""Connect to the checkpoint_live MySQL (Aurora/RDS) database.

Credentials are read from environment variables (populated from .env via
python-dotenv) - never hardcoded here.
"""
import os

import pymysql
import pymysql.cursors
from dotenv import load_dotenv

load_dotenv()

REQUIRED_DB_VARS = ["DB_HOST", "DB_PORT", "DB_USER", "DB_PASSWORD", "DB_NAME"]


class ConfigError(RuntimeError):
    """Raised when required .env configuration is missing."""


def _get_db_config():
    missing = [name for name in REQUIRED_DB_VARS if not os.environ.get(name)]
    if missing:
        raise ConfigError(
            f"Missing required database configuration in .env: {', '.join(missing)}"
        )
    return {
        "host": os.environ["DB_HOST"],
        "port": int(os.environ["DB_PORT"]),
        "user": os.environ["DB_USER"],
        "password": os.environ["DB_PASSWORD"],
        "database": os.environ["DB_NAME"],
        "cursorclass": pymysql.cursors.DictCursor,
    }


def get_connection(connect_timeout=15):
    config = _get_db_config()
    return pymysql.connect(connect_timeout=connect_timeout, **config)


def main():
    conn = get_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT VERSION() AS version, DATABASE() AS current_db;")
            print(cursor.fetchone())

            cursor.execute("SHOW TABLES;")
            print("Tables:")
            for row in cursor.fetchall():
                print(" -", list(row.values())[0])
    finally:
        conn.close()


if __name__ == "__main__":
    main()
