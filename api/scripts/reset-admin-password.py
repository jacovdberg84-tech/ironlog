#!/usr/bin/env python3
"""Reset IRONLOG admin password in local SQLite DB(s)."""
from __future__ import annotations

import base64
import hashlib
import os
import sqlite3
import sys
from pathlib import Path

DEFAULT_PASSWORD = "ChangeMe123!"
USERNAMES = ("admin", "Admin")

AUTH_COLUMNS = [
    ("password_hash", "TEXT"),
    ("department", "TEXT"),
    ("allowed_tabs", "TEXT"),
    ("roles_json", "TEXT"),
    ("allowed_locations", "TEXT"),
    ("setup_code_hash", "TEXT"),
    ("setup_code_expires_at", "TEXT"),
    ("pin_hash", "TEXT"),
]


def hash_password(plain: str) -> str:
    salt = os.urandom(16)
    digest = hashlib.scrypt(plain.encode(), salt=salt, n=16384, r=8, p=1, dklen=64)
    return "scrypt$" + base64.b64encode(salt).decode() + "$" + base64.b64encode(digest).decode()


def candidate_db_paths() -> list[Path]:
    paths: list[Path] = []
    api_dir = Path(__file__).resolve().parents[1]
    paths.append(api_dir / "db" / "ironlog.db")

    appdata = os.environ.get("APPDATA", "")
    if appdata:
        paths.append(Path(appdata) / "ironlog-api" / "data" / "db" / "ironlog.db")
        paths.append(Path(appdata) / "ironlog" / "data" / "db" / "ironlog.db")

    local = os.environ.get("LOCALAPPDATA", "")
    if local:
        paths.append(Path(local) / "ironlog-api" / "data" / "db" / "ironlog.db")
        paths.append(Path(local) / "ironlog" / "data" / "db" / "ironlog.db")

    if len(sys.argv) > 2:
        paths.insert(0, Path(sys.argv[2]).expanduser().resolve())

    seen: set[str] = set()
    out: list[Path] = []
    for p in paths:
        key = str(p).lower()
        if key not in seen:
            seen.add(key)
            out.append(p)
    return out


def ensure_auth_columns(conn: sqlite3.Connection) -> None:
    existing = {row[1] for row in conn.execute("PRAGMA table_info(users)").fetchall()}
    for name, col_type in AUTH_COLUMNS:
        if name not in existing:
            conn.execute(f"ALTER TABLE users ADD COLUMN {name} {col_type}")
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS auth_sessions (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          token TEXT NOT NULL UNIQUE,
          user_id INTEGER NOT NULL,
          expires_at TEXT NOT NULL,
          created_at TEXT NOT NULL DEFAULT (datetime('now')),
          FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        )
        """
    )
    conn.execute(
        """
        INSERT INTO users (username, full_name, role, active)
        SELECT 'admin', 'System Admin', 'admin', 1
        WHERE NOT EXISTS (SELECT 1 FROM users WHERE username = 'admin')
        """
    )


def reset_db(path: Path, password: str) -> bool:
    if not path.is_file():
        return False

    conn = sqlite3.connect(str(path))
    try:
        ensure_auth_columns(conn)
        stored = hash_password(password)
        updated = 0
        for username in USERNAMES:
            cur = conn.execute(
                """
                UPDATE users
                SET password_hash = ?, setup_code_hash = NULL, setup_code_expires_at = NULL
                WHERE username = ?
                """,
                (stored, username),
            )
            updated += cur.rowcount
        if updated == 0:
            cur = conn.execute(
                """
                UPDATE users
                SET password_hash = ?, setup_code_hash = NULL, setup_code_expires_at = NULL
                WHERE role = 'admin' AND COALESCE(active, 1) = 1
                """,
                (stored,),
            )
            updated += cur.rowcount
        conn.commit()
        users = conn.execute(
            """
            SELECT username, role,
              CASE WHEN password_hash IS NOT NULL AND length(trim(password_hash)) > 0 THEN 1 ELSE 0 END
            FROM users ORDER BY username
            """
        ).fetchall()
        print(f"OK {path}")
        for row in users:
            print(f"  {row[0]} ({row[1]}) password={'yes' if row[2] else 'no'}")
        return updated > 0
    finally:
        conn.close()


def main() -> int:
    password = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_PASSWORD
    if len(password) < 6:
        print("Password must be at least 6 characters.", file=sys.stderr)
        return 1

    touched = 0
    for path in candidate_db_paths():
        if reset_db(path, password):
            touched += 1

    if not touched:
        print("No database files were updated. Checked:")
        for path in candidate_db_paths():
            print(f"  {path} ({'exists' if path.is_file() else 'missing'})")
        return 1

    print()
    print("Sign in with:")
    print("  Username: admin")
    print(f"  Password: {password}")
    print("Restart IRONLOG / the API if it is already running.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
