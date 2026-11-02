"""Local SQLite database for users and SCIM API keys.

Auth is handled by Okta SSO; this DB only stores who is allowed to access the
portal and what role they have. Passwords are intentionally absent.

The DB file lives next to app.py (gitignored). Schema is created on demand by
init_db(); the first run with BOOTSTRAP_MANAGER_EMAIL set will seed a single
Manager so there is someone to manage everyone else.
"""

import hashlib
import os
import secrets
import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime, timezone

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "users.db")

ROLES = ("Read Only", "Consultant", "Manager", "Project Manager")

# Rendered into the users.role CHECK constraint so the schema stays in sync with ROLES.
_ROLE_CHECK_LIST = ", ".join(f"'{r}'" for r in ROLES)


def _now_iso():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _hash_token(raw):
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


@contextmanager
def get_conn():
    """Connection context manager. Foreign keys + row_factory enabled."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Schema + bootstrap
# ---------------------------------------------------------------------------

SCHEMA = f"""
CREATE TABLE IF NOT EXISTS users (
    id               TEXT PRIMARY KEY,
    email            TEXT UNIQUE NOT NULL COLLATE NOCASE,
    first_name       TEXT NOT NULL,
    last_name        TEXT NOT NULL,
    role             TEXT NOT NULL CHECK (role IN ({_ROLE_CHECK_LIST})),
    active           INTEGER NOT NULL DEFAULT 1,
    scim_external_id TEXT,
    created_at       TEXT NOT NULL,
    updated_at       TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_users_email ON users(email);

CREATE TABLE IF NOT EXISTS api_keys (
    id           TEXT PRIMARY KEY,
    label        TEXT NOT NULL,
    token_hash   TEXT UNIQUE NOT NULL,
    active       INTEGER NOT NULL DEFAULT 1,
    created_by   TEXT,
    created_at   TEXT NOT NULL,
    last_used_at TEXT
);
"""


def _migrate_users_role_check(conn):
    """Rebuild the users table when its role CHECK constraint predates the current ROLES.

    SQLite cannot ALTER a CHECK constraint in place, so we recreate the table with the
    up-to-date constraint and copy the rows over. No-op once every role is present.
    """
    row = conn.execute(
        "SELECT sql FROM sqlite_master WHERE type='table' AND name='users'"
    ).fetchone()
    if not row or all(f"'{r}'" in row["sql"] for r in ROLES):
        return
    print("[db] migrating users.role CHECK constraint to include current ROLES")
    conn.executescript(f"""
        CREATE TABLE users_new (
            id               TEXT PRIMARY KEY,
            email            TEXT UNIQUE NOT NULL COLLATE NOCASE,
            first_name       TEXT NOT NULL,
            last_name        TEXT NOT NULL,
            role             TEXT NOT NULL CHECK (role IN ({_ROLE_CHECK_LIST})),
            active           INTEGER NOT NULL DEFAULT 1,
            scim_external_id TEXT,
            created_at       TEXT NOT NULL,
            updated_at       TEXT NOT NULL
        );
        INSERT INTO users_new
            SELECT id, email, first_name, last_name, role, active,
                   scim_external_id, created_at, updated_at
            FROM users;
        DROP TABLE users;
        ALTER TABLE users_new RENAME TO users;
        CREATE INDEX IF NOT EXISTS idx_users_email ON users(email);
    """)


def init_db():
    """Create tables if needed; seed first Manager from env vars if table empty."""
    with get_conn() as conn:
        conn.executescript(SCHEMA)
        _migrate_users_role_check(conn)
        # Per-user column layout for the manager projects table (JSON). Plain ADD COLUMN — no rebuild.
        cols = {r["name"] for r in conn.execute("PRAGMA table_info(users)")}
        if "column_layout" not in cols:
            conn.execute("ALTER TABLE users ADD COLUMN column_layout TEXT")
        existing = conn.execute("SELECT COUNT(*) AS n FROM users").fetchone()["n"]
    if existing == 0:
        email = os.getenv("BOOTSTRAP_MANAGER_EMAIL")
        if email:
            first = os.getenv("BOOTSTRAP_MANAGER_FIRST_NAME", "Bootstrap")
            last  = os.getenv("BOOTSTRAP_MANAGER_LAST_NAME", "Manager")
            create_user(email=email, first_name=first, last_name=last, role="Manager")
            print(f"[db] seeded bootstrap Manager: {email}")


# ---------------------------------------------------------------------------
# Users CRUD
# ---------------------------------------------------------------------------

def _row_to_user(row):
    if row is None:
        return None
    d = dict(row)
    d["active"] = bool(d["active"])
    d["name"]   = f'{d["first_name"]} {d["last_name"]}'.strip()
    return d


def user_by_email(email):
    if not email:
        return None
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM users WHERE email = ? COLLATE NOCASE",
                           (email,)).fetchone()
    return _row_to_user(row)


def user_by_id(user_id):
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    return _row_to_user(row)


def set_column_layout(user_id, layout_json_str):
    """Persist a user's manager-table column layout (a JSON string, or None to clear)."""
    with get_conn() as conn:
        conn.execute("UPDATE users SET column_layout = ?, updated_at = ? WHERE id = ?",
                     (layout_json_str, _now_iso(), user_id))


def list_users(active_only=False):
    sql = "SELECT * FROM users"
    if active_only:
        sql += " WHERE active = 1"
    sql += " ORDER BY active DESC, last_name COLLATE NOCASE, first_name COLLATE NOCASE"
    with get_conn() as conn:
        rows = conn.execute(sql).fetchall()
    return [_row_to_user(r) for r in rows]


def create_user(*, email, first_name, last_name, role,
                active=True, scim_external_id=None, user_id=None):
    if role not in ROLES:
        raise ValueError(f"invalid role: {role!r}")
    if not email or not first_name or not last_name:
        raise ValueError("email, first_name, last_name are required")
    uid = user_id or str(uuid.uuid4())
    now = _now_iso()
    with get_conn() as conn:
        conn.execute(
            """INSERT INTO users
               (id, email, first_name, last_name, role, active,
                scim_external_id, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (uid, email, first_name, last_name, role, 1 if active else 0,
             scim_external_id, now, now),
        )
    return user_by_id(uid)


def update_user(user_id, **fields):
    allowed = {"email", "first_name", "last_name", "role", "active",
               "scim_external_id"}
    payload = {k: v for k, v in fields.items() if k in allowed}
    if "role" in payload and payload["role"] not in ROLES:
        raise ValueError(f"invalid role: {payload['role']!r}")
    if "active" in payload:
        payload["active"] = 1 if payload["active"] else 0
    if not payload:
        return user_by_id(user_id)
    payload["updated_at"] = _now_iso()
    cols = ", ".join(f"{k} = ?" for k in payload)
    args = list(payload.values()) + [user_id]
    with get_conn() as conn:
        conn.execute(f"UPDATE users SET {cols} WHERE id = ?", args)
    return user_by_id(user_id)


def set_active(user_id, active):
    return update_user(user_id, active=bool(active))


# ---------------------------------------------------------------------------
# API keys
# ---------------------------------------------------------------------------

def _row_to_apikey(row):
    if row is None:
        return None
    d = dict(row)
    d["active"] = bool(d["active"])
    return d


def apikey_create(label, created_by=None):
    """Generate a new API key. Returns (apikey_dict, raw_token).

    Raw token is shown to the user exactly once; only the hash is stored.
    """
    if not label:
        raise ValueError("label is required")
    raw = secrets.token_urlsafe(32)
    token_hash = _hash_token(raw)
    kid = str(uuid.uuid4())
    now = _now_iso()
    with get_conn() as conn:
        conn.execute(
            """INSERT INTO api_keys
               (id, label, token_hash, active, created_by, created_at)
               VALUES (?, ?, ?, 1, ?, ?)""",
            (kid, label, token_hash, created_by, now),
        )
        row = conn.execute("SELECT * FROM api_keys WHERE id = ?", (kid,)).fetchone()
    return _row_to_apikey(row), raw


def apikey_lookup_by_token(raw):
    """Return the api_key row matching the raw token, or None. Touches last_used_at."""
    if not raw:
        return None
    token_hash = _hash_token(raw)
    with get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM api_keys WHERE token_hash = ? AND active = 1",
            (token_hash,),
        ).fetchone()
        if row is None:
            return None
        conn.execute("UPDATE api_keys SET last_used_at = ? WHERE id = ?",
                     (_now_iso(), row["id"]))
    return _row_to_apikey(row)


def apikey_list():
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM api_keys ORDER BY active DESC, created_at DESC"
        ).fetchall()
    return [_row_to_apikey(r) for r in rows]


def apikey_deactivate(api_key_id):
    with get_conn() as conn:
        conn.execute("UPDATE api_keys SET active = 0 WHERE id = ?", (api_key_id,))
