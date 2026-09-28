"""Small persistent account and session store for the public dashboard."""

from __future__ import annotations

import hashlib
import hmac
import secrets
import sqlite3
import time
import uuid
from pathlib import Path

import os


DATA_DIR = Path(os.environ.get("APP_DATA_DIR", Path(__file__).resolve().parent)).resolve()
DATABASE_PATH = DATA_DIR / "accounts.sqlite3"
SESSION_DAYS = 14
PBKDF2_ROUNDS = 310_000


def connect() -> sqlite3.Connection:
    connection = sqlite3.connect(DATABASE_PATH, timeout=20)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def initialize() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with connect() as connection:
        connection.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id TEXT PRIMARY KEY,
                username TEXT NOT NULL COLLATE NOCASE UNIQUE,
                password_salt BLOB NOT NULL,
                password_hash BLOB NOT NULL,
                created_at REAL NOT NULL
            )
        """)
        connection.execute("""
            CREATE TABLE IF NOT EXISTS sessions (
                token_hash TEXT PRIMARY KEY,
                user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                expires_at REAL NOT NULL
            )
        """)
        connection.execute("CREATE INDEX IF NOT EXISTS sessions_expiry ON sessions(expires_at)")


def _password_hash(password: str, salt: bytes) -> bytes:
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ROUNDS, dklen=32)


def create_user(username: str, password: str) -> dict:
    username = username.strip()
    if not 3 <= len(username) <= 32 or not username.isascii() or not all(char.isalnum() or char in "_.-" for char in username):
        raise ValueError("Choose a username with 3–32 letters, numbers, dots, dashes, or underscores.")
    if not 12 <= len(password) <= 256:
        raise ValueError("Choose a password between 12 and 256 characters.")
    user_id = uuid.uuid4().hex
    salt = secrets.token_bytes(16)
    digest = _password_hash(password, salt)
    try:
        with connect() as connection:
            connection.execute(
                "INSERT INTO users (id, username, password_salt, password_hash, created_at) VALUES (?, ?, ?, ?, ?)",
                (user_id, username, salt, digest, time.time()),
            )
    except sqlite3.IntegrityError as error:
        raise ValueError("That username is already in use. Choose another.") from error
    return {"id": user_id, "username": username}


def authenticate(username: str, password: str) -> dict | None:
    if len(username) > 32 or len(password) > 256:
        return None
    with connect() as connection:
        row = connection.execute(
            "SELECT id, username, password_salt, password_hash FROM users WHERE username = ? COLLATE NOCASE",
            (username.strip(),),
        ).fetchone()
    if row is None:
        return None
    candidate = _password_hash(password, bytes(row["password_salt"]))
    if not hmac.compare_digest(candidate, bytes(row["password_hash"])):
        return None
    return {"id": row["id"], "username": row["username"]}


def create_session(user_id: str) -> str:
    token = secrets.token_urlsafe(32)
    token_hash = hashlib.sha256(token.encode("ascii")).hexdigest()
    with connect() as connection:
        now = time.time()
        connection.execute("DELETE FROM sessions WHERE expires_at <= ?", (now,))
        connection.execute(
            "INSERT INTO sessions (token_hash, user_id, expires_at) VALUES (?, ?, ?)",
            (token_hash, user_id, now + SESSION_DAYS * 24 * 60 * 60),
        )
    return token


def get_session(token: str) -> dict | None:
    if not token or len(token) > 128:
        return None
    token_hash = hashlib.sha256(token.encode("utf-8")).hexdigest()
    with connect() as connection:
        row = connection.execute(
            "SELECT users.id, users.username, sessions.expires_at "
            "FROM sessions JOIN users ON users.id = sessions.user_id WHERE sessions.token_hash = ?",
            (token_hash,),
        ).fetchone()
        if row is None:
            return None
        if row["expires_at"] <= time.time():
            connection.execute("DELETE FROM sessions WHERE token_hash = ?", (token_hash,))
            return None
    return {"id": row["id"], "username": row["username"]}


def delete_session(token: str) -> None:
    if not token:
        return
    token_hash = hashlib.sha256(token.encode("utf-8")).hexdigest()
    with connect() as connection:
        connection.execute("DELETE FROM sessions WHERE token_hash = ?", (token_hash,))


def all_user_ids() -> list[str]:
    with connect() as connection:
        return [row["id"] for row in connection.execute("SELECT id FROM users")]


def user_count() -> int:
    with connect() as connection:
        return int(connection.execute("SELECT COUNT(*) FROM users").fetchone()[0])
