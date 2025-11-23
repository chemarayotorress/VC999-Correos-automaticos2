"""Simple Flask backend to support multiusuario history and auth."""
from __future__ import annotations

import hashlib
import json
import os
import secrets
import sqlite3
import threading
import time
from contextlib import closing
from typing import Dict, Tuple

from flask import Flask, jsonify, request

APP = Flask(__name__)
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "backend.db")


def init_db():
    with closing(sqlite3.connect(DB_PATH)) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                license_key TEXT NOT NULL,
                active INTEGER NOT NULL DEFAULT 1
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS tokens (
                token TEXT PRIMARY KEY,
                user_id INTEGER NOT NULL,
                expires_at REAL NOT NULL,
                FOREIGN KEY(user_id) REFERENCES users(id)
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS quotes (
                id TEXT PRIMARY KEY,
                kind TEXT NOT NULL,
                created_at REAL NOT NULL,
                payload TEXT NOT NULL,
                user_id INTEGER,
                FOREIGN KEY(user_id) REFERENCES users(id)
            )
            """
        )
        conn.commit()


def _hash_password(password: str) -> str:
    salt = "vc999"
    return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()


def _ensure_admin():
    with closing(sqlite3.connect(DB_PATH)) as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM users")
        count = cur.fetchone()[0]
        if count == 0:
            cur.execute(
                "INSERT INTO users(username, password_hash, license_key, active) VALUES (?,?,?,1)",
                ("admin@vc999.com", _hash_password("admin"), "DEMO-ADMIN"),
            )
            conn.commit()


def _create_token(user_id: int) -> Tuple[str, float]:
    token = secrets.token_urlsafe(32)
    expires = time.time() + 3600 * 24
    with closing(sqlite3.connect(DB_PATH)) as conn:
        cur = conn.cursor()
        cur.execute("INSERT INTO tokens(token, user_id, expires_at) VALUES (?,?,?)", (token, user_id, expires))
        conn.commit()
    return token, expires


def _auth_user_from_token(token: str):
    with closing(sqlite3.connect(DB_PATH)) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT users.id, users.username, users.license_key, users.active, tokens.expires_at FROM tokens JOIN users ON users.id = tokens.user_id WHERE tokens.token = ?",
            (token,),
        )
        row = cur.fetchone()
        if not row:
            return None
        if row[3] != 1:
            return None
        if row[4] < time.time():
            return None
        return {"id": row[0], "username": row[1], "license_key": row[2], "expires_at": row[4]}


_startup_done = False
_startup_lock = threading.Lock()


def _startup():
    global _startup_done
    if _startup_done:
        return
    with _startup_lock:
        if _startup_done:
            return
        init_db()
        _ensure_admin()
        _startup_done = True


@APP.before_request
def _ensure_startup():
    _startup()


@APP.route("/api/auth/login", methods=["POST"])
def login():
    data = request.get_json(force=True, silent=True) or {}
    username = (data.get("username") or "").strip().lower()
    password = data.get("password") or ""
    license_key = (data.get("license_key") or "").strip()
    if not username or not password:
        return jsonify({"error": "missing_credentials"}), 400
    with closing(sqlite3.connect(DB_PATH)) as conn:
        cur = conn.cursor()
        cur.execute("SELECT id, password_hash, license_key, active FROM users WHERE lower(username)=?", (username,))
        row = cur.fetchone()
        if not row:
            return jsonify({"error": "invalid_user"}), 403
        user_id, pass_hash, license_db, active = row
        if active != 1 or (license_key and license_key != license_db):
            return jsonify({"error": "license_revoked"}), 403
        if pass_hash != _hash_password(password):
            return jsonify({"error": "wrong_password"}), 403
    token, expires = _create_token(user_id)
    return jsonify({"token": token, "expires_at": expires, "username": username, "license_key": license_db})


@APP.route("/api/auth/token", methods=["GET"])
def token_info():
    auth = request.headers.get("Authorization", "")
    if not auth.startswith("Bearer "):
        return jsonify({"error": "missing_token"}), 401
    token = auth.split(" ", 1)[1]
    user = _auth_user_from_token(token)
    if not user:
        return jsonify({"error": "invalid_token"}), 401
    return jsonify(user)


def _require_user():
    auth = request.headers.get("Authorization", "")
    if not auth.startswith("Bearer "):
        return None
    token = auth.split(" ", 1)[1]
    return _auth_user_from_token(token)


@APP.route("/api/quotes", methods=["GET"])
def list_quotes():
    user = _require_user()
    if not user:
        return jsonify({"error": "unauthorized"}), 401
    kind = request.args.get("type", "packaging")
    with closing(sqlite3.connect(DB_PATH)) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT id, payload, created_at FROM quotes WHERE kind=? AND (user_id=? OR user_id IS NULL) ORDER BY created_at DESC",
            (kind, user["id"]),
        )
        items = [
            {"id": row[0], "payload": row[1], "created_at": row[2]}
            for row in cur.fetchall()
        ]
    return jsonify({"items": items})


@APP.route("/api/quotes", methods=["POST"])
def create_quote():
    user = _require_user()
    if not user:
        return jsonify({"error": "unauthorized"}), 401
    data = request.get_json(force=True, silent=True) or {}
    kind = data.get("type", "packaging")
    payload = data.get("quote")
    if not isinstance(payload, dict):
        return jsonify({"error": "invalid_payload"}), 400
    quote_id = payload.get("id") or secrets.token_hex(8)
    with closing(sqlite3.connect(DB_PATH)) as conn:
        cur = conn.cursor()
        cur.execute(
            "REPLACE INTO quotes(id, kind, created_at, payload, user_id) VALUES (?,?,?,?,?)",
            (quote_id, kind, time.time(), json.dumps(payload), user["id"]),
        )
        conn.commit()
    return jsonify({"id": quote_id})


@APP.route("/api/quotes/<quote_id>", methods=["DELETE"])
def delete_quote(quote_id):
    user = _require_user()
    if not user:
        return jsonify({"error": "unauthorized"}), 401
    with closing(sqlite3.connect(DB_PATH)) as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM quotes WHERE id=? AND (user_id=? OR user_id IS NULL)", (quote_id, user["id"]))
        conn.commit()
    return jsonify({"status": "ok"})


@APP.route("/api/metrics", methods=["GET"])
def metrics():
    user = _require_user()
    if not user:
        return jsonify({"error": "unauthorized"}), 401
    with closing(sqlite3.connect(DB_PATH)) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT kind, payload FROM quotes WHERE user_id = ? OR user_id IS NULL",
            (user["id"],),
        )
        data: Dict[str, Dict[str, float]] = {}
        for kind, payload_json in cur.fetchall():
            try:
                payload = json.loads(payload_json)
            except Exception:
                payload = {}
            info = data.setdefault(kind, {"count": 0, "amount": 0.0})
            info["count"] += 1
            total = 0.0
            for key in ("total_numeric", "total_monto", "total"):
                value = payload.get(key)
                if isinstance(value, (int, float)):
                    total = float(value)
                    break
                if isinstance(value, str):
                    try:
                        clean = value.replace("US$", "").replace("$", "").replace(",", "").strip()
                        total = float(clean)
                        break
                    except ValueError:
                        continue
            info["amount"] += total
    return jsonify(data)


def _run_dev_server():
    _startup()
    print("Backend iniciado en http://127.0.0.1:5000 (Ctrl+C para detener)")
    APP.run(host="0.0.0.0", port=5000)


if __name__ == "__main__":
    _run_dev_server()
