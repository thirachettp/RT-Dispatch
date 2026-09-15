"""
SCDC - Vehicle Handling Task Management (MVP v3)
Single-file backend: FastAPI + SQLite.

Run:
    pip install -r requirements.txt
    python -m uvicorn main:app --reload
Then open http://localhost:8000 in a browser.
"""

import base64
import csv
import hashlib
import hmac
import io
import json
import logging
import os
import re
import secrets
import sqlite3
import time
import uuid
from contextlib import contextmanager
from datetime import datetime, timedelta
from typing import Optional, List

from fastapi import FastAPI, Depends, HTTPException, Header, UploadFile, File
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse, Response
from pydantic import BaseModel, Field

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from py_vapid import Vapid
    from pywebpush import webpush, WebPushException
    from cryptography.hazmat.primitives import serialization
    WEB_PUSH_AVAILABLE = True
except ImportError:
    WEB_PUSH_AVAILABLE = False

# Postgres is optional: set DATABASE_URL (e.g. on Vercel, where there is no
# persistent local disk for a SQLite file) to switch the whole app over to it.
# Leave it unset for local dev / any platform with a normal persistent
# filesystem (Replit, Render, Railway, etc.) and SQLite is used exactly as
# before — nothing else needs to change to run locally.
DATABASE_URL = os.environ.get("DATABASE_URL")
IS_POSTGRES = bool(DATABASE_URL)
_IMPORT_ERROR = None
if IS_POSTGRES:
    try:
        import psycopg2
        import psycopg2.extras
    except Exception as e:
        # If the Postgres driver can't even be imported (a known failure mode
        # for psycopg2-binary on some serverless runtimes), don't let it kill
        # the whole module at load time — that produces an opaque
        # FUNCTION_INVOCATION_FAILED on EVERY route, including /healthz.
        # Record it and let the app finish importing so /healthz can report it.
        _IMPORT_ERROR = f"{type(e).__name__}: {e}"
        IS_POSTGRES = False

log = logging.getLogger("uvicorn.error")

BASE_DIR = os.path.dirname(__file__)
DB_PATH = os.path.join(BASE_DIR, "scdc.db")
FRONTEND_DIR = os.path.join(BASE_DIR, "..", "frontend")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
if not IS_POSTGRES:
    os.makedirs(UPLOAD_DIR, exist_ok=True)

WAITING_TOO_LONG_MINUTES = 15
IN_PROGRESS_OVERDUE_MINUTES = 60
PENDING_PASSWORD = "PENDING"
MIN_PASSWORD_LENGTH = 6
VAPID_CLAIM_SUB = "mailto:admin@scdc.local"

app = FastAPI(title="SCDC Vehicle Handling Task Management (MVP)")


# ---------------------------------------------------------------------------
# Web Push (VAPID key pair generated once, stored in the database — not a
# file, see app_config table above for why)
# ---------------------------------------------------------------------------

def _load_vapid_from_db(db):
    if not WEB_PUSH_AVAILABLE:
        return None
    row = db.execute("SELECT value FROM app_config WHERE key='vapid_private_pem'").fetchone()
    if not row or not row["value"]:
        return None
    return Vapid.from_pem(row["value"].encode())


def ensure_vapid_keys():
    if not WEB_PUSH_AVAILABLE:
        return
    with get_db() as db:
        if db.execute("SELECT 1 FROM app_config WHERE key='vapid_private_pem'").fetchone():
            return
        v = Vapid()
        v.generate_keys()
        pem = v.private_pem().decode()
        db.execute(
            "INSERT INTO app_config (key, value) VALUES ('vapid_private_pem', ?) "
            "ON CONFLICT(key) DO NOTHING",
            (pem,),
        )
        log.info("Generated a new VAPID key pair and stored it in the database.")


def get_vapid_public_key_b64() -> Optional[str]:
    with get_db() as db:
        v = _load_vapid_from_db(db)
    if not v:
        return None
    raw = v.public_key.public_bytes(serialization.Encoding.X962, serialization.PublicFormat.UncompressedPoint)
    return base64.urlsafe_b64encode(raw).rstrip(b"=").decode()


def send_web_push_to_user(db, user_id: int, title: str, body: str, url: str = "/"):
    """Best-effort: never let a push failure break the request that triggered it
    (existing callers via notify() ignore the return value). Also returns a
    summary dict — {"attempted": N, "sent": M, "reason": str|None} — so the
    /push/test endpoint can tell the person in the UI, right when they enable
    notifications, whether delivery actually worked instead of just hoping."""
    if not WEB_PUSH_AVAILABLE:
        log.info("Web push skipped for user %s: pywebpush not installed", user_id)
        return {"attempted": 0, "sent": 0, "reason": "เซิร์ฟเวอร์ยังไม่ได้ติดตั้งไลบรารี Web Push"}
    vapid = _load_vapid_from_db(db)
    if not vapid:
        log.info("Web push skipped for user %s: no VAPID key in the database yet", user_id)
        return {"attempted": 0, "sent": 0, "reason": "เซิร์ฟเวอร์ยังไม่ได้ตั้งค่า VAPID key"}
    subs = db.execute("SELECT * FROM push_subscriptions WHERE user_id=?", (user_id,)).fetchall()
    if not subs:
        log.info("Web push skipped for user %s: no push subscription on file "
                 "(they haven't turned on notifications in Profile, or it never completed)", user_id)
        return {"attempted": 0, "sent": 0, "reason": "ยังไม่มีการเปิดใช้งานแจ้งเตือนไว้เลย"}
    payload = json.dumps({"title": title, "body": body, "url": url})
    sent = 0
    for s in subs:
        try:
            webpush(
                subscription_info={"endpoint": s["endpoint"], "keys": {"p256dh": s["p256dh"], "auth": s["auth"]}},
                data=payload,
                vapid_private_key=vapid,
                vapid_claims={"sub": VAPID_CLAIM_SUB},
                ttl=60,
            )
            sent += 1
            log.info("Web push sent to user %s (subscription %s)", user_id, s["id"])
        except WebPushException as e:
            status = getattr(getattr(e, "response", None), "status_code", None)
            if status in (404, 410, 403):
                # 404/410 = subscription gone. 403 almost always means the
                # push service rejected our VAPID signature outright (e.g.
                # "BadJwtToken") — this happens if the VAPID key ever changed
                # after this subscription was created, making it permanently
                # unusable. None of these are worth retrying, so clean them up
                # automatically instead of failing silently forever; the user
                # just needs to re-enable notifications once in Profile to
                # get a fresh subscription.
                db.execute("DELETE FROM push_subscriptions WHERE id=?", (s["id"],))
                log.info("Web push subscription %s for user %s is invalid (status %s) — removed it. "
                         "They'll need to toggle notifications off/on again in Profile.", s["id"], user_id, status)
            else:
                log.warning("Web push failed for subscription %s (user %s): %s", s["id"], user_id, e)
        except Exception as e:
            # Most commonly: the server has no outbound internet access to reach
            # the browser's push relay (fcm.googleapis.com / Mozilla autopush /
            # Apple push) — very possible on a LAN-only office deployment.
            log.warning("Web push error for subscription %s (user %s): %s", s["id"], user_id, e)
    return {
        "attempted": len(subs), "sent": sent,
        "reason": None if sent else "ส่งไม่สำเร็จทุกอุปกรณ์ที่เปิดแจ้งเตือนไว้ — ลองปิดแล้วเปิดใหม่อีกครั้ง",
    }


# ---------------------------------------------------------------------------
# DB helpers
# ---------------------------------------------------------------------------

if IS_POSTGRES:
    IntegrityError = psycopg2.IntegrityError

    def _convert_placeholders(query: str) -> str:
        """Convert SQLite's `?` placeholders to Postgres's `%s`, but only
        outside single-quoted string literals. A naive query.replace("?",
        "%s") also mangles any literal `?` character stored as DATA — e.g.
        this app's zone code_pattern values are regexes that legitimately
        contain `?` as an "optional group" marker (`([A-G][12])?`), which a
        blind replace corrupts and then makes psycopg2 choke on (it sees a
        phantom placeholder with nothing supplied to fill it). This tracks
        whether we're inside a '...' string as it scans, and leaves `?`
        alone whenever it's inside one."""
        result = []
        in_string = False
        i, n = 0, len(query)
        while i < n:
            ch = query[i]
            if ch == "'":
                if in_string and i + 1 < n and query[i + 1] == "'":
                    result.append("''")  # escaped '' inside a string literal
                    i += 2
                    continue
                in_string = not in_string
                result.append(ch)
            elif ch == "?" and not in_string:
                result.append("%s")
            else:
                result.append(ch)
            i += 1
        return "".join(result)

    class _PgCursor:
        """Wraps a psycopg2 cursor (backed by RealDictCursor, so rows already
        behave like dicts / support row["col"] the same way sqlite3.Row does)
        so call sites don't need to know which database is actually running."""
        def __init__(self, cur):
            self._cur = cur

        def fetchone(self):
            return self._cur.fetchone()

        def fetchall(self):
            return self._cur.fetchall()

        @property
        def rowcount(self):
            return self._cur.rowcount

    class _PgConnection:
        """Wraps a psycopg2 connection so the rest of the app can keep calling
        db.execute(sql_with_question_marks, params) / db.executescript(sql)
        exactly as it does today for sqlite3 — only this wrapper needs to know
        Postgres uses %s placeholders instead of ?."""
        def __init__(self, conn):
            self._conn = conn

        def execute(self, query, params=()):
            cur = self._conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            cur.execute(_convert_placeholders(query), params)
            return _PgCursor(cur)

        def executescript(self, script):
            # psycopg2 can run a multi-statement string in one call as long as
            # it contains no parameters, which is exactly what our DDL scripts are.
            cur = self._conn.cursor()
            cur.execute(script)

        def commit(self):
            self._conn.commit()

        def close(self):
            self._conn.close()
else:
    IntegrityError = sqlite3.IntegrityError


@contextmanager
def get_db():
    if IS_POSTGRES:
        conn = psycopg2.connect(DATABASE_URL)
        wrapped = _PgConnection(conn)
        try:
            yield wrapped
            conn.commit()
        finally:
            conn.close()
    else:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        try:
            yield conn
            conn.commit()
        finally:
            conn.close()


def hash_password(password: str, salt: Optional[str] = None) -> str:
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 100_000)
    return f"{salt}${digest.hex()}"


def verify_password(password: str, stored: str) -> bool:
    salt, _, digest_hex = stored.partition("$")
    if not salt or not digest_hex:
        return False
    check = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 100_000).hex()
    return hmac.compare_digest(check, digest_hex)


def now_iso() -> str:
    return datetime.utcnow().isoformat(timespec="seconds") + "Z"


def add_minutes_iso(minutes: int) -> str:
    return (datetime.utcnow() + timedelta(minutes=minutes)).isoformat(timespec="seconds") + "Z"


def generate_task_code(db) -> str:
    """TK{YYMMDD}-{seq}, e.g. TK260910-001. The date uses Thailand local time
    (UTC+7) — everything else in this app is stored in UTC, but a task code
    that flips to "tomorrow" at 7am local time (UTC midnight) would confuse
    warehouse staff who think in the local calendar day. The sequence resets
    to 001 each local day and is assigned via an atomic UPSERT+RETURNING, so
    two tasks created at the exact same instant can never collide (SQLite
    3.35+ required for RETURNING — same minimum version this app already
    needs elsewhere). If a single day ever exceeds 999 tasks the number just
    grows past 3 digits (TK260910-1000) rather than erroring or wrapping.
    """
    date_key = (datetime.utcnow() + timedelta(hours=7)).strftime("%y%m%d")
    row = db.execute(
        "INSERT INTO task_code_counters (date_key, next_seq) VALUES (?, 1) "
        "ON CONFLICT(date_key) DO UPDATE SET next_seq = task_code_counters.next_seq + 1 "
        "RETURNING next_seq",
        (date_key,),
    ).fetchone()
    return f"TK{date_key}-{row['next_seq']:03d}"


def valid_thai_phone(contact: str) -> bool:
    digits = re.sub(r"\D", "", contact or "")
    return bool(re.match(r"^0\d{8,9}$", digits))


def validate_and_normalize_email(db, email: Optional[str], exclude_user_id: Optional[int] = None) -> Optional[str]:
    """Shared by register/profile-edit/driver-creation/admin-edit so the
    @central.co.th + uniqueness rules are enforced identically everywhere an
    email can be set, not just at registration."""
    if not email:
        return None
    email = email.strip().lower()
    if not email.endswith("@central.co.th"):
        raise HTTPException(400, "อีเมลต้องเป็นโดเมน @central.co.th เท่านั้น")
    existing = db.execute("SELECT id FROM users WHERE email=?", (email,)).fetchone()
    if existing and existing["id"] != exclude_user_id:
        raise HTTPException(400, "อีเมลนี้มีผู้ใช้งานแล้ว")
    return email


# ---------------------------------------------------------------------------
# Schema
# ---------------------------------------------------------------------------

SCHEMA = """
CREATE TABLE IF NOT EXISTS vehicles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT UNIQUE NOT NULL,
    vehicle_type TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'Available',
    battery_level INTEGER
);

CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id TEXT UNIQUE NOT NULL,
    email TEXT,
    full_name TEXT,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL CHECK(role IN ('USER','DRIVER','ADMIN')),
    cost_center TEXT,
    contact TEXT,
    driver_status TEXT DEFAULT 'Not Checked-in',
    checked_in_vehicle_id INTEGER REFERENCES vehicles(id),
    created_at TEXT NOT NULL,
    deleted_at TEXT
);

-- Only enforces uniqueness among rows that actually have an email set, so
-- accounts without one (the common case) never collide with each other.
CREATE UNIQUE INDEX IF NOT EXISTS idx_users_email ON users(email) WHERE email IS NOT NULL;

CREATE TABLE IF NOT EXISTS vehicle_types (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    type_key TEXT UNIQUE NOT NULL,
    type_name_th TEXT NOT NULL,
    description TEXT
);

CREATE TABLE IF NOT EXISTS driver_licenses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    driver_id INTEGER NOT NULL REFERENCES users(id),
    vehicle_type TEXT NOT NULL,
    expiry_date TEXT,
    UNIQUE(driver_id, vehicle_type)
);

CREATE TABLE IF NOT EXISTS zones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    zone_key TEXT UNIQUE NOT NULL,
    zone_name_th TEXT NOT NULL,
    code_pattern TEXT NOT NULL,
    allow_multiple INTEGER NOT NULL DEFAULT 0,
    example_code TEXT,
    sort_order INTEGER NOT NULL DEFAULT 0,
    free_text INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS vehicle_type_zones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    vehicle_type TEXT NOT NULL,
    zone_key TEXT NOT NULL,
    UNIQUE(vehicle_type, zone_key)
);

CREATE TABLE IF NOT EXISTS task_rules (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    from_zone TEXT NOT NULL,
    to_zone TEXT NOT NULL,
    task_type TEXT NOT NULL,
    UNIQUE(from_zone, to_zone)
);

-- Backs the TK{YYMMDD}-{seq} task code format: one row per calendar day
-- (Thailand local time), atomically incremented via UPSERT+RETURNING so
-- concurrent task creation can never hand out the same code twice.
CREATE TABLE IF NOT EXISTS task_code_counters (
    date_key TEXT PRIMARY KEY,
    next_seq INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS tasks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_code TEXT UNIQUE NOT NULL,
    request_type TEXT NOT NULL,
    scheduled_at TEXT,
    bu TEXT,
    from_zone TEXT NOT NULL,
    from_location TEXT NOT NULL,
    to_zone TEXT NOT NULL,
    to_location TEXT NOT NULL,
    task_type TEXT NOT NULL,
    pallet_qty INTEGER NOT NULL,
    priority TEXT NOT NULL,
    remark TEXT,
    barcode_pallet_id TEXT,
    photo_url TEXT,
    requester_id INTEGER NOT NULL REFERENCES users(id),
    status TEXT NOT NULL DEFAULT 'Waiting',
    current_driver_id INTEGER REFERENCES users(id),
    current_vehicle_id INTEGER REFERENCES vehicles(id),
    accepted_at TEXT,
    accept_battery_level INTEGER,
    estimated_arrival TEXT,
    created_at TEXT NOT NULL,
    started_at TEXT,
    completed_at TEXT,
    cancel_reason TEXT,
    waiting_alert_sent INTEGER NOT NULL DEFAULT 0,
    rating INTEGER,
    rating_comment TEXT,
    blocked_reason TEXT,
    blocked_at TEXT
);

CREATE TABLE IF NOT EXISTS task_assignments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_id INTEGER NOT NULL REFERENCES tasks(id),
    driver_id INTEGER NOT NULL REFERENCES users(id),
    vehicle_id INTEGER NOT NULL REFERENCES vehicles(id),
    assigned_by INTEGER NOT NULL REFERENCES users(id),
    assigned_at TEXT NOT NULL,
    unassigned_at TEXT,
    assignment_status TEXT NOT NULL DEFAULT 'Active',
    reason TEXT
);

-- A driver can never have two simultaneously-Active assignment rows on the
-- same task. This is the database-level backstop against race conditions
-- when multiple requests (self-assign/join/take-over/admin-add) hit the same
-- task at nearly the same instant — enforced atomically by SQLite itself,
-- independent of any application-level timing.
CREATE UNIQUE INDEX IF NOT EXISTS idx_one_active_assignment_per_driver
ON task_assignments(task_id, driver_id) WHERE assignment_status='Active';

CREATE TABLE IF NOT EXISTS task_status_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_id INTEGER NOT NULL REFERENCES tasks(id),
    from_status TEXT,
    to_status TEXT NOT NULL,
    changed_by INTEGER REFERENCES users(id),
    changed_at TEXT NOT NULL,
    note TEXT
);

CREATE TABLE IF NOT EXISTS task_messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_id INTEGER NOT NULL REFERENCES tasks(id),
    sender_id INTEGER NOT NULL REFERENCES users(id),
    message TEXT NOT NULL,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS team_chat_messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sender_id INTEGER NOT NULL REFERENCES users(id),
    message TEXT NOT NULL,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS breakdowns (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    target_type TEXT NOT NULL CHECK(target_type IN ('driver','vehicle')),
    target_id INTEGER NOT NULL,
    target_label TEXT,
    description TEXT,
    reported_by INTEGER NOT NULL REFERENCES users(id),
    status TEXT NOT NULL DEFAULT 'Open',
    created_at TEXT NOT NULL,
    resolved_at TEXT,
    resolution_note TEXT
);

CREATE TABLE IF NOT EXISTS audit_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    actor_id INTEGER REFERENCES users(id),
    action TEXT NOT NULL,
    details TEXT,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS notifications (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL REFERENCES users(id),
    message TEXT NOT NULL,
    task_id INTEGER,
    kind TEXT NOT NULL DEFAULT 'task',
    created_at TEXT NOT NULL,
    is_read INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS sessions (
    token TEXT PRIMARY KEY,
    user_id INTEGER NOT NULL REFERENCES users(id),
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS push_subscriptions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL REFERENCES users(id),
    endpoint TEXT UNIQUE NOT NULL,
    p256dh TEXT NOT NULL,
    auth TEXT NOT NULL,
    created_at TEXT NOT NULL
);

-- Small key/value table for server-wide settings that used to be stored as
-- files (the VAPID private key, notably) — a database row survives exactly
-- as long as the rest of the app's data does, unlike a file on disk, which
-- can vanish on any redeploy/container restart depending on the hosting
-- platform (guaranteed on Vercel; common on other serverless/autoscale setups).
CREATE TABLE IF NOT EXISTS app_config (
    key TEXT PRIMARY KEY,
    value TEXT
);

-- Task photos live in the database now too, for the same reason as the VAPID
-- key above: no dependency on a local disk that may not persist. Old
-- photo_url values pointing at /uploads/... (from before this change) keep
-- working via the static file mount further down for backward compatibility.
CREATE TABLE IF NOT EXISTS task_photos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_id INTEGER NOT NULL REFERENCES tasks(id),
    content_type TEXT NOT NULL,
    data BLOB NOT NULL,
    created_at TEXT NOT NULL
);
"""


def init_db():
    if IS_POSTGRES:
        with get_db() as db:
            # Vercel (and any serverless platform) can spin up several instances
            # that all hit init_db() at once on the first requests. Without
            # coordination they race to CREATE the same tables and one loses
            # with a "duplicate key ... pg_type" error, which then crashed
            # startup. A transaction-scoped advisory lock serializes this:
            # whichever instance grabs the lock first does the setup, the
            # others wait, then find the tables already there and do nothing.
            db.execute("SELECT pg_advisory_xact_lock(913472001)")
            existing = db.execute(
                "SELECT 1 FROM information_schema.tables WHERE table_name='users'"
            ).fetchone()
            fresh = not existing
            if fresh:
                db.executescript(_postgres_schema())
                seed(db)
            else:
                # Self-heal: a broken earlier deploy may have created the tables
                # but left master data (vehicle types, zones, etc.) empty, which
                # makes it impossible to create vehicles or tasks. Refill any
                # empty master table on every boot — cheap, and idempotent.
                seed_master_data(db)
        # No migrate_db() here: a Postgres database always starts from today's
        # schema, so there's never an older SQLite-era shape to evolve away from.
        return

    fresh = not os.path.exists(DB_PATH)
    with get_db() as db:
        db.executescript(SCHEMA)
        if fresh:
            seed(db)
        else:
            seed_master_data(db)  # same self-heal for the SQLite path
    if not fresh:
        migrate_db()


def _postgres_schema() -> str:
    """The same schema, translated for Postgres: AUTOINCREMENT -> SERIAL,
    BLOB -> BYTEA. GROUP_CONCAT-style aggregation isn't touched here since
    that's a query-time function handled separately in list_drivers()."""
    return (
        SCHEMA
        .replace("INTEGER PRIMARY KEY AUTOINCREMENT", "SERIAL PRIMARY KEY")
        .replace("BLOB", "BYTEA")
    )


def migrate_db():
    """Best-effort auto-migration for SQLite databases created by an older
    version of this app, so a schema change doesn't force deleting existing
    data every time. New installs never hit this — CREATE TABLE already
    matches SCHEMA. Requires SQLite 3.35+ (2021) for DROP COLUMN; if
    unavailable, logs a clear message instead of crashing on first use.
    Postgres never calls this — see init_db()."""
    with get_db() as db:
        for table, old_columns in (
            ("tasks", ["max_rack_level"]),
            ("vehicles", ["max_rack_level", "capacity_pallet"]),
        ):
            existing = [r["name"] for r in db.execute(f"PRAGMA table_info({table})").fetchall()]
            for col in old_columns:
                if col in existing:
                    try:
                        db.execute(f"ALTER TABLE {table} DROP COLUMN {col}")
                        log.info(f"Migrated database: dropped old column {table}.{col}")
                    except sqlite3.OperationalError as e:
                        log.warning(
                            f"Could not auto-migrate {table}.{col} ({e}). "
                            f"If you see schema errors, delete backend/scdc.db and restart "
                            f"(this resets all data) — your SQLite version may be too old "
                            f"for automatic migration (needs 3.35+)."
                        )

        for table, new_columns in (
            ("users", [("deleted_at", "TEXT"), ("email", "TEXT")]),
            ("breakdowns", [("resolution_note", "TEXT")]),
            ("tasks", [("bu", "TEXT")]),
            ("zones", [("free_text", "INTEGER NOT NULL DEFAULT 0")]),
        ):
            existing = [r["name"] for r in db.execute(f"PRAGMA table_info({table})").fetchall()]
            for col, coltype in new_columns:
                if col not in existing:
                    db.execute(f"ALTER TABLE {table} ADD COLUMN {col} {coltype}")
                    log.info(f"Migrated database: added new column {table}.{col}")

        # Data fixes for databases that existed before this round: GATE now
        # allows multiple locations, and there's a new EMPTY_PALLET zone for
        # requisitioning empty pallets. Fresh installs already get both via
        # seed() — this only matters for upgrading an existing database.
        db.execute("UPDATE zones SET allow_multiple=1 WHERE zone_key='GATE'")
        if not db.execute("SELECT 1 FROM zones WHERE zone_key='EMPTY_PALLET'").fetchone():
            db.execute(
                "INSERT INTO zones (zone_key, zone_name_th, code_pattern, allow_multiple, example_code, sort_order, free_text) "
                "VALUES ('EMPTY_PALLET','พาเลทเปล่า','^.+$',0,'พาเลทเปล่า-จุดจ่าย',7,1)"
            )
            log.info("Migrated database: added EMPTY_PALLET zone")


def seed_master_data(db):
    """Master/reference data (vehicle types, zones, the type<->zone matrix,
    task rules). Split out from user seeding and made idempotent PER TABLE so
    it self-heals: if an earlier broken deploy left these tables empty (tables
    created but seed never completed), this refills whatever is missing on the
    next boot instead of staying empty forever. Safe to call every startup."""
    if not db.execute("SELECT 1 FROM vehicle_types LIMIT 1").fetchone():
        db.execute(
            "INSERT INTO vehicle_types (type_key, type_name_th, description) VALUES "
            "('RT','รถ Reach Truck','ยกสินค้าขึ้นชั้นวางสูง'),"
            "('PE','รถ Pallet Truck (ไฟฟ้า)','ลากพาเลทระยะสั้น'),"
            "('FORKLIFT','รถโฟล์คลิฟท์','ยกของทั่วไป')"
        )
        log.info("Seeded master data: vehicle_types")

    if not db.execute("SELECT 1 FROM zones LIMIT 1").fetchone():
        db.execute(
            "INSERT INTO zones (zone_key, zone_name_th, code_pattern, allow_multiple, example_code, sort_order, free_text) VALUES "
            "('CONCRETE_YARD','ลานปูน','^ลานปูน-[1-7]$',0,'ลานปูน-3',1,0),"
            "('GATE','ประตู','^ประตู-([1-9]|[1-7][0-9]|8[0-5])$',1,'ประตู-12',2,0),"
            "('SORT_YARD','ลาน Sort','^ลาน Sort$',0,'ลาน Sort',3,0),"
            "('MEZZANINE','Mezzanine','^Mezzanine-[1-4]-(ซ้าย \\(Cross dock\\)|ขวา \\(Double Deep\\))$',0,'Mezzanine-2-ซ้าย (Cross dock)',4,0),"
            "('SELECTIVE_RACK','Selective Rack','^B[AB][A-H](0[1-9]|[12][0-9]|3[0-9])([A-G][12])?$',1,'BAA01A1',5,0),"
            "('DOUBLE_DEEP','Double Deep','^AA[A-F](0[1-9]|[1-6][0-9]|7[0-4])([A-G][12])?$',1,'AAA01G1',6,0),"
            "('EMPTY_PALLET','พาเลทเปล่า','^.+$',0,'พาเลทเปล่า-จุดจ่าย',7,1)"
        )
        log.info("Seeded master data: zones")

    if not db.execute("SELECT 1 FROM vehicle_type_zones LIMIT 1").fetchone():
        # Default: every vehicle type can work every zone. Admin narrows this
        # down later via the vehicle-type <-> zone matrix in Settings.
        zone_keys = ["CONCRETE_YARD", "GATE", "SORT_YARD", "MEZZANINE", "SELECTIVE_RACK", "DOUBLE_DEEP"]
        for vt in ("RT", "PE", "FORKLIFT"):
            for zk in zone_keys:
                db.execute("INSERT INTO vehicle_type_zones (vehicle_type, zone_key) VALUES (?,?) ON CONFLICT DO NOTHING", (vt, zk))
        log.info("Seeded master data: vehicle_type_zones")

    if not db.execute("SELECT 1 FROM task_rules LIMIT 1").fetchone():
        db.execute(
            "INSERT INTO task_rules (from_zone, to_zone, task_type) VALUES "
            "('CONCRETE_YARD','SELECTIVE_RACK','Putaway'),"
            "('CONCRETE_YARD','DOUBLE_DEEP','Putaway'),"
            "('SELECTIVE_RACK','GATE','Picking'),"
            "('DOUBLE_DEEP','GATE','Picking'),"
            "('SELECTIVE_RACK','CONCRETE_YARD','Replenishment'),"
            "('DOUBLE_DEEP','CONCRETE_YARD','Replenishment')"
        )
        log.info("Seeded master data: task_rules")


def seed(db):
    ts = now_iso()

    def add_user(employee_id, password, role, full_name=None, cost_center=None, contact=None, pending=False):
        cur = db.execute(
            "INSERT INTO users (employee_id, full_name, password_hash, role, cost_center, contact, "
            "driver_status, created_at) VALUES (?,?,?,?,?,?,?,?) RETURNING id",
            (employee_id, full_name, PENDING_PASSWORD if pending else hash_password(password), role,
             cost_center, contact, "Not Checked-in" if role == "DRIVER" else None, ts),
        )
        return cur.fetchone()["id"]

    # First-run bootstrap only: create exactly one admin account with a
    # securely random password (never hardcoded), so there's a way to log in
    # and set up real users/vehicles from the app itself. Printed once to the
    # server console/log — grab it now, log in, and change the password
    # immediately (Profile -> เปลี่ยนรหัสผ่าน). No demo USER/DRIVER accounts
    # or demo vehicles are created; add real ones from the Admin UI.
    bootstrap_password = secrets.token_urlsafe(9)
    add_user("admin", bootstrap_password, "ADMIN", full_name="Administrator")
    log.warning("=" * 60)
    log.warning("FIRST-TIME SETUP - INITIAL ADMIN ACCOUNT CREATED")
    log.warning("  Employee ID: admin")
    log.warning("  Password:    %s", bootstrap_password)
    log.warning("  Log in now and change this password immediately.")
    log.warning("  This message only appears once, on first run (when")
    log.warning("  scdc.db does not exist yet). Write the password down now.")
    log.warning("=" * 60)

    seed_master_data(db)


# ---------------------------------------------------------------------------
# Task type inference (data-driven: zones + task_rules tables)
# ---------------------------------------------------------------------------

def infer_task_type(db, from_zone: str, to_zone: str) -> str:
    row = db.execute("SELECT task_type FROM task_rules WHERE from_zone=? AND to_zone=?", (from_zone, to_zone)).fetchone()
    return row["task_type"] if row else "Transfer"


def validate_location_string(db, zone_key: str, value: str) -> str:
    zone = db.execute("SELECT * FROM zones WHERE zone_key=?", (zone_key,)).fetchone()
    if not zone:
        raise HTTPException(400, f"ไม่พบโซน {zone_key}")
    value = (value or "").strip()
    if not value:
        raise HTTPException(400, "กรุณาระบุตำแหน่ง")
    raw_parts = [p.strip() for p in value.split(",") if p.strip()]
    # Silently drop duplicate codes (case-insensitive, order preserved) - e.g. pasted
    # from Excel with repeats. Keep original casing since some zones (Mezzanine, Sort
    # Yard) have meaningful mixed-case Thai/English text, not just alphanumeric codes.
    seen = set()
    parts = []
    for p in raw_parts:
        key = p.upper()
        if key not in seen:
            seen.add(key)
            parts.append(p)
    if not zone["allow_multiple"] and len(parts) > 1:
        raise HTTPException(400, f"โซน {zone['zone_name_th']} เลือกได้จุดเดียวเท่านั้น")
    if zone["free_text"]:
        return ",".join(parts)
    pattern = re.compile(zone["code_pattern"], re.IGNORECASE)
    for p in parts:
        if not pattern.match(p):
            raise HTTPException(400, f"รูปแบบตำแหน่งไม่ถูกต้อง: '{p}' (ตัวอย่างที่ถูกต้อง: {zone['example_code']})")
    return ",".join(parts)


def zone_capability_issue(db, vehicle_type: str, zone_key: str) -> Optional[str]:
    row = db.execute(
        "SELECT 1 FROM vehicle_type_zones WHERE vehicle_type=? AND zone_key=?", (vehicle_type, zone_key)
    ).fetchone()
    if row:
        return None
    zone = db.execute("SELECT zone_name_th FROM zones WHERE zone_key=?", (zone_key,)).fetchone()
    zone_name = zone["zone_name_th"] if zone else zone_key
    return f"ประเภทรถ {vehicle_type} ไม่สามารถทำงานในโซน {zone_name} ได้"


# ---------------------------------------------------------------------------
# State machine
# ---------------------------------------------------------------------------

ALLOWED_TRANSITIONS = {
    "Waiting": {"Assigned", "Cancelled"},
    "Assigned": {"In Progress", "Cancelled", "Waiting"},
    "In Progress": {"Completed", "Cancelled", "Pause"},
    "Pause": {"In Progress", "Cancelled", "Waiting"},
    "Completed": set(),
    "Cancelled": set(),
}


def transition(db, task_row, to_status: str, changed_by: int, note: str = None):
    from_status = task_row["status"]
    if to_status not in ALLOWED_TRANSITIONS.get(from_status, set()):
        raise HTTPException(400, f"Invalid transition: {from_status} -> {to_status}")
    db.execute("UPDATE tasks SET status=? WHERE id=?", (to_status, task_row["id"]))
    db.execute(
        "INSERT INTO task_status_history (task_id, from_status, to_status, changed_by, changed_at, note) "
        "VALUES (?,?,?,?,?,?)",
        (task_row["id"], from_status, to_status, changed_by, now_iso(), note),
    )


def notify(db, user_id: int, message: str, task_id: Optional[int] = None, kind: str = "task"):
    db.execute(
        "INSERT INTO notifications (user_id, message, task_id, kind, created_at) VALUES (?,?,?,?,?)",
        (user_id, message, task_id, kind, now_iso()),
    )
    push_url = f"/?task={task_id}" if (kind in ("task", "chat") and task_id) else "/"
    send_web_push_to_user(db, user_id, "SCDC", message, push_url)


def notify_admins(db, message: str, task_id: Optional[int] = None, kind: str = "task"):
    for a in db.execute("SELECT id FROM users WHERE role='ADMIN'").fetchall():
        notify(db, a["id"], message, task_id, kind)


def audit(db, actor_id: Optional[int], action: str, details: str = ""):
    db.execute(
        "INSERT INTO audit_logs (actor_id, action, details, created_at) VALUES (?,?,?,?)",
        (actor_id, action, details, now_iso()),
    )


def free_up_driver_and_vehicle(db, driver_id: Optional[int], vehicle_id: Optional[int]):
    if driver_id:
        row = db.execute("SELECT driver_status FROM users WHERE id=?", (driver_id,)).fetchone()
        if row and row["driver_status"] not in ("Breakdown",):
            db.execute("UPDATE users SET driver_status='Available' WHERE id=?", (driver_id,))
    if vehicle_id:
        row = db.execute("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
        if row and row["status"] != "Breakdown":
            db.execute("UPDATE vehicles SET status='Available' WHERE id=?", (vehicle_id,))


def check_waiting_too_long(db):
    cutoff = (datetime.utcnow() - timedelta(minutes=WAITING_TOO_LONG_MINUTES)).isoformat(timespec="seconds") + "Z"
    rows = db.execute(
        "SELECT * FROM tasks WHERE status='Waiting' AND waiting_alert_sent=0 AND created_at<?", (cutoff,)
    ).fetchall()
    for t in rows:
        notify_admins(db, f"Task {t['task_code']} has been waiting over {WAITING_TOO_LONG_MINUTES} min", t["id"])
        db.execute("UPDATE tasks SET waiting_alert_sent=1 WHERE id=?", (t["id"],))


def task_chat_allowed(db, task_row, user) -> bool:
    if user["role"] == "ADMIN":
        return True
    if task_row["requester_id"] == user["id"]:
        return True
    ever_driver = db.execute(
        "SELECT 1 FROM task_assignments WHERE task_id=? AND driver_id=?", (task_row["id"], user["id"])
    ).fetchone()
    return bool(ever_driver)


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

class LoginBody(BaseModel):
    employee_id: str  # accepts either an employee ID or a @central.co.th email — see login()
    password: str


class RegisterBody(BaseModel):
    employee_id: str
    email: Optional[str] = None
    full_name: str
    password: str
    cost_center: str = Field(..., min_length=5, max_length=5)
    contact: str


class SetPasswordBody(BaseModel):
    employee_id: str
    new_password: str


class ChangePasswordBody(BaseModel):
    old_password: str
    new_password: str


class ProfileUpdateBody(BaseModel):
    full_name: Optional[str] = None
    contact: Optional[str] = None
    cost_center: Optional[str] = None
    email: Optional[str] = None


def current_user(authorization: Optional[str] = Header(None)):
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(401, "Missing token")
    token = authorization.removeprefix("Bearer ").strip()
    with get_db() as db:
        row = db.execute(
            "SELECT users.* FROM sessions JOIN users ON users.id = sessions.user_id WHERE token=?",
            (token,),
        ).fetchone()
    if not row or row["deleted_at"]:
        raise HTTPException(401, "Invalid or expired token")
    return dict(row)


def require_role(*roles):
    def checker(user=Depends(current_user)):
        if user["role"] not in roles:
            raise HTTPException(403, f"Requires role: {roles}")
        return user
    return checker


@app.post("/auth/login")
def login(body: LoginBody):
    identifier = body.employee_id.strip()
    with get_db() as db:
        row = db.execute(
            "SELECT * FROM users WHERE employee_id=? OR email=?", (identifier, identifier)
        ).fetchone()
        if not row or row["deleted_at"]:
            raise HTTPException(404, "ยังไม่มีบัญชีนี้ โปรดสมัครสมาชิก")
        if row["password_hash"] == PENDING_PASSWORD:
            raise HTTPException(428, "First login: please set your password")
        if not verify_password(body.password, row["password_hash"]):
            raise HTTPException(401, "รหัสผ่านไม่ถูกต้อง หากลืมรหัสโปรดติดต่อ Admin")
        if row["role"] == "DRIVER":
            # Every fresh login starts "not on a vehicle yet" until they check in again.
            db.execute(
                "UPDATE users SET driver_status='Not Checked-in', checked_in_vehicle_id=NULL "
                "WHERE id=? AND driver_status NOT IN ('Busy','Breakdown')",
                (row["id"],),
            )
        # Single active session per account: a fresh login anywhere signs out
        # any other device/browser this account was logged into elsewhere.
        db.execute("DELETE FROM sessions WHERE user_id=?", (row["id"],))
        token = secrets.token_urlsafe(32)
        db.execute(
            "INSERT INTO sessions (token, user_id, created_at) VALUES (?,?,?)",
            (token, row["id"], now_iso()),
        )
        return {"token": token, "role": row["role"], "employee_id": row["employee_id"]}


@app.post("/auth/register")
def register(body: RegisterBody):
    with get_db() as db:
        if db.execute("SELECT 1 FROM users WHERE employee_id=?", (body.employee_id,)).fetchone():
            raise HTTPException(400, "รหัสพนักงานนี้มีผู้ใช้งานแล้ว")
        email = validate_and_normalize_email(db, body.email)
        if not body.cost_center.isdigit():
            raise HTTPException(400, "Cost center must be 5 digits")
        if not body.full_name.strip():
            raise HTTPException(400, "Full name is required")
        if not valid_thai_phone(body.contact):
            raise HTTPException(400, "Contact number looks invalid (expects a 9-10 digit Thai phone number)")
        if len(body.password) < MIN_PASSWORD_LENGTH:
            raise HTTPException(400, f"รหัสผ่านต้องมีอย่างน้อย {MIN_PASSWORD_LENGTH} ตัวอักษร")
        db.execute(
            "INSERT INTO users (employee_id, email, full_name, password_hash, role, cost_center, contact, created_at) "
            "VALUES (?,?,?,?,?,?,?,?)",
            (body.employee_id, email, body.full_name.strip(), hash_password(body.password), "USER",
             body.cost_center, body.contact, now_iso()),
        )
        return {"ok": True}


@app.post("/auth/set-password")
def set_password(body: SetPasswordBody):
    if len(body.new_password) < MIN_PASSWORD_LENGTH:
        raise HTTPException(400, f"รหัสผ่านต้องมีอย่างน้อย {MIN_PASSWORD_LENGTH} ตัวอักษร")
    with get_db() as db:
        row = db.execute("SELECT * FROM users WHERE employee_id=?", (body.employee_id,)).fetchone()
        if not row or row["password_hash"] != PENDING_PASSWORD:
            raise HTTPException(400, "This account does not need first-time password setup")
        db.execute("UPDATE users SET password_hash=? WHERE id=?", (hash_password(body.new_password), row["id"]))
        return {"ok": True}


@app.post("/auth/change-password")
def change_password(body: ChangePasswordBody, user=Depends(current_user)):
    if len(body.new_password) < MIN_PASSWORD_LENGTH:
        raise HTTPException(400, f"รหัสผ่านต้องมีอย่างน้อย {MIN_PASSWORD_LENGTH} ตัวอักษร")
    with get_db() as db:
        row = db.execute("SELECT * FROM users WHERE id=?", (user["id"],)).fetchone()
        if not verify_password(body.old_password, row["password_hash"]):
            raise HTTPException(400, "Current password is incorrect")
        db.execute("UPDATE users SET password_hash=? WHERE id=?", (hash_password(body.new_password), user["id"]))
        return {"ok": True}


@app.get("/me")
def me(user=Depends(current_user)):
    return {"employee_id": user["employee_id"], "role": user["role"], "id": user["id"],
            "full_name": user["full_name"], "contact": user["contact"], "cost_center": user["cost_center"],
            "email": user["email"],
            "driver_status": user["driver_status"], "checked_in_vehicle_id": user["checked_in_vehicle_id"]}


@app.put("/me")
def update_me(body: ProfileUpdateBody, user=Depends(current_user)):
    """Everyone can edit their own name/contact/email; only USER has a cost center to edit.
    Employee ID and role are never editable here."""
    fields = {}
    if body.full_name is not None:
        if not body.full_name.strip():
            raise HTTPException(400, "Full name cannot be empty")
        fields["full_name"] = body.full_name.strip()
    if body.contact is not None:
        if not valid_thai_phone(body.contact):
            raise HTTPException(400, "Contact number looks invalid (expects a 9-10 digit Thai phone number)")
        fields["contact"] = body.contact
    if body.cost_center is not None:
        if user["role"] != "USER":
            raise HTTPException(400, "Only USER accounts have a cost center")
        if not body.cost_center.isdigit() or len(body.cost_center) != 5:
            raise HTTPException(400, "Cost center must be 5 digits")
        fields["cost_center"] = body.cost_center
    if not fields and body.email is None:
        return {"ok": True}
    with get_db() as db:
        if body.email is not None:
            fields["email"] = validate_and_normalize_email(db, body.email, exclude_user_id=user["id"])
        if not fields:
            return {"ok": True}
        set_clause = ", ".join(f"{k}=?" for k in fields.keys())
        db.execute(f"UPDATE users SET {set_clause} WHERE id=?", list(fields.values()) + [user["id"]])
        return {"ok": True}


# ---------------------------------------------------------------------------
# Admin: create driver accounts (first-login password setup)
# ---------------------------------------------------------------------------

class NewDriverBody(BaseModel):
    employee_id: str
    email: Optional[str] = None
    full_name: Optional[str] = None
    initial_vehicle_types: List[str] = []


@app.post("/admin/drivers")
def create_driver(body: NewDriverBody, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        if db.execute("SELECT 1 FROM users WHERE employee_id=?", (body.employee_id,)).fetchone():
            raise HTTPException(400, "Employee ID already exists")
        email = validate_and_normalize_email(db, body.email)
        cur = db.execute(
            "INSERT INTO users (employee_id, email, full_name, password_hash, role, driver_status, created_at) "
            "VALUES (?,?,?,?,?,?,?) RETURNING id",
            (body.employee_id, email, body.full_name, PENDING_PASSWORD, "DRIVER", "Not Checked-in", now_iso()),
        )
        driver_id = cur.fetchone()["id"]
        for vt in body.initial_vehicle_types:
            db.execute(
                "INSERT INTO driver_licenses (driver_id, vehicle_type, expiry_date) VALUES (?,?,NULL) "
                "ON CONFLICT DO NOTHING",
                (driver_id, vt),
            )
        audit(db, user["id"], "create_driver", body.employee_id)
        return {"ok": True}


@app.get("/admin/users")
def list_all_users(user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        rows = db.execute(
            "SELECT id, employee_id, email, full_name, role, contact, cost_center, driver_status "
            "FROM users WHERE deleted_at IS NULL ORDER BY role, employee_id"
        ).fetchall()
        return [dict(r) for r in rows]


class UserAdminUpdateBody(BaseModel):
    full_name: Optional[str] = None
    email: Optional[str] = None
    contact: Optional[str] = None
    cost_center: Optional[str] = None
    role: Optional[str] = None


@app.put("/admin/users/{user_id}")
def admin_update_user(user_id: int, body: UserAdminUpdateBody, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบผู้ใช้")
        fields = body.dict(exclude_unset=True, exclude={"email"})
        if "role" in fields and fields["role"] not in ("USER", "DRIVER", "ADMIN"):
            raise HTTPException(400, "role ไม่ถูกต้อง")
        if fields.get("cost_center"):
            if not fields["cost_center"].isdigit() or len(fields["cost_center"]) != 5:
                raise HTTPException(400, "Cost center ต้องเป็นตัวเลข 5 หลัก")
        if fields.get("contact") and not valid_thai_phone(fields["contact"]):
            raise HTTPException(400, "เบอร์ติดต่อไม่ถูกต้อง")
        if body.email is not None:
            fields["email"] = validate_and_normalize_email(db, body.email, exclude_user_id=user_id)
        if not fields:
            return {"ok": True}
        set_clause = ", ".join(f"{k}=?" for k in fields.keys())
        db.execute(f"UPDATE users SET {set_clause} WHERE id=?", list(fields.values()) + [user_id])
        if fields.get("role") == "DRIVER" and row["role"] != "DRIVER":
            db.execute("UPDATE users SET driver_status='Not Checked-in' WHERE id=?", (user_id,))
        audit(db, user["id"], "admin_update_user", f"{row['employee_id']}: {fields}")
        return {"ok": True}


@app.delete("/admin/users/{user_id}")
def admin_delete_user(user_id: int, user=Depends(require_role("ADMIN"))):
    """Soft-delete: keeps the row (and every historical task/audit/chat record
    that references it) but blocks login and hides the account from normal
    lists. A hard DELETE would either fail on foreign-key constraints (if the
    user has any history) or silently orphan that history (if constraints
    were relaxed) — neither is acceptable for a real deployment."""
    with get_db() as db:
        row = db.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบผู้ใช้")
        if row["id"] == user["id"]:
            raise HTTPException(400, "ลบบัญชีตัวเองไม่ได้")
        if row["role"] == "ADMIN":
            other_admins = db.execute(
                "SELECT COUNT(*) n FROM users WHERE role='ADMIN' AND deleted_at IS NULL AND id!=?", (user_id,)
            ).fetchone()["n"]
            if other_admins == 0:
                raise HTTPException(400, "ลบไม่ได้ เพราะเป็น Admin คนสุดท้ายในระบบ")
        db.execute("UPDATE users SET deleted_at=? WHERE id=?", (now_iso(), user_id))
        db.execute("DELETE FROM sessions WHERE user_id=?", (user_id,))  # force logout everywhere
        audit(db, user["id"], "delete_user", row["employee_id"])
        return {"ok": True}


@app.post("/admin/users/{user_id}/reset-password")
def admin_reset_password(user_id: int, user=Depends(require_role("ADMIN"))):
    """'Forgot password' for Admin to trigger on someone else's behalf: puts
    the account back into the same first-login state new accounts start in,
    so the user logs in with any password and is prompted to set a new one
    (reuses the existing /auth/set-password flow, no new mechanism needed)."""
    with get_db() as db:
        row = db.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบผู้ใช้")
        db.execute("UPDATE users SET password_hash=? WHERE id=?", (PENDING_PASSWORD, user_id))
        db.execute("DELETE FROM sessions WHERE user_id=?", (user_id,))  # force logout everywhere
        audit(db, user["id"], "reset_password", row["employee_id"])
        return {"ok": True}


@app.get("/admin/users/{user_id}/licenses")
def get_driver_licenses(user_id: int, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        rows = db.execute("SELECT * FROM driver_licenses WHERE driver_id=?", (user_id,)).fetchall()
        return [dict(r) for r in rows]


class LicenseBody(BaseModel):
    vehicle_type: str
    expiry_date: Optional[str] = None  # None = permanent, never expires


@app.post("/admin/users/{user_id}/licenses")
def set_driver_license(user_id: int, body: LicenseBody, user=Depends(require_role("ADMIN"))):
    if body.expiry_date:
        try:
            datetime.strptime(body.expiry_date, "%Y-%m-%d")
        except ValueError:
            raise HTTPException(400, "วันหมดอายุต้องอยู่ในรูปแบบ YYYY-MM-DD (เช่น 2027-12-31)")
    with get_db() as db:
        db.execute(
            "INSERT INTO driver_licenses (driver_id, vehicle_type, expiry_date) VALUES (?,?,?) "
            "ON CONFLICT(driver_id, vehicle_type) DO UPDATE SET expiry_date=excluded.expiry_date",
            (user_id, body.vehicle_type, body.expiry_date),
        )
        audit(db, user["id"], "set_license", f"driver {user_id}: {body.vehicle_type} exp={body.expiry_date}")
        return {"ok": True}


@app.delete("/admin/users/{user_id}/licenses/{vehicle_type}")
def delete_driver_license(user_id: int, vehicle_type: str, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        db.execute("DELETE FROM driver_licenses WHERE driver_id=? AND vehicle_type=?", (user_id, vehicle_type))
        audit(db, user["id"], "delete_license", f"driver {user_id}: {vehicle_type}")
        return {"ok": True}


# ---------------------------------------------------------------------------
# Zones, Vehicle-Type <-> Zone capability matrix, Task Rules (Admin-configurable)
# ---------------------------------------------------------------------------

class ZoneBody(BaseModel):
    zone_key: str
    zone_name_th: str
    code_pattern: str
    allow_multiple: bool = False
    example_code: Optional[str] = None
    free_text: bool = False


@app.get("/zones")
def list_zones(user=Depends(current_user)):
    with get_db() as db:
        return [dict(r) for r in db.execute("SELECT * FROM zones ORDER BY sort_order, id").fetchall()]


@app.post("/zones")
def add_zone(body: ZoneBody, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        if db.execute("SELECT 1 FROM zones WHERE zone_key=?", (body.zone_key,)).fetchone():
            raise HTTPException(400, "รหัสโซนนี้มีอยู่แล้ว")
        try:
            re.compile(body.code_pattern)
        except re.error as e:
            raise HTTPException(400, f"รูปแบบ pattern ไม่ถูกต้อง: {e}")
        max_order = db.execute("SELECT COALESCE(MAX(sort_order),0) m FROM zones").fetchone()["m"]
        db.execute(
            "INSERT INTO zones (zone_key, zone_name_th, code_pattern, allow_multiple, example_code, sort_order, free_text) "
            "VALUES (?,?,?,?,?,?,?)",
            (body.zone_key, body.zone_name_th, body.code_pattern, int(body.allow_multiple),
             body.example_code, max_order + 1, int(body.free_text)),
        )
        audit(db, user["id"], "add_zone", body.zone_key)
        return {"ok": True}


@app.delete("/zones/{zone_key}")
def delete_zone(zone_key: str, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        in_use = db.execute(
            "SELECT 1 FROM task_rules WHERE from_zone=? OR to_zone=?", (zone_key, zone_key)
        ).fetchone()
        if in_use:
            raise HTTPException(400, "ลบไม่ได้ เพราะยังมีกฎประเภทงานอ้างอิงโซนนี้อยู่ กรุณาลบกฎที่เกี่ยวข้องก่อน")
        db.execute("DELETE FROM zones WHERE zone_key=?", (zone_key,))
        db.execute("DELETE FROM vehicle_type_zones WHERE zone_key=?", (zone_key,))
        audit(db, user["id"], "delete_zone", zone_key)
        return {"ok": True}


@app.get("/vehicle-type-zones")
def get_vehicle_type_zones(user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        vehicle_types = [r["vehicle_type"] for r in db.execute(
            "SELECT DISTINCT vehicle_type FROM vehicles ORDER BY vehicle_type"
        ).fetchall()]
        zones = [dict(r) for r in db.execute("SELECT zone_key, zone_name_th FROM zones ORDER BY sort_order, id").fetchall()]
        rows = db.execute("SELECT vehicle_type, zone_key FROM vehicle_type_zones").fetchall()
        matrix = {}
        for vt in vehicle_types:
            matrix[vt] = {z["zone_key"]: False for z in zones}
        for r in rows:
            if r["vehicle_type"] in matrix:
                matrix[r["vehicle_type"]][r["zone_key"]] = True
        return {"vehicle_types": vehicle_types, "zones": zones, "matrix": matrix}


class VehicleTypeZoneBody(BaseModel):
    vehicle_type: str
    zone_key: str
    allowed: bool


@app.post("/vehicle-type-zones")
def set_vehicle_type_zone(body: VehicleTypeZoneBody, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        if body.allowed:
            db.execute(
                "INSERT INTO vehicle_type_zones (vehicle_type, zone_key) VALUES (?,?) "
                "ON CONFLICT DO NOTHING",
                (body.vehicle_type, body.zone_key),
            )
        else:
            db.execute(
                "DELETE FROM vehicle_type_zones WHERE vehicle_type=? AND zone_key=?",
                (body.vehicle_type, body.zone_key),
            )
        audit(db, user["id"], "set_vehicle_type_zone", f"{body.vehicle_type}/{body.zone_key}={body.allowed}")
        return {"ok": True}


class TaskRuleBody(BaseModel):
    from_zone: str
    to_zone: str
    task_type: str


@app.get("/task-rules")
def list_task_rules(user=Depends(current_user)):
    with get_db() as db:
        return [dict(r) for r in db.execute("SELECT * FROM task_rules ORDER BY id").fetchall()]


@app.post("/task-rules")
def add_task_rule(body: TaskRuleBody, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        db.execute(
            "INSERT INTO task_rules (from_zone, to_zone, task_type) VALUES (?,?,?) "
            "ON CONFLICT(from_zone, to_zone) DO UPDATE SET task_type=excluded.task_type",
            (body.from_zone, body.to_zone, body.task_type),
        )
        audit(db, user["id"], "set_task_rule", f"{body.from_zone}->{body.to_zone}={body.task_type}")
        return {"ok": True}


@app.delete("/task-rules/{rule_id}")
def delete_task_rule(rule_id: int, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        db.execute("DELETE FROM task_rules WHERE id=?", (rule_id,))
        audit(db, user["id"], "delete_task_rule", str(rule_id))
        return {"ok": True}


# ---------------------------------------------------------------------------
# Tasks
# ---------------------------------------------------------------------------

BU_LIST = ["CDS", "MUJI", "PGE", "RBS", "SSP", "GG", "B2S", "OFM", "CMG", "KMNY"]


class TaskCreate(BaseModel):
    request_type: str = "Request Now"
    scheduled_at: Optional[str] = None
    bu: str
    from_zone: str
    from_location: str
    to_zone: str
    to_location: str
    pallet_qty: int = Field(..., gt=0)
    priority: str = "Normal"
    remark: Optional[str] = None
    barcode_pallet_id: Optional[str] = None


class TaskUpdate(BaseModel):
    from_zone: Optional[str] = None
    from_location: Optional[str] = None
    to_zone: Optional[str] = None
    to_location: Optional[str] = None
    scheduled_at: Optional[str] = None
    pallet_qty: Optional[int] = Field(None, gt=0)
    priority: Optional[str] = None
    remark: Optional[str] = None
    barcode_pallet_id: Optional[str] = None


class PriorityBody(BaseModel):
    priority: str


def task_to_dict(row, db, viewer_role: Optional[str] = None) -> dict:
    d = dict(row)
    # Rating/comment are hidden from Drivers specifically (so they don't see
    # how they were scored) — but not from the User who submitted it, since
    # they need `rating` to tell "already reviewed" from "awaiting review"
    # (hiding it from them too broke that tracking: the rate button would
    # never disappear). Admin always sees everything.
    if viewer_role == "DRIVER":
        d["rating"] = None
        d["rating_comment"] = None
    d["driver_employee_id"] = None
    d["driver_name"] = None
    d["vehicle_code"] = None
    if d.get("current_driver_id"):
        u = db.execute("SELECT employee_id, full_name FROM users WHERE id=?", (d["current_driver_id"],)).fetchone()
        if u:
            d["driver_employee_id"] = u["employee_id"]
            d["driver_name"] = u["full_name"]
    if d.get("current_vehicle_id"):
        v = db.execute("SELECT code FROM vehicles WHERE id=?", (d["current_vehicle_id"],)).fetchone()
        d["vehicle_code"] = v["code"] if v else None
    req = db.execute("SELECT employee_id, full_name, contact FROM users WHERE id=?", (d["requester_id"],)).fetchone()
    if req:
        d["requester_employee_id"] = req["employee_id"]
        d["requester_name"] = req["full_name"]
        d["requester_contact"] = req["contact"]
    # Multi-driver: everyone currently Active on this task (primary = current_driver_id, first joined).
    active_rows = db.execute(
        "SELECT task_assignments.*, users.employee_id, users.full_name FROM task_assignments "
        "JOIN users ON users.id = task_assignments.driver_id "
        "WHERE task_id=? AND assignment_status='Active' ORDER BY assigned_at", (d["id"],)
    ).fetchall()
    drivers = []
    for r in active_rows:
        v = db.execute("SELECT code FROM vehicles WHERE id=?", (r["vehicle_id"],)).fetchone()
        drivers.append({
            "employee_id": r["employee_id"], "full_name": r["full_name"],
            "vehicle_code": v["code"] if v else None,
        })
    d["drivers"] = drivers
    d["other_driver_count"] = max(0, len(drivers) - 1)
    return d


@app.post("/tasks")
def create_task(body: TaskCreate, user=Depends(require_role("USER"))):
    if body.request_type == "Booking" and not body.scheduled_at:
        raise HTTPException(400, "การจองล่วงหน้าต้องระบุวันเวลา")
    if body.bu not in BU_LIST:
        raise HTTPException(400, f"BU ไม่ถูกต้อง ต้องเป็นหนึ่งใน: {', '.join(BU_LIST)}")
    with get_db() as db:
        from_location = validate_location_string(db, body.from_zone, body.from_location)
        to_location = validate_location_string(db, body.to_zone, body.to_location)
        task_type = infer_task_type(db, body.from_zone, body.to_zone)
        code = generate_task_code(db)
        cur = db.execute(
            "INSERT INTO tasks (task_code, request_type, scheduled_at, bu, from_zone, from_location, to_zone, "
            "to_location, task_type, pallet_qty, priority, remark, barcode_pallet_id, "
            "requester_id, status, created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) RETURNING id",
            (code, body.request_type, body.scheduled_at, body.bu, body.from_zone, from_location, body.to_zone,
             to_location, task_type, body.pallet_qty, body.priority, body.remark,
             body.barcode_pallet_id, user["id"], "Waiting", now_iso()),
        )
        task_id = cur.fetchone()["id"]
        db.execute(
            "INSERT INTO task_status_history (task_id, from_status, to_status, changed_by, changed_at) "
            "VALUES (?,?,?,?,?)", (task_id, None, "Waiting", user["id"], now_iso()),
        )
        notify_admins(db, f"New task {code} created by {user['employee_id']}", task_id)
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        result = task_to_dict(row, db, viewer_role=user["role"])
        available = db.execute(
            "SELECT COUNT(*) n FROM users WHERE role='DRIVER' AND driver_status='Available' AND deleted_at IS NULL"
        ).fetchone()["n"]
        if available == 0:
            result["warning"] = "ตอนนี้ยังไม่มีคนขับว่าง งานอาจต้องใช้เวลาสักหน่อยกว่าจะมีคนขับรับงาน"
        return result


@app.put("/tasks/{task_id}")
def update_task(task_id: int, body: TaskUpdate, user=Depends(require_role("USER"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if row["requester_id"] != user["id"]:
            raise HTTPException(403, "Not your task")
        if row["status"] != "Waiting":
            raise HTTPException(400, "Task can only be modified before work starts (status=Waiting)")
        fields = body.dict(exclude_unset=True)
        if not fields:
            return task_to_dict(row, db, viewer_role=user["role"])
        from_zone = fields.get("from_zone", row["from_zone"])
        to_zone = fields.get("to_zone", row["to_zone"])
        if "from_location" in fields or "from_zone" in fields:
            fields["from_location"] = validate_location_string(db, from_zone, fields.get("from_location", row["from_location"]))
            fields["from_zone"] = from_zone
        if "to_location" in fields or "to_zone" in fields:
            fields["to_location"] = validate_location_string(db, to_zone, fields.get("to_location", row["to_location"]))
            fields["to_zone"] = to_zone
        task_type = infer_task_type(db, from_zone, to_zone)
        set_clause = ", ".join(f"{k}=?" for k in fields.keys())
        params = list(fields.values()) + [task_type, task_id]
        db.execute(f"UPDATE tasks SET {set_clause}, task_type=? WHERE id=?", params)
        audit(db, user["id"], "modify_task", f"task {row['task_code']}: {fields}")
        fresh = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        return task_to_dict(fresh, db, viewer_role=user["role"])


@app.post("/tasks/{task_id}/photo")
def upload_task_photo(task_id: int, file: UploadFile = File(...), user=Depends(require_role("USER"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row or row["requester_id"] != user["id"]:
            raise HTTPException(404, "Task not found")
        content = file.file.read()
        content_type = file.content_type or "image/jpeg"
        cur = db.execute(
            "INSERT INTO task_photos (task_id, content_type, data, created_at) VALUES (?,?,?,?) RETURNING id",
            (task_id, content_type, content, now_iso()),
        )
        photo_id = cur.fetchone()["id"]
        url = f"/tasks/{task_id}/photo-data/{photo_id}"
        db.execute("UPDATE tasks SET photo_url=? WHERE id=?", (url, task_id))
        return {"ok": True, "photo_url": url}


@app.get("/tasks/{task_id}/photo-data/{photo_id}")
def get_task_photo(task_id: int, photo_id: int, user=Depends(current_user)):
    with get_db() as db:
        row = db.execute(
            "SELECT * FROM task_photos WHERE id=? AND task_id=?", (photo_id, task_id)
        ).fetchone()
        if not row:
            raise HTTPException(404, "Photo not found")
        return Response(content=bytes(row["data"]), media_type=row["content_type"])


@app.post("/tasks/{task_id}/priority")
def change_priority(task_id: int, body: PriorityBody, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if row["status"] in ("Completed", "Cancelled"):
            raise HTTPException(400, f"Cannot change priority of a {row['status']} task")
        old = row["priority"]
        db.execute("UPDATE tasks SET priority=? WHERE id=?", (body.priority, task_id))
        audit(db, user["id"], "change_priority", f"task {row['task_code']}: {old} -> {body.priority}")
        notify(db, row["requester_id"], f"Task {row['task_code']} priority changed to {body.priority}", task_id)
        return {"ok": True}


@app.get("/tasks")
def list_tasks(status_filter: Optional[str] = None, user=Depends(current_user)):
    with get_db() as db:
        if user["role"] == "ADMIN":
            check_waiting_too_long(db)
        if user["role"] == "USER":
            q = "SELECT * FROM tasks WHERE requester_id=?"
            params = [user["id"]]
        elif user["role"] == "DRIVER":
            # Drivers see: any task they've ever had an assignment on (any status,
            # for history — including as a joiner, not just primary driver), plus
            # every non-terminal task (so they can self-assign/join/take-over).
            q = ("SELECT DISTINCT tasks.* FROM tasks "
                 "LEFT JOIN task_assignments ON task_assignments.task_id = tasks.id AND task_assignments.driver_id=? "
                 "WHERE (task_assignments.driver_id IS NOT NULL OR tasks.status NOT IN ('Completed','Cancelled'))")
            params = [user["id"]]
        else:
            q = "SELECT * FROM tasks WHERE 1=1"
            params = []
        if status_filter:
            q += " AND status=?"
            params.append(status_filter)
        q += " ORDER BY created_at DESC"
        rows = db.execute(q, params).fetchall()
        return [task_to_dict(r, db, viewer_role=user["role"]) for r in rows]


@app.get("/tasks/{task_id}")
def get_task(task_id: int, user=Depends(current_user)):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        history = db.execute(
            "SELECT * FROM task_assignments WHERE task_id=? ORDER BY assigned_at", (task_id,)
        ).fetchall()
        status_hist = db.execute(
            "SELECT * FROM task_status_history WHERE task_id=? ORDER BY changed_at", (task_id,)
        ).fetchall()
        d = task_to_dict(row, db, viewer_role=user["role"])
        d["assignment_history"] = [dict(h) for h in history]
        d["status_history"] = [dict(h) for h in status_hist]
        return d


@app.post("/tasks/{task_id}/cancel")
def cancel_task(task_id: int, reason: Optional[str] = None, user=Depends(current_user)):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if user["role"] == "USER":
            if row["requester_id"] != user["id"] or row["status"] not in ("Waiting", "Assigned"):
                raise HTTPException(403, "Cannot cancel this task")
        elif user["role"] != "ADMIN":
            raise HTTPException(403, "Not allowed")
        transition(db, row, "Cancelled", user["id"], note=reason)
        db.execute("UPDATE tasks SET cancel_reason=? WHERE id=?", (reason, task_id))
        free_up_driver_and_vehicle(db, row["current_driver_id"], row["current_vehicle_id"])
        db.execute(
            "UPDATE task_assignments SET unassigned_at=?, assignment_status='Cancelled' "
            "WHERE task_id=? AND assignment_status='Active'", (now_iso(), task_id),
        )
        audit(db, user["id"], "cancel_task", f"task {row['task_code']}: {reason or ''}")
        notify(db, row["requester_id"], f"Task {row['task_code']} was cancelled", task_id)
        if row["current_driver_id"]:
            notify(db, row["current_driver_id"], f"Task {row['task_code']} was cancelled", task_id)
        return {"ok": True}


class SendBackBody(BaseModel):
    reason: str


@app.post("/tasks/{task_id}/send-back")
def send_back_task(task_id: int, body: SendBackBody, user=Depends(require_role("ADMIN"))):
    """Admin kicks a Waiting task back to the User to fix (e.g. not enough PE
    vehicles right now) instead of leaving it stuck or cancelling outright."""
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if row["status"] != "Waiting":
            raise HTTPException(400, "ตีกลับได้เฉพาะงานที่ยังรอดำเนินการ (Waiting) เท่านั้น")
        db.execute("UPDATE tasks SET blocked_reason=?, blocked_at=? WHERE id=?", (body.reason, now_iso(), task_id))
        audit(db, user["id"], "send_back_task", f"task {row['task_code']}: {body.reason}")
        notify(db, row["requester_id"], f"งาน {row['task_code']} ถูกตีกลับ: {body.reason}", task_id)
        return {"ok": True}


@app.post("/tasks/{task_id}/resubmit")
def resubmit_task(task_id: int, user=Depends(require_role("USER"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if row["requester_id"] != user["id"]:
            raise HTTPException(403, "Not your task")
        if row["status"] != "Waiting" or not row["blocked_reason"]:
            raise HTTPException(400, "งานนี้ไม่ได้ถูกตีกลับ")
        db.execute(
            "UPDATE tasks SET blocked_reason=NULL, blocked_at=NULL, waiting_alert_sent=0 WHERE id=?", (task_id,)
        )
        audit(db, user["id"], "resubmit_task", f"task {row['task_code']}")
        notify_admins(db, f"งาน {row['task_code']} ถูกส่งกลับมาใหม่โดย {user['employee_id']}", task_id)
        return {"ok": True}


class AssignBody(BaseModel):
    driver_employee_id: str
    reason: Optional[str] = None
    estimated_minutes: Optional[int] = None
    force: bool = False


_STATUS_TH = {
    "Available": "ว่าง", "Busy": "ไม่ว่าง", "Pause": "พัก", "Breakdown": "ขัดข้อง", "Not Checked-in": "ยังไม่ขึ้นรถ",
}


def eligibility_issues(db, driver, vehicle, task_row) -> List[str]:
    issues = []
    if driver["driver_status"] not in ("Available",):
        issues.append(f"คนขับมีสถานะ '{_STATUS_TH.get(driver['driver_status'], driver['driver_status'])}' ไม่ใช่ 'ว่าง'")
    if vehicle["status"] not in ("Available",):
        issues.append(f"รถมีสถานะ '{_STATUS_TH.get(vehicle['status'], vehicle['status'])}' ไม่ใช่ 'ว่าง'")
    lic = db.execute(
        "SELECT 1 FROM driver_licenses WHERE driver_id=? AND vehicle_type=?", (driver["id"], vehicle["vehicle_type"])
    ).fetchone()
    if not lic:
        issues.append(f"คนขับไม่มีใบอนุญาตขับรถประเภท {vehicle['vehicle_type']}")
    from_issue = zone_capability_issue(db, vehicle["vehicle_type"], task_row["from_zone"])
    if from_issue:
        issues.append(from_issue)
    if task_row["to_zone"] != task_row["from_zone"]:
        to_issue = zone_capability_issue(db, vehicle["vehicle_type"], task_row["to_zone"])
        if to_issue:
            issues.append(to_issue)
    return issues


def _do_assign(db, task_row, driver, vehicle, assigned_by_id, reason, estimated_minutes, is_reassign):
    if is_reassign:
        # Fetch every currently-active driver/vehicle BEFORE closing their rows, so
        # multi-driver tasks get everyone freed up + notified, not just the one
        # driver shown as "primary" on the task.
        displaced = db.execute(
            "SELECT * FROM task_assignments WHERE task_id=? AND assignment_status='Active'", (task_row["id"],)
        ).fetchall()
        db.execute(
            "UPDATE task_assignments SET unassigned_at=?, assignment_status='Replaced' "
            "WHERE task_id=? AND assignment_status='Active'",
            (now_iso(), task_row["id"]),
        )
        for d in displaced:
            if d["driver_id"] == driver["id"]:
                continue  # same driver being re-confirmed onto the task, nothing to free/notify
            notify(db, d["driver_id"], f"You were unassigned from task {task_row['task_code']}", task_row["id"])
            free_up_driver_and_vehicle(db, d["driver_id"], d["vehicle_id"])
        if task_row["status"] == "Assigned":
            db.execute("UPDATE tasks SET accepted_at=NULL WHERE id=?", (task_row["id"],))

    try:
        db.execute(
            "INSERT INTO task_assignments (task_id, driver_id, vehicle_id, assigned_by, assigned_at, "
            "assignment_status, reason) VALUES (?,?,?,?,?,?,?)",
            (task_row["id"], driver["id"], vehicle["id"], assigned_by_id, now_iso(), "Active", reason),
        )
    except IntegrityError:
        raise HTTPException(409, "คนขับคนนี้มีการมอบหมายที่ Active อยู่กับงานนี้อยู่แล้ว (อาจเกิดจากกดซ้ำ หรือมีคนอื่นทำพร้อมกัน)")
    eta = add_minutes_iso(estimated_minutes) if estimated_minutes else task_row["estimated_arrival"]
    db.execute(
        "UPDATE tasks SET current_driver_id=?, current_vehicle_id=?, estimated_arrival=? WHERE id=?",
        (driver["id"], vehicle["id"], eta, task_row["id"]),
    )
    db.execute("UPDATE users SET driver_status='Busy' WHERE id=?", (driver["id"],))
    db.execute("UPDATE vehicles SET status='Busy' WHERE id=?", (vehicle["id"],))
    if task_row["status"] == "Waiting":
        # Atomic claim: the WHERE clause re-checks status at write time (not
        # just from the stale in-memory task_row read earlier), so if two
        # requests race to claim the same Waiting task, only one succeeds —
        # the loser gets a clean 409 and everything it wrote in this request
        # (including the task_assignments insert above) rolls back together,
        # since get_db() only commits if the whole request completes without
        # raising.
        cur = db.execute(
            "UPDATE tasks SET status='Assigned', blocked_reason=NULL, blocked_at=NULL "
            "WHERE id=? AND status='Waiting'",
            (task_row["id"],),
        )
        if cur.rowcount == 0:
            raise HTTPException(409, "งานนี้ถูกมอบหมายไปแล้วโดยคนอื่นในระหว่างนี้ กรุณารีเฟรชหน้าจอ")
        db.execute(
            "INSERT INTO task_status_history (task_id, from_status, to_status, changed_by, changed_at) "
            "VALUES (?,?,?,?,?)",
            (task_row["id"], "Waiting", "Assigned", assigned_by_id, now_iso()),
        )

    notify(db, driver["id"], f"Task {task_row['task_code']} assigned to you", task_row["id"])
    notify(db, task_row["requester_id"], f"Task {task_row['task_code']} assigned to a driver", task_row["id"])
    if is_reassign:
        notify_admins(db, f"Task {task_row['task_code']} reassigned to {driver['employee_id']}", task_row["id"])


@app.post("/tasks/{task_id}/assign")
def assign_task(task_id: int, body: AssignBody, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if row["status"] not in ("Waiting",):
            raise HTTPException(400, "Task is not in Waiting status; use /reassign instead")
        driver = db.execute("SELECT * FROM users WHERE employee_id=? AND role='DRIVER'",
                             (body.driver_employee_id,)).fetchone()
        if not driver:
            raise HTTPException(404, "ไม่พบคนขับคนนี้")
        if not driver["checked_in_vehicle_id"]:
            raise HTTPException(400, "คนขับยังไม่ได้เช็คอินรถ ไม่สามารถมอบหมายงานได้")
        vehicle = db.execute("SELECT * FROM vehicles WHERE id=?", (driver["checked_in_vehicle_id"],)).fetchone()
        issues = eligibility_issues(db, driver, vehicle, row)
        if issues and not body.force:
            raise HTTPException(409, "เงื่อนไขไม่ผ่าน: " + "; ".join(issues) +
                                 " (ยืนยันด้วย force=true เพื่อมอบหมายทั้งที่ไม่ตรงเงื่อนไข)")
        if issues and body.force:
            audit(db, user["id"], "manual_override_assign", f"task {row['task_code']}: {'; '.join(issues)}")
        _do_assign(db, row, driver, vehicle, user["id"], body.reason, body.estimated_minutes, is_reassign=False)
        return {"ok": True, "overridden_issues": issues if issues else None}


@app.post("/tasks/{task_id}/reassign")
def reassign_task(task_id: int, body: AssignBody, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if row["status"] not in ("Assigned", "In Progress", "Pause"):
            raise HTTPException(400, f"Cannot reassign task in status {row['status']}")
        driver = db.execute("SELECT * FROM users WHERE employee_id=? AND role='DRIVER'",
                             (body.driver_employee_id,)).fetchone()
        if not driver:
            raise HTTPException(404, "ไม่พบคนขับคนนี้")
        if not driver["checked_in_vehicle_id"]:
            raise HTTPException(400, "คนขับยังไม่ได้เช็คอินรถ ไม่สามารถมอบหมายงานได้")
        vehicle = db.execute("SELECT * FROM vehicles WHERE id=?", (driver["checked_in_vehicle_id"],)).fetchone()
        issues = eligibility_issues(db, driver, vehicle, row)
        if issues and not body.force:
            raise HTTPException(409, "เงื่อนไขไม่ผ่าน: " + "; ".join(issues) +
                                 " (ยืนยันด้วย force=true เพื่อมอบหมายทั้งที่ไม่ตรงเงื่อนไข)")
        if issues and body.force:
            audit(db, user["id"], "manual_override_reassign", f"task {row['task_code']}: {'; '.join(issues)}")
        _do_assign(db, row, driver, vehicle, user["id"], body.reason, body.estimated_minutes, is_reassign=True)
        return {"ok": True, "overridden_issues": issues if issues else None}


class SelfAssignBody(BaseModel):
    battery_level: int


@app.post("/tasks/{task_id}/self-assign")
def self_assign_task(task_id: int, body: SelfAssignBody, user=Depends(require_role("DRIVER"))):
    """A driver takes a Waiting task themselves, using the vehicle they already checked in with."""
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if row["status"] != "Waiting":
            raise HTTPException(400, "Task is no longer waiting")
        driver = db.execute("SELECT * FROM users WHERE id=?", (user["id"],)).fetchone()
        if not driver["checked_in_vehicle_id"]:
            raise HTTPException(400, "Please check in with a vehicle first")
        if driver["driver_status"] != "Available":
            raise HTTPException(400, f"You are currently {driver['driver_status']}, not Available")
        vehicle = db.execute("SELECT * FROM vehicles WHERE id=?", (driver["checked_in_vehicle_id"],)).fetchone()
        issues = eligibility_issues(db, driver, vehicle, row)
        if issues:
            raise HTTPException(409, "Cannot take this task: " + "; ".join(issues))
        _do_assign(db, row, driver, vehicle, user["id"], "Self-assigned by driver", None, is_reassign=False)
        db.execute("UPDATE tasks SET accepted_at=?, accept_battery_level=? WHERE id=?",
                   (now_iso(), body.battery_level, task_id))
        db.execute("UPDATE vehicles SET battery_level=? WHERE id=?", (body.battery_level, vehicle["id"]))
        notify_admins(db, f"Driver {user['employee_id']} self-assigned task {row['task_code']}", task_id)
        return {"ok": True}


@app.post("/tasks/{task_id}/take-over")
def take_over_task(task_id: int, body: SelfAssignBody, user=Depends(require_role("DRIVER"))):
    """A driver takes over a task another driver is already working, using their
    own checked-in vehicle. Functionally a driver-initiated reassign."""
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if row["status"] not in ("Assigned", "In Progress", "Pause"):
            raise HTTPException(400, "This task cannot be taken over right now")
        if row["current_driver_id"] == user["id"]:
            raise HTTPException(400, "This is already your task")
        driver = db.execute("SELECT * FROM users WHERE id=?", (user["id"],)).fetchone()
        if not driver["checked_in_vehicle_id"]:
            raise HTTPException(400, "Please check in with a vehicle first")
        if driver["driver_status"] != "Available":
            raise HTTPException(400, f"You are currently {driver['driver_status']}, not Available")
        vehicle = db.execute("SELECT * FROM vehicles WHERE id=?", (driver["checked_in_vehicle_id"],)).fetchone()
        issues = eligibility_issues(db, driver, vehicle, row)
        if issues:
            raise HTTPException(409, "Cannot take over this task: " + "; ".join(issues))
        _do_assign(db, row, driver, vehicle, user["id"], "Taken over by another driver", None, is_reassign=True)
        db.execute("UPDATE tasks SET accepted_at=?, accept_battery_level=? WHERE id=?",
                   (now_iso(), body.battery_level, task_id))
        db.execute("UPDATE vehicles SET battery_level=? WHERE id=?", (body.battery_level, vehicle["id"]))
        notify_admins(db, f"Driver {user['employee_id']} took over task {row['task_code']}", task_id)
        return {"ok": True}


class AcceptBody(BaseModel):
    battery_level: int


@app.post("/tasks/{task_id}/accept")
def accept_task(task_id: int, body: AcceptBody, user=Depends(require_role("DRIVER"))):
    """Driver acknowledges the assignment and reports current battery %.
    Status stays 'Assigned'; accept time is recorded separately so Admin/User
    can see the driver has acknowledged but not yet started."""
    with get_db() as db:
        cur = db.execute(
            "UPDATE tasks SET accepted_at=?, accept_battery_level=? "
            "WHERE id=? AND status='Assigned' AND current_driver_id=? AND accepted_at IS NULL",
            (now_iso(), body.battery_level, task_id, user["id"]),
        )
        if cur.rowcount == 0:
            raise HTTPException(409, "Task already accepted, not assigned to you, or not in Assigned status")
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if row["current_vehicle_id"]:
            db.execute("UPDATE vehicles SET battery_level=? WHERE id=?", (body.battery_level, row["current_vehicle_id"]))
        notify(db, row["requester_id"], f"Driver accepted task {row['task_code']}", task_id)
        notify_admins(db, f"Driver {user['employee_id']} accepted task {row['task_code']}", task_id)
        return {"ok": True}


@app.post("/tasks/{task_id}/start")
def start_task(task_id: int, user=Depends(require_role("DRIVER"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if row["current_driver_id"] != user["id"]:
            raise HTTPException(403, "Not your task")
        if row["status"] != "Assigned" or row["accepted_at"] is None:
            raise HTTPException(400, "Task must be accepted before it can be started")
        transition(db, row, "In Progress", user["id"])
        db.execute("UPDATE tasks SET started_at=? WHERE id=?", (now_iso(), task_id))
        notify(db, row["requester_id"], f"Driver started task {row['task_code']}", task_id)
        return {"ok": True}


@app.post("/tasks/{task_id}/complete")
def complete_task(task_id: int, user=Depends(require_role("DRIVER"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        my_assignment = db.execute(
            "SELECT * FROM task_assignments WHERE task_id=? AND driver_id=? AND assignment_status='Active'",
            (task_id, user["id"]),
        ).fetchone()
        if not my_assignment:
            raise HTTPException(403, "Not your task")
        if row["status"] not in ("In Progress",):
            raise HTTPException(400, "Task already completed or not in progress")
        db.execute(
            "UPDATE task_assignments SET unassigned_at=?, assignment_status='Completed' WHERE id=?",
            (now_iso(), my_assignment["id"]),
        )
        free_up_driver_and_vehicle(db, user["id"], my_assignment["vehicle_id"])
        remaining = db.execute(
            "SELECT COUNT(*) n FROM task_assignments WHERE task_id=? AND assignment_status='Active'", (task_id,)
        ).fetchone()["n"]
        if remaining == 0:
            transition(db, row, "Completed", user["id"])
            db.execute("UPDATE tasks SET completed_at=? WHERE id=?", (now_iso(), task_id))
            notify(db, row["requester_id"], f"Task {row['task_code']} completed", task_id)
        else:
            # Someone else is still working; if the "primary" driver shown on the task
            # just finished, hand the primary slot to another still-active driver.
            if row["current_driver_id"] == user["id"]:
                nxt = db.execute(
                    "SELECT * FROM task_assignments WHERE task_id=? AND assignment_status='Active' "
                    "ORDER BY assigned_at LIMIT 1", (task_id,),
                ).fetchone()
                if nxt:
                    db.execute("UPDATE tasks SET current_driver_id=?, current_vehicle_id=? WHERE id=?",
                               (nxt["driver_id"], nxt["vehicle_id"], task_id))
            notify(db, row["requester_id"],
                   f"Driver {user['employee_id']} finished their part of task {row['task_code']} "
                   f"({remaining} still working)", task_id)
        return {"ok": True}


@app.post("/tasks/{task_id}/admin-complete")
def admin_complete_task(task_id: int, user=Depends(require_role("ADMIN"))):
    """Admin manually force-completes a stuck/orphaned task, bypassing the normal
    driver-must-complete-their-own-part flow."""
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if row["status"] not in ("In Progress", "Pause"):
            raise HTTPException(
                400, f"ทำเครื่องหมายเสร็จไม่ได้ในสถานะ {row['status']} "
                     f"(งานที่ยังไม่เริ่ม ใช้ 'ยกเลิก' หรือ 'มอบหมายใหม่' แทน)"
            )
        active_rows = db.execute(
            "SELECT * FROM task_assignments WHERE task_id=? AND assignment_status='Active'", (task_id,)
        ).fetchall()
        for a in active_rows:
            db.execute(
                "UPDATE task_assignments SET unassigned_at=?, assignment_status='Completed' WHERE id=?",
                (now_iso(), a["id"]),
            )
            free_up_driver_and_vehicle(db, a["driver_id"], a["vehicle_id"])
        transition(db, row, "Completed", user["id"], note="Force-completed by admin")
        db.execute("UPDATE tasks SET completed_at=? WHERE id=?", (now_iso(), task_id))
        audit(db, user["id"], "admin_complete_task", f"task {row['task_code']}")
        notify(db, row["requester_id"], f"Task {row['task_code']} marked completed by admin", task_id)
        return {"ok": True}


class JoinTaskBody(BaseModel):
    battery_level: int


@app.post("/tasks/{task_id}/join")
def join_task(task_id: int, body: JoinTaskBody, user=Depends(require_role("DRIVER"))):
    """A driver adds themself to a task someone else is already working, alongside
    them (not replacing) — for jobs that genuinely need more than one person."""
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if row["status"] not in ("Assigned", "In Progress", "Pause"):
            raise HTTPException(400, "งานนี้ยังไม่เริ่ม หรือปิดงานไปแล้ว")
        if db.execute(
            "SELECT 1 FROM task_assignments WHERE task_id=? AND driver_id=? AND assignment_status='Active'",
            (task_id, user["id"]),
        ).fetchone():
            raise HTTPException(400, "คุณอยู่ในงานนี้อยู่แล้ว")
        driver = db.execute("SELECT * FROM users WHERE id=?", (user["id"],)).fetchone()
        if not driver["checked_in_vehicle_id"]:
            raise HTTPException(400, "กรุณาเช็คอินรถก่อน")
        if driver["driver_status"] != "Available":
            raise HTTPException(400, f"คุณมีสถานะ {_STATUS_TH.get(driver['driver_status'], driver['driver_status'])} ไม่ใช่ 'ว่าง'")
        vehicle = db.execute("SELECT * FROM vehicles WHERE id=?", (driver["checked_in_vehicle_id"],)).fetchone()
        issues = eligibility_issues(db, driver, vehicle, row)
        if issues:
            raise HTTPException(409, "เข้าร่วมไม่ได้: " + "; ".join(issues))
        try:
            db.execute(
                "INSERT INTO task_assignments (task_id, driver_id, vehicle_id, assigned_by, assigned_at, "
                "assignment_status, reason) VALUES (?,?,?,?,?,?,?)",
                (task_id, driver["id"], vehicle["id"], user["id"], now_iso(), "Active", "Joined by driver"),
            )
        except IntegrityError:
            raise HTTPException(409, "คุณเพิ่งเข้าร่วมงานนี้ไปแล้ว (อาจเกิดจากกดซ้ำ)")
        db.execute("UPDATE users SET driver_status='Busy' WHERE id=?", (driver["id"],))
        db.execute("UPDATE vehicles SET status='Busy', battery_level=? WHERE id=?", (body.battery_level, vehicle["id"]))
        notify(db, row["requester_id"], f"Driver {user['employee_id']} joined task {row['task_code']}", task_id)
        notify_admins(db, f"Driver {user['employee_id']} joined task {row['task_code']}", task_id)
        return {"ok": True}


class AddDriverBody(BaseModel):
    driver_employee_id: str


@app.post("/tasks/{task_id}/add-driver")
def admin_add_driver(task_id: int, body: AddDriverBody, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if row["status"] not in ("Assigned", "In Progress", "Pause"):
            raise HTTPException(400, "เพิ่มคนขับได้เฉพาะงานที่กำลังดำเนินการเท่านั้น")
        driver = db.execute(
            "SELECT * FROM users WHERE employee_id=? AND role='DRIVER'", (body.driver_employee_id,)
        ).fetchone()
        if not driver:
            raise HTTPException(404, "ไม่พบคนขับคนนี้")
        if db.execute(
            "SELECT 1 FROM task_assignments WHERE task_id=? AND driver_id=? AND assignment_status='Active'",
            (task_id, driver["id"]),
        ).fetchone():
            raise HTTPException(400, "คนขับคนนี้อยู่ในงานนี้อยู่แล้ว")
        if not driver["checked_in_vehicle_id"]:
            raise HTTPException(400, "คนขับยังไม่ได้เช็คอินรถ")
        vehicle = db.execute("SELECT * FROM vehicles WHERE id=?", (driver["checked_in_vehicle_id"],)).fetchone()
        issues = eligibility_issues(db, driver, vehicle, row)
        if issues:
            raise HTTPException(409, "เพิ่มไม่ได้: " + "; ".join(issues))
        try:
            db.execute(
                "INSERT INTO task_assignments (task_id, driver_id, vehicle_id, assigned_by, assigned_at, "
                "assignment_status, reason) VALUES (?,?,?,?,?,?,?)",
                (task_id, driver["id"], vehicle["id"], user["id"], now_iso(), "Active", "Added by admin"),
            )
        except IntegrityError:
            raise HTTPException(409, "คนขับคนนี้เพิ่งถูกเพิ่มเข้างานนี้ไปแล้ว (อาจเกิดจากกดซ้ำ)")
        db.execute("UPDATE users SET driver_status='Busy' WHERE id=?", (driver["id"],))
        db.execute("UPDATE vehicles SET status='Busy' WHERE id=?", (vehicle["id"],))
        notify(db, driver["id"], f"You were added to task {row['task_code']}", task_id)
        notify(db, row["requester_id"], f"Driver {driver['employee_id']} added to task {row['task_code']}", task_id)
        return {"ok": True}


@app.post("/tasks/{task_id}/leave")
def leave_task(task_id: int, user=Depends(require_role("DRIVER"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        my_assignment = db.execute(
            "SELECT * FROM task_assignments WHERE task_id=? AND driver_id=? AND assignment_status='Active'",
            (task_id, user["id"]),
        ).fetchone()
        if not my_assignment:
            raise HTTPException(400, "คุณไม่ได้อยู่ในงานนี้")
        active_count = db.execute(
            "SELECT COUNT(*) n FROM task_assignments WHERE task_id=? AND assignment_status='Active'", (task_id,)
        ).fetchone()["n"]
        if active_count <= 1 and row["status"] == "In Progress":
            raise HTTPException(
                400, "ไม่สามารถออกจากงานได้ เพราะเป็นคนเดียวที่เหลืออยู่ขณะกำลังทำงาน "
                     "กรุณาให้คนอื่นเข้าร่วมหรือติดต่อ Admin ก่อน"
            )
        db.execute(
            "UPDATE task_assignments SET unassigned_at=?, assignment_status='Left' WHERE id=?",
            (now_iso(), my_assignment["id"]),
        )
        free_up_driver_and_vehicle(db, user["id"], my_assignment["vehicle_id"])
        if active_count <= 1:
            # Last driver, task hadn't started yet -> send it back to the pool.
            db.execute(
                "UPDATE tasks SET current_driver_id=NULL, current_vehicle_id=NULL, accepted_at=NULL WHERE id=?",
                (task_id,),
            )
            transition(db, row, "Waiting", user["id"], note="Last driver left before starting")
            notify_admins(db, f"งาน {row['task_code']} ไม่มีคนขับแล้ว กลับไปรอมอบหมายใหม่", task_id)
            return {"ok": True}
        if row["current_driver_id"] == user["id"]:
            nxt = db.execute(
                "SELECT * FROM task_assignments WHERE task_id=? AND assignment_status='Active' "
                "ORDER BY assigned_at LIMIT 1", (task_id,),
            ).fetchone()
            if nxt:
                db.execute("UPDATE tasks SET current_driver_id=?, current_vehicle_id=? WHERE id=?",
                           (nxt["driver_id"], nxt["vehicle_id"], task_id))
        notify(db, row["requester_id"], f"Driver {user['employee_id']} left task {row['task_code']}", task_id)
        notify_admins(db, f"Driver {user['employee_id']} left task {row['task_code']}", task_id)
        return {"ok": True}


class RatingBody(BaseModel):
    rating: int
    comment: Optional[str] = None


@app.post("/tasks/{task_id}/rating")
def rate_task(task_id: int, body: RatingBody, user=Depends(require_role("USER"))):
    if body.rating < 1 or body.rating > 5:
        raise HTTPException(400, "คะแนนต้องอยู่ระหว่าง 1-5 ดาว")
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if row["requester_id"] != user["id"]:
            raise HTTPException(403, "Not your task")
        if row["status"] != "Completed":
            raise HTTPException(400, "ให้คะแนนได้เฉพาะงานที่เสร็จสิ้นแล้ว")
        db.execute("UPDATE tasks SET rating=?, rating_comment=? WHERE id=?", (body.rating, body.comment, task_id))
        if row["current_driver_id"]:
            notify(db, row["current_driver_id"], f"งาน {row['task_code']} ได้รับคะแนน {body.rating} ดาว", task_id)
        audit(db, user["id"], "rate_task", f"task {row['task_code']}: {body.rating} stars")
        return {"ok": True}


class PauseBody(BaseModel):
    category: str  # 'user' or 'vehicle'
    reason: str


@app.post("/tasks/{task_id}/pause")
def pause_task(task_id: int, body: PauseBody, user=Depends(require_role("DRIVER"))):
    if body.category not in ("user", "vehicle"):
        raise HTTPException(400, "category must be 'user' or 'vehicle'")
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        my_assignment = db.execute(
            "SELECT * FROM task_assignments WHERE task_id=? AND driver_id=? AND assignment_status='Active'",
            (task_id, user["id"]),
        ).fetchone()
        if not my_assignment:
            raise HTTPException(403, "Not your task")
        note = f"[{body.category}] {body.reason}"

        # A problem on the requester's side (location/goods not ready) blocks
        # everyone regardless of how many vehicles are on the job — pause the
        # whole task.
        if body.category == "user":
            transition(db, row, "Pause", user["id"], note=note)
            notify(db, row["requester_id"], f"Task {row['task_code']} paused: {body.reason}", task_id)
            notify_admins(db, f"Task {row['task_code']} paused (user): {body.reason}", task_id)
            return {"ok": True}

        # A vehicle problem only stops the reporting driver. If someone else
        # is still actively working the task, the task itself stays "In
        # Progress" — only pause it if this was the only driver left.
        others = db.execute(
            "SELECT COUNT(*) n FROM task_assignments WHERE task_id=? AND assignment_status='Active' AND driver_id!=?",
            (task_id, user["id"]),
        ).fetchone()["n"]
        if others == 0:
            transition(db, row, "Pause", user["id"], note=note)
            notify(db, row["requester_id"], f"Task {row['task_code']} paused: {body.reason}", task_id)
            notify_admins(db, f"Task {row['task_code']} paused (vehicle): {body.reason}", task_id)
            return {"ok": True}

        # Others remain active: step this driver back (like /leave) and flag
        # their vehicle as broken down, but leave the task itself running.
        db.execute(
            "UPDATE task_assignments SET unassigned_at=?, assignment_status='Left' WHERE id=?",
            (now_iso(), my_assignment["id"]),
        )
        db.execute("UPDATE users SET driver_status='Available' WHERE id=?", (user["id"],))
        db.execute("UPDATE vehicles SET status='Breakdown' WHERE id=?", (my_assignment["vehicle_id"],))
        vehicle_row = db.execute("SELECT code FROM vehicles WHERE id=?", (my_assignment["vehicle_id"],)).fetchone()
        db.execute(
            "INSERT INTO breakdowns (target_type, target_id, target_label, description, reported_by, status, created_at) "
            "VALUES ('vehicle',?,?,?,?,?,?)",
            (my_assignment["vehicle_id"], vehicle_row["code"] if vehicle_row else "", body.reason, user["id"], "Open", now_iso()),
        )
        if row["current_driver_id"] == user["id"]:
            nxt = db.execute(
                "SELECT * FROM task_assignments WHERE task_id=? AND assignment_status='Active' "
                "ORDER BY assigned_at LIMIT 1", (task_id,),
            ).fetchone()
            if nxt:
                db.execute("UPDATE tasks SET current_driver_id=?, current_vehicle_id=? WHERE id=?",
                           (nxt["driver_id"], nxt["vehicle_id"], task_id))
        notify(db, row["requester_id"],
               f"Driver {user['employee_id']} had a vehicle problem on task {row['task_code']} "
               f"({body.reason}) — task continues with the remaining driver(s)", task_id)
        notify_admins(db, f"Vehicle problem reported on task {row['task_code']} by {user['employee_id']}: "
                          f"{body.reason} (task still In Progress, {others} driver(s) remain)", task_id)
        return {"ok": True}


@app.post("/tasks/{task_id}/resume")
def resume_task(task_id: int, user=Depends(require_role("DRIVER"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        my_assignment = db.execute(
            "SELECT 1 FROM task_assignments WHERE task_id=? AND driver_id=? AND assignment_status='Active'",
            (task_id, user["id"]),
        ).fetchone()
        if not my_assignment:
            raise HTTPException(403, "Not your task")
        transition(db, row, "In Progress", user["id"])
        return {"ok": True}


# ---------------------------------------------------------------------------
# Task chat (scoped to task participants: requester, any driver ever assigned, admins)
# ---------------------------------------------------------------------------

class ChatMessageBody(BaseModel):
    message: str


@app.get("/tasks/{task_id}/messages")
def get_task_messages(task_id: int, user=Depends(current_user)):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if not task_chat_allowed(db, row, user):
            raise HTTPException(403, "You are not part of this task's conversation")
        rows = db.execute(
            "SELECT task_messages.*, users.employee_id AS sender_employee_id, users.role AS sender_role "
            "FROM task_messages JOIN users ON users.id = task_messages.sender_id "
            "WHERE task_id=? ORDER BY created_at", (task_id,)
        ).fetchall()
        return [dict(r) for r in rows]


@app.post("/tasks/{task_id}/messages")
def post_task_message(task_id: int, body: ChatMessageBody, user=Depends(current_user)):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Task not found")
        if not task_chat_allowed(db, row, user):
            raise HTTPException(403, "You are not part of this task's conversation")
        if not body.message.strip():
            raise HTTPException(400, "Message cannot be empty")
        db.execute(
            "INSERT INTO task_messages (task_id, sender_id, message, created_at) VALUES (?,?,?,?)",
            (task_id, user["id"], body.message.strip(), now_iso()),
        )
        others = set()
        if row["requester_id"] != user["id"]:
            others.add(row["requester_id"])
        active_drivers = db.execute(
            "SELECT driver_id FROM task_assignments WHERE task_id=? AND assignment_status='Active'", (task_id,)
        ).fetchall()
        for r in active_drivers:
            if r["driver_id"] != user["id"]:
                others.add(r["driver_id"])
        for uid in others:
            notify(db, uid, f"New message on task {row['task_code']}", task_id, kind="chat")
        return {"ok": True}


# ---------------------------------------------------------------------------
# Team chat (Admin <-> Driver only, not visible to USER) with @mentions
# ---------------------------------------------------------------------------

class TeamChatBody(BaseModel):
    message: str


@app.get("/team-chat/participants")
def team_chat_participants(user=Depends(require_role("ADMIN", "DRIVER"))):
    with get_db() as db:
        rows = db.execute(
            "SELECT employee_id, full_name, role FROM users WHERE role IN ('ADMIN','DRIVER') "
            "AND deleted_at IS NULL ORDER BY role, employee_id"
        ).fetchall()
        return [dict(r) for r in rows]


@app.get("/team-chat/messages")
def get_team_chat_messages(user=Depends(require_role("ADMIN", "DRIVER"))):
    with get_db() as db:
        rows = db.execute(
            "SELECT team_chat_messages.*, users.employee_id AS sender_employee_id, users.full_name AS sender_name, "
            "users.role AS sender_role FROM team_chat_messages JOIN users ON users.id = team_chat_messages.sender_id "
            "ORDER BY team_chat_messages.created_at LIMIT 200"
        ).fetchall()
        return [dict(r) for r in rows]


@app.post("/team-chat/messages")
def post_team_chat_message(body: TeamChatBody, user=Depends(require_role("ADMIN", "DRIVER"))):
    message = body.message.strip()
    if not message:
        raise HTTPException(400, "Message cannot be empty")
    with get_db() as db:
        db.execute(
            "INSERT INTO team_chat_messages (sender_id, message, created_at) VALUES (?,?,?)",
            (user["id"], message, now_iso()),
        )
        # Parse @employee_id mentions and notify each matched, real ADMIN/DRIVER user.
        mentioned_ids = set(re.findall(r"@(\w+)", message))
        for emp_id in mentioned_ids:
            target = db.execute(
                "SELECT id FROM users WHERE employee_id=? AND role IN ('ADMIN','DRIVER') AND deleted_at IS NULL", (emp_id,)
            ).fetchone()
            if target and target["id"] != user["id"]:
                notify(db, target["id"], f"{user['employee_id']} mentioned you in Team Chat", kind="team_chat")
        return {"ok": True}


# ---------------------------------------------------------------------------
# Drivers / Vehicles / Check-in / Dispatch suggestion
# ---------------------------------------------------------------------------

@app.get("/drivers/availability-summary")
def drivers_availability_summary(user=Depends(current_user)):
    """Aggregate counts only (no names/details) — safe to expose to USER and
    DRIVER, not just ADMIN, so a requester can see at a glance whether a
    driver is likely to pick up their task soon."""
    with get_db() as db:
        rows = db.execute(
            "SELECT driver_status, COUNT(*) n FROM users "
            "WHERE role='DRIVER' AND deleted_at IS NULL GROUP BY driver_status"
        ).fetchall()
        counts = {r["driver_status"]: r["n"] for r in rows}
        available = counts.get("Available", 0)
        busy = counts.get("Busy", 0)
        breakdown = counts.get("Breakdown", 0)
        not_checked_in = counts.get("Not Checked-in", 0)
        return {
            "available": available, "busy": busy, "breakdown": breakdown,
            "not_checked_in": not_checked_in,
            "total": available + busy + breakdown + not_checked_in,
        }


@app.get("/drivers")
def list_drivers(user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        concat_fn = "STRING_AGG(driver_licenses.vehicle_type, ',')" if IS_POSTGRES else "GROUP_CONCAT(driver_licenses.vehicle_type)"
        # expiry_date is stored as YYYY-MM-DD TEXT. SQLite compares it fine with
        # date('now'), but Postgres refuses text-vs-date ("operator does not
        # exist: text < date"). Since ISO YYYY-MM-DD sorts correctly as plain
        # strings, compare against today's date rendered AS TEXT on Postgres.
        today_expr = "to_char(CURRENT_DATE,'YYYY-MM-DD')" if IS_POSTGRES else "date('now')"
        rows = db.execute(
            "SELECT users.id, users.employee_id, users.full_name, users.driver_status, users.checked_in_vehicle_id, "
            f"{concat_fn} AS license_types, "
            f"SUM(CASE WHEN driver_licenses.expiry_date IS NOT NULL AND driver_licenses.expiry_date < {today_expr} "
            "THEN 1 ELSE 0 END) AS expired_count "
            "FROM users LEFT JOIN driver_licenses ON driver_licenses.driver_id = users.id "
            "WHERE users.role='DRIVER' AND users.deleted_at IS NULL GROUP BY users.id"
        ).fetchall()
        return [dict(r) for r in rows]


class VehicleTypeCreate(BaseModel):
    type_key: str
    type_name_th: str
    description: Optional[str] = None


@app.get("/vehicle-types")
def list_vehicle_types(user=Depends(current_user)):
    with get_db() as db:
        return [dict(r) for r in db.execute("SELECT * FROM vehicle_types ORDER BY id").fetchall()]


@app.post("/vehicle-types")
def create_vehicle_type(body: VehicleTypeCreate, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        key = body.type_key.strip().upper()
        if db.execute("SELECT 1 FROM vehicle_types WHERE type_key=?", (key,)).fetchone():
            raise HTTPException(400, "ประเภทรถนี้มีอยู่แล้ว")
        db.execute("INSERT INTO vehicle_types (type_key, type_name_th, description) VALUES (?,?,?)",
                   (key, body.type_name_th, body.description))
        audit(db, user["id"], "create_vehicle_type", key)
        return {"ok": True}


@app.delete("/vehicle-types/{type_key}")
def delete_vehicle_type(type_key: str, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        if db.execute("SELECT 1 FROM vehicles WHERE vehicle_type=?", (type_key,)).fetchone():
            raise HTTPException(400, "ลบไม่ได้ เพราะยังมีรถที่ใช้ประเภทนี้อยู่")
        db.execute("DELETE FROM vehicle_types WHERE type_key=?", (type_key,))
        db.execute("DELETE FROM vehicle_type_zones WHERE vehicle_type=?", (type_key,))
        audit(db, user["id"], "delete_vehicle_type", type_key)
        return {"ok": True}


@app.get("/vehicles")
def list_vehicles(user=Depends(current_user)):
    with get_db() as db:
        rows = db.execute("SELECT * FROM vehicles").fetchall()
        return [dict(r) for r in rows]


class VehicleCreate(BaseModel):
    code: str
    vehicle_type: str
    battery_level: int = 100


@app.post("/vehicles")
def create_vehicle(body: VehicleCreate, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        if db.execute("SELECT 1 FROM vehicles WHERE code=?", (body.code,)).fetchone():
            raise HTTPException(400, "รหัสรถนี้มีอยู่แล้ว")
        if not db.execute("SELECT 1 FROM vehicle_types WHERE type_key=?", (body.vehicle_type,)).fetchone():
            raise HTTPException(400, "ไม่พบประเภทรถนี้ กรุณาเพิ่มประเภทรถก่อน")
        db.execute(
            "INSERT INTO vehicles (code, vehicle_type, status, battery_level) VALUES (?,?,'Available',?)",
            (body.code, body.vehicle_type, body.battery_level),
        )
        audit(db, user["id"], "create_vehicle", body.code)
        return {"ok": True}


@app.delete("/vehicles/{vehicle_id}")
def delete_vehicle(vehicle_id: int, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Vehicle not found")
        active_task = db.execute(
            "SELECT 1 FROM tasks WHERE current_vehicle_id=? AND status NOT IN ('Completed','Cancelled')",
            (vehicle_id,),
        ).fetchone()
        if active_task:
            raise HTTPException(400, "ลบไม่ได้ เพราะรถคันนี้กำลังถูกใช้งานในงานที่ยังไม่เสร็จ")
        if db.execute("SELECT 1 FROM users WHERE checked_in_vehicle_id=?", (vehicle_id,)).fetchone():
            raise HTTPException(400, "ลบไม่ได้ เพราะมีคนขับเช็คอินรถคันนี้อยู่")
        db.execute("DELETE FROM vehicles WHERE id=?", (vehicle_id,))
        audit(db, user["id"], "delete_vehicle", row["code"])
        return {"ok": True}


class VehicleUpdateBody(BaseModel):
    vehicle_type: Optional[str] = None


@app.put("/vehicles/{vehicle_id}")
def update_vehicle(vehicle_id: int, body: VehicleUpdateBody, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
        if not row:
            raise HTTPException(404, "Vehicle not found")
        fields = body.dict(exclude_unset=True)
        if not fields:
            return {"ok": True}
        if "vehicle_type" in fields and not db.execute(
            "SELECT 1 FROM vehicle_types WHERE type_key=?", (fields["vehicle_type"],)
        ).fetchone():
            raise HTTPException(400, "ไม่พบประเภทรถนี้")
        set_clause = ", ".join(f"{k}=?" for k in fields.keys())
        db.execute(f"UPDATE vehicles SET {set_clause} WHERE id=?", list(fields.values()) + [vehicle_id])
        audit(db, user["id"], "update_vehicle", f"{row['code']}: {fields}")
        return {"ok": True}


@app.get("/vehicles/available-for-checkin")
def vehicles_available_for_checkin(user=Depends(require_role("DRIVER"))):
    """Vehicles nobody else is currently checked into (excludes Breakdown too)."""
    with get_db() as db:
        rows = db.execute(
            "SELECT * FROM vehicles WHERE status != 'Breakdown' AND id NOT IN "
            "(SELECT checked_in_vehicle_id FROM users WHERE role='DRIVER' "
            "AND checked_in_vehicle_id IS NOT NULL AND id != ?) ORDER BY code",
            (user["id"],),
        ).fetchall()
        return [dict(r) for r in rows]


class DriverStatusBody(BaseModel):
    status: str


@app.post("/drivers/status")
def set_driver_status(body: DriverStatusBody, user=Depends(require_role("DRIVER"))):
    if body.status not in ("Not Checked-in", "Available", "Busy", "Pause", "Breakdown"):
        raise HTTPException(400, "Invalid status")
    with get_db() as db:
        db.execute("UPDATE users SET driver_status=? WHERE id=?", (body.status, user["id"]))
        if body.status == "Breakdown":
            notify_admins(db, f"Driver {user['employee_id']} set status to Breakdown")
        return {"ok": True}


class CheckInBody(BaseModel):
    vehicle_code: str
    battery_level: int


@app.post("/vehicles/check-in")
def vehicle_check_in(body: CheckInBody, user=Depends(require_role("DRIVER"))):
    with get_db() as db:
        v = db.execute("SELECT * FROM vehicles WHERE code=?", (body.vehicle_code,)).fetchone()
        if not v:
            raise HTTPException(404, "Vehicle not found")
        if v["status"] == "Breakdown":
            raise HTTPException(400, "Vehicle is marked as Breakdown; cannot check in")
        db.execute("UPDATE vehicles SET battery_level=? WHERE id=?", (body.battery_level, v["id"]))
        db.execute("UPDATE users SET checked_in_vehicle_id=? WHERE id=?", (v["id"], user["id"]))
        cur = db.execute(
            "SELECT driver_status FROM users WHERE id=?", (user["id"],)
        ).fetchone()
        if cur["driver_status"] == "Not Checked-in":
            db.execute("UPDATE users SET driver_status='Available' WHERE id=?", (user["id"],))
        audit(db, user["id"], "vehicle_check_in", f"{body.vehicle_code} battery={body.battery_level}%")
        return {"ok": True}


@app.post("/vehicles/check-out")
def vehicle_check_out(user=Depends(require_role("DRIVER"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM users WHERE id=?", (user["id"],)).fetchone()
        if row["driver_status"] == "Busy":
            raise HTTPException(400, "ไม่สามารถเช็คเอาท์ได้ขณะกำลังมีงานอยู่ กรุณาทำงานให้เสร็จหรือหยุดงานก่อน")
        if not row["checked_in_vehicle_id"]:
            raise HTTPException(400, "คุณยังไม่ได้เช็คอินรถ")
        v = db.execute("SELECT code FROM vehicles WHERE id=?", (row["checked_in_vehicle_id"],)).fetchone()
        db.execute("UPDATE users SET checked_in_vehicle_id=NULL, driver_status='Not Checked-in' WHERE id=?", (user["id"],))
        audit(db, user["id"], "vehicle_check_out", v["code"] if v else "")
        return {"ok": True}


@app.get("/dispatch/suggest/{task_id}")
def suggest_dispatch(task_id: int, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        task = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not task:
            raise HTTPException(404, "Task not found")
        vehicles = db.execute("SELECT * FROM vehicles WHERE status='Available'").fetchall()
        suggestions = []
        for v in vehicles:
            drivers = db.execute(
                "SELECT DISTINCT users.* FROM users JOIN driver_licenses ON driver_licenses.driver_id = users.id "
                "WHERE users.role='DRIVER' AND users.driver_status='Available' AND users.deleted_at IS NULL "
                "AND driver_licenses.vehicle_type=?",
                (v["vehicle_type"],),
            ).fetchall()
            for d in drivers:
                suggestions.append({
                    "driver_employee_id": d["employee_id"],
                    "vehicle_code": v["code"],
                    "reason": f"Driver authorized for {v['vehicle_type']}, both available",
                })
        return {"suggestions": suggestions[:5]}


# ---------------------------------------------------------------------------
# Driver stats (for the History dashboard)
# ---------------------------------------------------------------------------

@app.get("/drivers/me/stats")
def my_driver_stats(period: str = "day", date: Optional[str] = None, user=Depends(require_role("DRIVER"))):
    if period not in ("day", "month"):
        raise HTTPException(400, "period must be 'day' or 'month'")
    if not date:
        date = datetime.utcnow().strftime("%Y-%m-%d" if period == "day" else "%Y-%m")
    try:
        if period == "day":
            start = datetime.strptime(date, "%Y-%m-%d")
            end = start + timedelta(days=1)
        else:
            start = datetime.strptime(date, "%Y-%m")
            end = (start + timedelta(days=32)).replace(day=1)
    except ValueError:
        raise HTTPException(400, "Invalid date format")
    start_iso = start.isoformat(timespec="seconds") + "Z"
    end_iso = end.isoformat(timespec="seconds") + "Z"
    with get_db() as db:
        # Multi-driver aware: credit = task's pallet_qty split evenly among everyone
        # who actually finished their part of that task (assignment_status='Completed').
        my_rows = db.execute(
            "SELECT task_assignments.task_id, tasks.pallet_qty FROM task_assignments "
            "JOIN tasks ON tasks.id = task_assignments.task_id "
            "WHERE task_assignments.driver_id=? AND task_assignments.assignment_status='Completed' "
            "AND tasks.status='Completed' AND tasks.completed_at>=? AND tasks.completed_at<?",
            (user["id"], start_iso, end_iso),
        ).fetchall()
        task_count = len(my_rows)
        pallet_total = 0.0
        for r in my_rows:
            participants = db.execute(
                "SELECT COUNT(*) n FROM task_assignments WHERE task_id=? AND assignment_status='Completed'",
                (r["task_id"],),
            ).fetchone()["n"]
            pallet_total += r["pallet_qty"] / max(1, participants)
        return {"period": period, "date": date, "task_count": task_count,
                "pallet_total": round(pallet_total, 1)}


# ---------------------------------------------------------------------------
# Breakdowns
# ---------------------------------------------------------------------------

class BreakdownBody(BaseModel):
    target_type: str
    vehicle_code: Optional[str] = None
    description: Optional[str] = None


@app.post("/breakdowns")
def report_breakdown(body: BreakdownBody, user=Depends(require_role("DRIVER"))):
    with get_db() as db:
        if body.target_type == "driver":
            target_id = user["id"]
            label = user["employee_id"]
            db.execute("UPDATE users SET driver_status='Breakdown' WHERE id=?", (user["id"],))
        elif body.target_type == "vehicle":
            if not body.vehicle_code:
                raise HTTPException(400, "vehicle_code required")
            v = db.execute("SELECT * FROM vehicles WHERE code=?", (body.vehicle_code,)).fetchone()
            if not v:
                raise HTTPException(404, "Vehicle not found")
            target_id = v["id"]
            label = v["code"]
            db.execute("UPDATE vehicles SET status='Breakdown' WHERE id=?", (v["id"],))
        else:
            raise HTTPException(400, "target_type must be 'driver' or 'vehicle'")
        db.execute(
            "INSERT INTO breakdowns (target_type, target_id, target_label, description, reported_by, status, created_at) "
            "VALUES (?,?,?,?,?,?,?)",
            (body.target_type, target_id, label, body.description, user["id"], "Open", now_iso()),
        )
        notify_admins(db, f"Breakdown reported: {body.target_type} {label} by {user['employee_id']}")
        return {"ok": True}


@app.get("/breakdowns")
def list_breakdowns(user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        rows = db.execute("SELECT * FROM breakdowns ORDER BY created_at DESC").fetchall()
        return [dict(r) for r in rows]


class ResolveBreakdownBody(BaseModel):
    resolution_note: Optional[str] = None


@app.post("/breakdowns/{breakdown_id}/resolve")
def resolve_breakdown(breakdown_id: int, body: ResolveBreakdownBody = ResolveBreakdownBody(), user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        b = db.execute("SELECT * FROM breakdowns WHERE id=?", (breakdown_id,)).fetchone()
        if not b:
            raise HTTPException(404, "Breakdown not found")
        note = (body.resolution_note or "").strip() or None
        db.execute(
            "UPDATE breakdowns SET status='Resolved', resolved_at=?, resolution_note=? WHERE id=?",
            (now_iso(), note, breakdown_id),
        )
        if b["target_type"] == "driver":
            db.execute("UPDATE users SET driver_status='Available' WHERE id=?", (b["target_id"],))
        else:
            db.execute("UPDATE vehicles SET status='Available' WHERE id=?", (b["target_id"],))
        msg = f"ปัญหาที่คุณแจ้ง ({b['target_label']}) ได้รับการแก้ไขแล้ว"
        if note:
            msg += f": {note}"
        notify(db, b["reported_by"], msg)
        audit(db, user["id"], "resolve_breakdown", f"{b['target_label']}" + (f": {note}" if note else ""))
        return {"ok": True}


# ---------------------------------------------------------------------------
# Notifications
# ---------------------------------------------------------------------------

class BroadcastBody(BaseModel):
    role: str
    message: str


@app.get("/notifications")
def get_notifications(user=Depends(current_user)):
    with get_db() as db:
        rows = db.execute(
            "SELECT * FROM notifications WHERE user_id=? ORDER BY created_at DESC LIMIT 50", (user["id"],)
        ).fetchall()
        return [dict(r) for r in rows]


@app.post("/notifications/{notif_id}/read")
def read_notification(notif_id: int, user=Depends(current_user)):
    with get_db() as db:
        db.execute("UPDATE notifications SET is_read=1 WHERE id=? AND user_id=?", (notif_id, user["id"]))
        return {"ok": True}


@app.post("/notifications/broadcast")
def broadcast_notification(body: BroadcastBody, user=Depends(require_role("ADMIN"))):
    message = body.message.strip()
    if not message:
        raise HTTPException(400, "กรุณากรอกข้อความประกาศ")
    with get_db() as db:
        if body.role == "ALL":
            targets = db.execute("SELECT id FROM users WHERE role IN ('USER','DRIVER') AND deleted_at IS NULL").fetchall()
        else:
            targets = db.execute("SELECT id FROM users WHERE role=? AND deleted_at IS NULL", (body.role,)).fetchall()
        for t in targets:
            notify(db, t["id"], f"[Admin] {message}")
        audit(db, user["id"], "broadcast", f"{body.role}: {message}")
        return {"ok": True, "sent_to": len(targets)}


# ---------------------------------------------------------------------------
# Web Push subscriptions
# ---------------------------------------------------------------------------

@app.get("/push/vapid-public-key")
def push_vapid_public_key():
    key = get_vapid_public_key_b64()
    if not key:
        raise HTTPException(503, "Web Push ยังไม่พร้อมใช้งานบนเซิร์ฟเวอร์นี้ (ต้องรันผ่าน HTTPS)")
    return {"publicKey": key}


class PushSubscribeBody(BaseModel):
    endpoint: str
    keys: dict


@app.post("/push/subscribe")
def push_subscribe(body: PushSubscribeBody, user=Depends(current_user)):
    p256dh = body.keys.get("p256dh")
    auth = body.keys.get("auth")
    if not p256dh or not auth:
        raise HTTPException(400, "Invalid subscription payload")
    with get_db() as db:
        db.execute(
            "INSERT INTO push_subscriptions (user_id, endpoint, p256dh, auth, created_at) VALUES (?,?,?,?,?) "
            "ON CONFLICT(endpoint) DO UPDATE SET user_id=excluded.user_id, p256dh=excluded.p256dh, auth=excluded.auth",
            (user["id"], body.endpoint, p256dh, auth, now_iso()),
        )
        return {"ok": True}


class PushUnsubscribeBody(BaseModel):
    endpoint: str


@app.post("/push/unsubscribe")
def push_unsubscribe(body: PushUnsubscribeBody, user=Depends(current_user)):
    with get_db() as db:
        db.execute("DELETE FROM push_subscriptions WHERE endpoint=? AND user_id=?", (body.endpoint, user["id"]))
        return {"ok": True}


@app.get("/push/status")
def push_status(user=Depends(current_user)):
    with get_db() as db:
        count = db.execute(
            "SELECT COUNT(*) n FROM push_subscriptions WHERE user_id=?", (user["id"],)
        ).fetchone()["n"]
        return {"subscribed": count > 0, "available": WEB_PUSH_AVAILABLE and bool(get_vapid_public_key_b64())}


@app.post("/push/test")
def push_test(user=Depends(current_user)):
    """Called right after subscribing, so the person gets an immediate, honest
    answer — 'yes this actually works' or a specific reason why not — instead
    of just seeing a toggle flip to on with no proof anything was delivered."""
    with get_db() as db:
        result = send_web_push_to_user(
            db, user["id"], "ทดสอบการแจ้งเตือน",
            "ถ้าคุณเห็นข้อความนี้ แปลว่าการแจ้งเตือนใช้งานได้แล้ว!",
        )
    if result["attempted"] == 0:
        raise HTTPException(400, result["reason"])
    if result["sent"] == 0:
        raise HTTPException(502, result["reason"])
    return {"ok": True, "sent": result["sent"], "attempted": result["attempted"]}


# ---------------------------------------------------------------------------
# Audit log (Admin)
# ---------------------------------------------------------------------------

@app.get("/audit-logs")
def get_audit_logs(user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        rows = db.execute(
            "SELECT audit_logs.*, users.employee_id AS actor FROM audit_logs "
            "LEFT JOIN users ON users.id = audit_logs.actor_id ORDER BY audit_logs.created_at DESC LIMIT 100"
        ).fetchall()
        return [dict(r) for r in rows]


# ---------------------------------------------------------------------------
# Dashboard summary (Admin)
# ---------------------------------------------------------------------------

@app.get("/dashboard")
def dashboard(user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        check_waiting_too_long(db)
        task_counts = {}
        for s in ["Waiting", "Assigned", "In Progress", "Pause", "Completed", "Cancelled"]:
            c = db.execute("SELECT COUNT(*) n FROM tasks WHERE status=?", (s,)).fetchone()
            task_counts[s] = c["n"]
        cutoff = (datetime.utcnow() - timedelta(minutes=WAITING_TOO_LONG_MINUTES)).isoformat(timespec="seconds") + "Z"
        delayed = db.execute("SELECT COUNT(*) n FROM tasks WHERE status='Waiting' AND created_at<?", (cutoff,)).fetchone()["n"]
        task_counts["Delayed"] = delayed

        avail_drivers = db.execute("SELECT COUNT(*) n FROM users WHERE role='DRIVER' AND driver_status='Available' AND deleted_at IS NULL").fetchone()["n"]
        total_drivers = db.execute("SELECT COUNT(*) n FROM users WHERE role='DRIVER' AND deleted_at IS NULL").fetchone()["n"]
        busy_drivers = db.execute("SELECT COUNT(*) n FROM users WHERE role='DRIVER' AND driver_status='Busy' AND deleted_at IS NULL").fetchone()["n"]
        breakdown_drivers = db.execute("SELECT COUNT(*) n FROM users WHERE role='DRIVER' AND driver_status='Breakdown' AND deleted_at IS NULL").fetchone()["n"]

        avail_vehicles = db.execute("SELECT COUNT(*) n FROM vehicles WHERE status='Available'").fetchone()["n"]
        total_vehicles = db.execute("SELECT COUNT(*) n FROM vehicles").fetchone()["n"]
        busy_vehicles = db.execute("SELECT COUNT(*) n FROM vehicles WHERE status='Busy'").fetchone()["n"]
        breakdown_vehicles = db.execute("SELECT COUNT(*) n FROM vehicles WHERE status='Breakdown'").fetchone()["n"]

        return {
            "task": task_counts,
            "people": {"ว่าง": avail_drivers, "ไม่ว่าง": busy_drivers, "ขัดข้อง": breakdown_drivers, "ทั้งหมด": total_drivers},
            "vehicle": {"ว่าง": avail_vehicles, "ไม่ว่าง": busy_vehicles, "ขัดข้อง": breakdown_vehicles, "ทั้งหมด": total_vehicles},
        }


@app.get("/dashboard/task-volume")
def dashboard_task_volume(days: int = 7, user=Depends(require_role("ADMIN"))):
    days = max(1, min(days, 31))
    statuses = ["Waiting", "Assigned", "In Progress", "Pause", "Completed", "Cancelled"]
    with get_db() as db:
        today = datetime.utcnow().date()
        series = []
        for i in range(days - 1, -1, -1):
            d = today - timedelta(days=i)
            d_str = d.isoformat()
            row_counts = {}
            for s in statuses:
                c = db.execute(
                    "SELECT COUNT(*) n FROM tasks WHERE date(created_at)=? AND status=?", (d_str, s)
                ).fetchone()
                row_counts[s] = c["n"]
            series.append({"date": d_str, **row_counts})

        def count_since(days_back):
            since = (datetime.utcnow() - timedelta(days=days_back)).isoformat(timespec="seconds") + "Z"
            return db.execute("SELECT COUNT(*) n FROM tasks WHERE created_at>=?", (since,)).fetchone()["n"]

        today_str = today.isoformat()
        today_count = db.execute("SELECT COUNT(*) n FROM tasks WHERE date(created_at)=?", (today_str,)).fetchone()["n"]
        summary = {
            "today": today_count,
            "this_week": count_since(7),
            "this_month": count_since(30),
        }
        return {"series": series, "summary": summary}


@app.get("/dashboard/overview")
def dashboard_overview(period: str = "today", user=Depends(require_role("ADMIN"))):
    """Management KPIs: 'how are we performing' — period-bound flow metrics with
    trend vs the immediately preceding period of equal length, plus live snapshot
    metrics (in-progress/overdue/utilization) that don't have a meaningful trend."""
    days = {"today": 1, "week": 7, "month": 30}.get(period, 1)
    now = datetime.utcnow()
    cur_start, cur_end = now - timedelta(days=days), now
    prev_start, prev_end = now - timedelta(days=days * 2), cur_start

    def iso(dt):
        return dt.isoformat(timespec="seconds") + "Z"

    with get_db() as db:
        def flow_metrics(start, end):
            total = db.execute(
                "SELECT COUNT(*) n FROM tasks WHERE created_at>=? AND created_at<?", (iso(start), iso(end))
            ).fetchone()["n"]
            completed_rows = db.execute(
                "SELECT pallet_qty, started_at, completed_at FROM tasks "
                "WHERE status='Completed' AND completed_at>=? AND completed_at<?",
                (iso(start), iso(end)),
            ).fetchall()
            completed = len(completed_rows)
            total_pallets = sum(r["pallet_qty"] for r in completed_rows)
            durations_min, total_hours = [], 0.0
            for r in completed_rows:
                if r["started_at"] and r["completed_at"]:
                    try:
                        s = datetime.fromisoformat(r["started_at"].replace("Z", ""))
                        e = datetime.fromisoformat(r["completed_at"].replace("Z", ""))
                        mins = (e - s).total_seconds() / 60
                        if mins > 0:
                            durations_min.append(mins)
                            total_hours += mins / 60
                    except ValueError:
                        pass
            avg_process = sum(durations_min) / len(durations_min) if durations_min else 0
            qty_per_hour = (total_pallets / total_hours) if total_hours > 0 else 0
            completion_rate = (completed / total * 100) if total > 0 else 0
            return {
                "total_jobs": total, "completed": completed,
                "completion_rate": round(completion_rate, 1),
                "avg_process_time_minutes": round(avg_process, 1),
                "qty_per_hour": round(qty_per_hour, 1),
            }

        cur = flow_metrics(cur_start, cur_end)
        prev = flow_metrics(prev_start, prev_end)

        def pct_change(c, p):
            return round((c - p) / p * 100, 1) if p else None

        trend = {k: pct_change(cur[k], prev[k]) for k in cur}

        in_progress = db.execute("SELECT COUNT(*) n FROM tasks WHERE status='In Progress'").fetchone()["n"]
        waiting_overdue = db.execute(
            "SELECT COUNT(*) n FROM tasks WHERE status='Waiting' AND created_at<?",
            (iso(now - timedelta(minutes=WAITING_TOO_LONG_MINUTES)),),
        ).fetchone()["n"]
        inprogress_overdue = db.execute(
            "SELECT COUNT(*) n FROM tasks WHERE status='In Progress' AND started_at<?",
            (iso(now - timedelta(minutes=IN_PROGRESS_OVERDUE_MINUTES)),),
        ).fetchone()["n"]
        overdue = waiting_overdue + inprogress_overdue

        total_drivers = db.execute("SELECT COUNT(*) n FROM users WHERE role='DRIVER' AND deleted_at IS NULL").fetchone()["n"]
        busy_drivers = db.execute("SELECT COUNT(*) n FROM users WHERE role='DRIVER' AND driver_status='Busy' AND deleted_at IS NULL").fetchone()["n"]
        total_vehicles = db.execute("SELECT COUNT(*) n FROM vehicles").fetchone()["n"]
        busy_vehicles = db.execute("SELECT COUNT(*) n FROM vehicles WHERE status='Busy'").fetchone()["n"]

        return {
            "period": period,
            "kpi": {
                **cur,
                "in_progress": in_progress,
                "overdue": overdue,
                "people_utilization": round(busy_drivers / total_drivers * 100, 1) if total_drivers else 0,
                "vehicle_utilization": round(busy_vehicles / total_vehicles * 100, 1) if total_vehicles else 0,
            },
            "trend": trend,
        }


STATUS_TH = {"Waiting":"รอดำเนินการ","Assigned":"มอบหมายแล้ว","In Progress":"กำลังทำ",
             "Pause":"หยุดชั่วคราว","Completed":"เสร็จสิ้น","Cancelled":"ยกเลิก"}
PRIORITY_TH = {"High":"สูง","Normal":"ปกติ","Low":"ต่ำ"}
REQUEST_TYPE_TH = {"Request Now":"ขอทันที","Booking":"จองล่วงหน้า"}


def _parse_iso(s):
    if not s:
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", ""))
    except ValueError:
        return None


def _fmt_dt(s):
    dt = _parse_iso(s)
    return dt.strftime("%Y-%m-%d %H:%M") if dt else ""


def _minutes_between(a, b):
    da, db_ = _parse_iso(a), _parse_iso(b)
    if not da or not db_:
        return ""
    return round((db_ - da).total_seconds() / 60)


def _xlsx_style_sheet(ws, headers, rows, col_widths, note=None):
    header_fill = PatternFill("solid", fgColor="12151A")
    header_font = Font(name="Arial", bold=True, color="F5A623", size=10)
    body_font = Font(name="Arial", size=10)
    note_font = Font(name="Arial", size=9, italic=True, color="666666")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")

    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = border
    for r in rows:
        ws.append(r)
    for r in range(2, len(rows) + 2):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = body_font
            cell.border = border
            cell.alignment = wrap
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 32
    if note:
        note_row = len(rows) + 3
        ws.cell(row=note_row, column=1, value=note).font = note_font


@app.get("/dashboard/export")
def export_tasks(period: str = "all", user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        if period == "all":
            tasks = db.execute("SELECT * FROM tasks ORDER BY created_at DESC").fetchall()
        else:
            days = {"today": 1, "week": 7, "month": 30}.get(period, 1)
            since = (datetime.utcnow() - timedelta(days=days)).isoformat(timespec="seconds") + "Z"
            tasks = db.execute("SELECT * FROM tasks WHERE created_at>=? ORDER BY created_at DESC", (since,)).fetchall()

        zone_names = {z["zone_key"]: z["zone_name_th"] for z in db.execute("SELECT * FROM zones").fetchall()}

        assignments_by_task = {}
        for a in db.execute(
            "SELECT task_assignments.*, users.full_name FROM task_assignments "
            "JOIN users ON users.id = task_assignments.driver_id"
        ).fetchall():
            assignments_by_task.setdefault(a["task_id"], []).append(dict(a))
        vehicle_code_by_id = {v["id"]: v["code"] for v in db.execute("SELECT id, code FROM vehicles").fetchall()}

        history_by_task = {}
        for h in db.execute("SELECT * FROM task_status_history ORDER BY task_id, changed_at").fetchall():
            history_by_task.setdefault(h["task_id"], []).append(dict(h))

        send_back_by_task = {}
        for log_row in db.execute("SELECT * FROM audit_logs WHERE action='send_back_task' ORDER BY created_at").fetchall():
            m = re.match(r"task (\S+): (.*)", log_row["details"] or "")
            if m:
                send_back_by_task.setdefault(m.group(1), []).append((log_row["created_at"], m.group(2)))

        req_ids = list({t["requester_id"] for t in tasks})
        req_by_id = {}
        if req_ids:
            placeholders = ",".join("?" * len(req_ids))
            req_by_id = {u["id"]: dict(u) for u in db.execute(
                f"SELECT * FROM users WHERE id IN ({placeholders})", req_ids
            ).fetchall()}

        task_rows = []
        pause_rows = []
        driver_name_by_id = {}

        for t in tasks:
            t = dict(t)
            req = req_by_id.get(t["requester_id"], {})
            assigns = assignments_by_task.get(t["id"], [])
            driver_names = []
            seen_d = set()
            for a in assigns:
                if a["driver_id"] not in seen_d:
                    seen_d.add(a["driver_id"])
                    driver_names.append(a["full_name"] or "")
                    driver_name_by_id[a["driver_id"]] = a["full_name"] or ""
            vehicle_codes = []
            seen_v = set()
            for a in assigns:
                code = vehicle_code_by_id.get(a["vehicle_id"])
                if code and a["vehicle_id"] not in seen_v:
                    seen_v.add(a["vehicle_id"])
                    vehicle_codes.append(code)

            hist = history_by_task.get(t["id"], [])
            pause_count = 0
            total_pause_minutes = 0
            open_pause_start = None
            for h in hist:
                if h["to_status"] == "Pause":
                    pause_count += 1
                    open_pause_start = h["changed_at"]
                    note = h["note"] or ""
                    m = re.match(r"\[(\w+)\]\s*(.*)", note)
                    category_th = "ฝั่งผู้ขอ" if m and m.group(1) == "user" else ("ฝั่งรถ" if m else "")
                    reason = m.group(2) if m else note
                    pause_rows.append({
                        "task_code": t["task_code"],
                        "driver": driver_name_by_id.get(h["changed_by"], ""),
                        "start": h["changed_at"], "end": None,
                        "category": category_th, "reason": reason,
                    })
                elif h["from_status"] == "Pause" and h["to_status"] == "In Progress" and open_pause_start:
                    dur = _minutes_between(open_pause_start, h["changed_at"])
                    if pause_rows and pause_rows[-1]["task_code"] == t["task_code"] and pause_rows[-1]["end"] is None:
                        pause_rows[-1]["end"] = h["changed_at"]
                        pause_rows[-1]["duration"] = dur
                    if isinstance(dur, (int, float)):
                        total_pause_minutes += dur
                    open_pause_start = None

            sb = send_back_by_task.get(t["task_code"], [])
            if t["blocked_reason"]:
                blocked_label = "ใช่ (ยังไม่แก้ไข)"
                blocked_reason, blocked_at = t["blocked_reason"], t["blocked_at"]
            elif sb:
                blocked_label = "เคย (แก้ไขและส่งใหม่แล้ว)"
                blocked_at, blocked_reason = sb[-1]
            else:
                blocked_label, blocked_reason, blocked_at = "ไม่", "", ""

            task_rows.append([
                t["task_code"], STATUS_TH.get(t["status"], t["status"]), t["task_type"],
                PRIORITY_TH.get(t["priority"], t["priority"]), REQUEST_TYPE_TH.get(t["request_type"], t["request_type"]),
                _fmt_dt(t["scheduled_at"]),
                zone_names.get(t["from_zone"], t["from_zone"]), t["from_location"],
                zone_names.get(t["to_zone"], t["to_zone"]), t["to_location"],
                t["pallet_qty"],
                req.get("full_name") or "", req.get("employee_id") or "", req.get("cost_center") or "", req.get("contact") or "",
                ", ".join(driver_names), ", ".join(vehicle_codes),
                _fmt_dt(t["created_at"]), _fmt_dt(t["accepted_at"]), _fmt_dt(t["started_at"]), _fmt_dt(t["completed_at"]),
                _minutes_between(t["created_at"], t["accepted_at"]),
                _minutes_between(t["started_at"], t["completed_at"]),
                blocked_label, blocked_reason, _fmt_dt(blocked_at),
                t["cancel_reason"] or "",
                t["rating"] or "", t["rating_comment"] or "",
                pause_count, total_pause_minutes if pause_count else 0,
            ])

        headers1 = [
            "รหัสงาน", "สถานะ", "ประเภทงาน", "ความสำคัญ", "ประเภทคำขอ", "วันเวลานัดหมาย (Booking)",
            "โซนต้นทาง", "ตำแหน่งต้นทาง", "โซนปลายทาง", "ตำแหน่งปลายทาง",
            "จำนวนพาเลท",
            "ผู้ขอ - ชื่อ", "ผู้ขอ - รหัสพนักงาน", "ผู้ขอ - Cost Center", "ผู้ขอ - เบอร์ติดต่อ",
            "คนขับ (ทั้งหมดที่เกี่ยวข้อง)", "รถ (ทั้งหมดที่เกี่ยวข้อง)",
            "สร้างเมื่อ", "ตอบรับเมื่อ (Accept)", "เริ่มงานเมื่อ (Start)", "เสร็จงานเมื่อ (Complete)",
            "เวลารอ (นาที)", "เวลาที่ใช้ทำงานจริง (นาที)",
            "ถูกตีกลับ?", "เหตุผลตีกลับ", "เวลาตีกลับ",
            "เหตุผลยกเลิก",
            "คะแนน (1-5)", "คอมเมนต์คะแนน",
            "จำนวนครั้งที่หยุด", "เวลารวมที่หยุด (นาที)",
        ]
        headers2 = ["รหัสงาน", "คนขับที่รายงาน", "เริ่มหยุดเมื่อ", "กลับมาทำงานเมื่อ",
                    "ระยะเวลาที่หยุด (นาที)", "หมวดปัญหา", "รายละเอียดที่แจ้ง"]
        rows2 = [
            [p["task_code"], p["driver"], _fmt_dt(p["start"]),
             _fmt_dt(p["end"]) if p["end"] else "(ยังไม่กลับมาทำงาน)",
             p.get("duration", "(ยังไม่จบ)") if p["end"] else "(ยังไม่จบ)",
             p["category"], p["reason"]]
            for p in pause_rows
        ]

        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "งาน"
        _xlsx_style_sheet(ws1, headers1, task_rows,
                           col_widths=[13,11,12,10,12,20, 13,15,13,15, 10, 16,14,12,14, 26,16,
                                       17,17,17,17, 10,12, 16,20,17, 24, 8,26, 10,12])
        ws2 = wb.create_sheet("ช่วงเวลาที่หยุด")
        _xlsx_style_sheet(ws2, headers2, rows2, col_widths=[13,16,17,20,16,12,30])

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return StreamingResponse(
            buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=scdc_export_{period}.xlsx"},
        )


# ---------------------------------------------------------------------------
# Static frontend + uploads
# ---------------------------------------------------------------------------

try:
    if _IMPORT_ERROR:
        # The Postgres driver failed to import. Don't fall through to init_db()
        # here — on Vercel that would try to create a SQLite file on a
        # read-only filesystem and crash again. Skip init entirely and let
        # /healthz report the import error as the root cause.
        raise RuntimeError(f"Postgres driver import failed: {_IMPORT_ERROR}")
    init_db()
    ensure_vapid_keys()
except Exception as e:
    # On a normal server (Replit, a VM) this runs once at startup and any DB
    # problem should surface loudly. But on Vercel this module is imported to
    # handle a request, and a raised exception here means EVERY request dies
    # with FUNCTION_INVOCATION_FAILED and no useful message. Log the real
    # error so it shows up in Runtime Logs, and let the app finish importing;
    # the health/diagnostics route below can then report what went wrong
    # instead of the whole function being dead on arrival.
    log.error("Startup init failed: %r", e)
    _STARTUP_ERROR = repr(e)
else:
    _STARTUP_ERROR = None


@app.get("/healthz")
def healthz():
    """Plain diagnostics endpoint that never touches the DB at import time —
    so even if startup init failed, this still answers and tells you WHY,
    turning an opaque 500 into a readable message."""
    return {
        "ok": _STARTUP_ERROR is None and _IMPORT_ERROR is None,
        "startup_error": _STARTUP_ERROR,
        "import_error": _IMPORT_ERROR,
        "database": "postgres" if IS_POSTGRES else "sqlite",
        "database_url_set": bool(DATABASE_URL),
        "web_push": WEB_PUSH_AVAILABLE,
    }


class AdminResetBody(BaseModel):
    secret: str
    new_password: str


@app.post("/admin/emergency-reset-password")
def emergency_reset_admin_password(body: AdminResetBody):
    """Recovery hatch for when the one-time bootstrap admin password was
    missed (e.g. it scrolled out of the deploy logs). Gated by a secret set
    as the ADMIN_RESET_SECRET environment variable — if that variable isn't
    set, this endpoint is disabled entirely, so it can't be abused on a
    normal running system. Resets the password of the earliest-created admin
    account. Set the env var, call this once, then delete the env var (or
    leave it — it's still safe, but removing it closes the hatch)."""
    reset_secret = os.environ.get("ADMIN_RESET_SECRET")
    if not reset_secret:
        raise HTTPException(404, "Not found")
    if not hmac.compare_digest(body.secret, reset_secret):
        raise HTTPException(403, "Invalid secret")
    if len(body.new_password) < MIN_PASSWORD_LENGTH:
        raise HTTPException(400, f"รหัสผ่านต้องมีอย่างน้อย {MIN_PASSWORD_LENGTH} ตัวอักษร")
    with get_db() as db:
        admin = db.execute(
            "SELECT id, employee_id FROM users WHERE role='ADMIN' AND deleted_at IS NULL "
            "ORDER BY id LIMIT 1"
        ).fetchone()
        if not admin:
            raise HTTPException(404, "ไม่พบบัญชี admin")
        db.execute("UPDATE users SET password_hash=? WHERE id=?",
                   (hash_password(body.new_password), admin["id"]))
        db.execute("DELETE FROM sessions WHERE user_id=?", (admin["id"],))
        return {"ok": True, "employee_id": admin["employee_id"],
                "message": "ตั้งรหัสผ่านใหม่แล้ว ล็อกอินได้เลย และอย่าลืมลบ ADMIN_RESET_SECRET ออกจาก environment variables"}

if not IS_POSTGRES:
    # Only relevant for disk-backed deployments — serves photos uploaded
    # before the DB-storage change above, or during local/Replit-style dev.
    # Skipped entirely on Postgres/Vercel, where UPLOAD_DIR is never created
    # and every task photo goes through /tasks/{id}/photo-data/{photo_id} instead.
    try:
        app.mount("/uploads", StaticFiles(directory=UPLOAD_DIR), name="uploads")
    except RuntimeError as e:
        log.warning("Could not mount /uploads (%s) — skipping, old photo links won't resolve", e)

try:
    app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")
except RuntimeError as e:
    # A missing frontend/ directory here (e.g. if a deployment platform's
    # bundler didn't include it) used to crash the entire app at import time
    # — every request would fail with FUNCTION_INVOCATION_FAILED, not just
    # the ones for static files. Log it and keep going so the API routes
    # (and any other static route below) still work even if this one can't.
    log.warning("Could not mount /static (%s) — the SPA frontend won't be served, but the API will still respond", e)


@app.get("/")
def index():
    return FileResponse(os.path.join(FRONTEND_DIR, "index.html"))


@app.get("/service-worker.js")
def service_worker():
    # Served at the root path (not /static/...) so its default scope covers the whole app.
    return FileResponse(os.path.join(FRONTEND_DIR, "service-worker.js"), media_type="application/javascript")


@app.get("/manifest.json")
def manifest():
    return FileResponse(os.path.join(FRONTEND_DIR, "manifest.json"), media_type="application/manifest+json")


if __name__ == "__main__":
    # Lets the app run directly with `python main.py`, which is how Replit
    # (and similar platforms) typically start a Python app. Respects the
    # $PORT environment variable these platforms assign automatically,
    # falling back to 8000 for plain local use. This is in addition to (not
    # instead of) running via `uvicorn main:app ...` on the command line,
    # which still works exactly as before for local dev / LAN deployment.
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
