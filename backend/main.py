"""
SCDC - Vehicle Handling Task Management (MVP v3)
Single-file backend: FastAPI + SQLite.

Run:
    pip install -r requirements.txt
    python -m uvicorn main:app --reload
Then open http://localhost:8000 in a browser.
"""

import base64
import contextvars
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
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
from datetime import datetime, timedelta
from typing import Optional, List

from fastapi import FastAPI, Depends, HTTPException, Header, UploadFile, File
from fastapi.middleware.gzip import GZipMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse, Response
from starlette.background import BackgroundTask
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
VAPID_CLAIM_SUB = "mailto:admin@scdc.local"   # legacy fallback only — see resolve_vapid_subject()


# Apple's push service validates the VAPID "sub" (contact) claim and answers
# 403 BadJwtToken unless it is a real mailto: address or https: URL on a real
# domain. Google's push service doesn't check it, which is why push worked on
# Android while failing on every iPhone. The legacy value above uses ".local",
# a reserved domain that can never resolve, so it is never used for Apple.
_UNUSABLE_SUFFIXES = (".local", ".localhost", ".invalid", ".test", ".example", ".internal")


def _subject_host(subject: str):
    s = (subject or "").strip()
    if s.startswith("mailto:"):
        addr = s[len("mailto:"):]
        return addr.rsplit("@", 1)[1].lower() if "@" in addr else None
    if s.startswith("https://"):
        return s[len("https://"):].split("/")[0].split(":")[0].lower() or None
    return None


def subject_is_usable(subject: str) -> bool:
    host = _subject_host(subject)
    if not host or "." not in host or host == "localhost":
        return False
    return not host.endswith(_UNUSABLE_SUFFIXES)


def resolve_vapid_subject():
    """First usable contact from: VAPID_SUBJECT, then this deployment's own
    public URL (Vercel sets these automatically). Returns (subject, source).

    A bare email in VAPID_SUBJECT is accepted and given its mailto: prefix —
    a missing scheme is an easy slip in a dashboard, and it fails exactly as
    silently as a bad domain does.
    """
    raw = (os.environ.get("VAPID_SUBJECT") or "").strip()
    if raw and "@" in raw and not raw.startswith(("mailto:", "https://")):
        raw = "mailto:" + raw
    candidates = [
        (raw, "VAPID_SUBJECT"),
        ("https://" + os.environ["VERCEL_PROJECT_PRODUCTION_URL"]
         if os.environ.get("VERCEL_PROJECT_PRODUCTION_URL") else "", "VERCEL_PROJECT_PRODUCTION_URL"),
        ("https://" + os.environ["VERCEL_URL"] if os.environ.get("VERCEL_URL") else "", "VERCEL_URL"),
    ]
    for value, source in candidates:
        if value and subject_is_usable(value):
            return value, source
    return VAPID_CLAIM_SUB, "fallback (ใช้กับ iPhone ไม่ได้)"


VAPID_SUBJECT, VAPID_SUBJECT_SOURCE = resolve_vapid_subject()

# ---------------------------------------------------------------------------
# Build version — BUMP THIS ON EVERY RELEASE.
# Format: TTLOD.yymmdd.hhmm, using Bangkok time at the moment the build is cut.
# This is the single source of truth: the number is substituted into the page
# when index.html is served, so the frontend can't drift out of step with it.
# ---------------------------------------------------------------------------
APP_VERSION = "TTLOD.261028.0827"

# Advance booking ("จองล่วงหน้า") is switched off, not deleted.
#
# Nothing offers it any more and the server refuses it, but the columns, the
# stored appointment times and the report column all stay exactly as they
# were: jobs booked before this still show their time and still export
# correctly. Turning it back on is this one line here and its twin in
# index.html — nothing else has to be rebuilt.
BOOKING_ENABLED = False

# Thai labels for statuses that appear inside user-facing error messages. The
# frontend has its own copies for rendering; these exist so a message built on
# the server doesn't leak an English status word into a Thai-only UI.
TASK_STATUS_LABEL_TH = {
    "Waiting": "รอรับงาน", "Assigned": "มอบหมายแล้ว", "In Progress": "กำลังดำเนินการ",
    "Pause": "หยุดชั่วคราว", "Completed": "เสร็จสิ้น", "Cancelled": "ยกเลิกแล้ว",
}
PRIORITY_LABEL_TH = {"Normal": "ปกติ", "Urgent": "ด่วน", "Critical": "ด่วนมาก"}
DRIVER_STATUS_LABEL_TH = {
    "Available": "ว่าง", "Busy": "ติดงาน", "Breakdown": "รถเสีย", "Offline": "ออฟไลน์",
}

app = FastAPI(title="CRL ForkliftDispatch", description="Warehouse Forklift Request & Dispatch")

# /app-data returns a large JSON document on every refresh and was being sent
# uncompressed. It is highly repetitive JSON, so gzip typically cuts it by
# ~80% — a direct win on warehouse mobile connections.
app.add_middleware(GZipMiddleware, minimum_size=1000)


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


def send_web_push_to_user(db, user_id: int, title: str, body: str, url: str = "/",
                          tag: Optional[str] = None):
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
    # The tag decides whether a new notification STACKS or REPLACES on the
    # phone. Every notification used to carry the same one, so three new jobs
    # arriving together showed up as a single line and the first two were never
    # seen. One tag per job (so a later update about that job replaces the
    # earlier one, which is what you want) and otherwise one tag per message.
    payload = json.dumps({"title": title, "body": body, "url": url,
                          "tag": tag or f"scdc-{secrets.token_hex(4)}"})

    def _one(s):
        """Returns (sent?, subscription_id_to_delete_or_None, (status, reason) or None)."""
        try:
            webpush(
                subscription_info={"endpoint": s["endpoint"], "keys": {"p256dh": s["p256dh"], "auth": s["auth"]}},
                data=payload,
                vapid_private_key=vapid,
                vapid_claims={"sub": VAPID_SUBJECT},
                ttl=60,
            )
            log.info("Web push sent to user %s (subscription %s)", user_id, s["id"])
            return True, None, None
        except WebPushException as e:
            resp = getattr(e, "response", None)
            status = getattr(resp, "status_code", None)
            reason = None
            try:
                reason = (resp.json() or {}).get("reason")
            except Exception:
                try:
                    reason = (resp.text or "").strip()[:120] or None
                except Exception:
                    reason = None
            # Only a subscription that is genuinely dead is deleted:
            #   404/410             — the browser dropped it
            #   403 VapidPkHashMismatch — it was made for a different server key
            # A 403 BadJwtToken is different: it means OUR signature was
            # refused (Apple rejects an unusable VAPID subject this way). The
            # subscription is fine. The previous version deleted it anyway and
            # told the person to switch notifications off and on — which could
            # never work, since every new subscription hit the same 403 and was
            # deleted too.
            if status in (404, 410) or (status == 403 and reason == "VapidPkHashMismatch"):
                log.info("Push subscription %s for user %s is dead (%s %s) — removing it",
                         s["id"], user_id, status, reason)
                return False, s["id"], (status, reason)
            log.warning("Web push refused for subscription %s (user %s): %s %s",
                        s["id"], user_id, status, reason)
            return False, None, (status, reason)
        except Exception as e:
            # Most commonly: the server has no outbound internet access to reach
            # the browser's push relay (fcm.googleapis.com / Mozilla autopush /
            # Apple push) — very possible on a LAN-only office deployment.
            log.warning("Web push error for subscription %s (user %s): %s", s["id"], user_id, e)
            return False, None, (None, str(e)[:120])

    # Devices are pushed to CONCURRENTLY. Each webpush() is a blocking HTTPS
    # call out to FCM/Mozilla/Apple; done in sequence, a user with three
    # devices meant three round-trips stacked end to end. The DB is
    # deliberately not touched from these threads (a psycopg2 connection is
    # not safe to share) — failures come back as ids and are deleted below,
    # on this thread.
    sent = 0
    stale, errors = [], []
    results = [_one(subs[0])] if len(subs) == 1 else None
    if results is None:
        with ThreadPoolExecutor(max_workers=min(8, len(subs))) as pool:
            results = list(pool.map(_one, subs))
    for ok, bad, err in results:
        sent += 1 if ok else 0
        if bad:
            stale.append(bad)
        if err:
            errors.append({"status": err[0], "reason": err[1]})
    for sub_id in stale:
        db.execute("DELETE FROM push_subscriptions WHERE id=?", (sub_id,))
    return {"attempted": len(subs), "sent": sent, "errors": errors,
            "reason": None if sent else _explain_push_failure(errors)}


def _explain_push_failure(errors) -> str:
    """Turn the push service's answer into something a person can act on.
    "Try switching it off and on" is only offered when it can actually help."""
    reasons = {e.get("reason") for e in errors}
    statuses = {e.get("status") for e in errors}
    if "BadJwtToken" in reasons:
        return ("Apple ปฏิเสธลายเซ็นของเซิร์ฟเวอร์ (BadJwtToken) — เป็นปัญหาการตั้งค่าฝั่งเซิร์ฟเวอร์ "
                "ปิด-เปิดแจ้งเตือนใหม่จะไม่ช่วย กรุณาแจ้งผู้ดูแลระบบให้ตั้งค่า VAPID_SUBJECT")
    if "VapidPkHashMismatch" in reasons or statuses & {404, 410}:
        return "การลงทะเบียนแจ้งเตือนบนเครื่องนี้หมดอายุแล้ว — กรุณาปิดแล้วเปิดแจ้งเตือนใหม่อีกครั้ง"
    if None in statuses:
        return "เซิร์ฟเวอร์ติดต่อบริการแจ้งเตือนไม่ได้ (" + "; ".join(
            str(e.get("reason")) for e in errors if e.get("status") is None)[:120] + ")"
    detail = ", ".join(f"{e.get('status')} {e.get('reason') or ''}".strip() for e in errors)
    return f"บริการแจ้งเตือนปฏิเสธการส่ง ({detail})"


# --------------------------------------------------------------------------
# Deferred push delivery
# --------------------------------------------------------------------------
# notify() used to call send_web_push_to_user() inline, so the HTTP request
# that triggered a notification sat waiting for every push to be delivered to
# Google/Apple before it could return. notify_ops() made that far worse: it
# looped over every admin, and each admin over each of their devices. With 5
# admins on 2 devices, a driver tapping "รับงาน" waited on 10 sequential
# outbound HTTPS calls before the button released.
#
# Pushes raised during a request are now collected here and flushed AFTER the
# response has been sent (Starlette BackgroundTask — still inside the same
# serverless invocation, so nothing is lost on Vercel). A contextvar holding a
# shared list is what carries them: FastAPI copies the context into the worker
# thread that runs the endpoint, and because the list object itself is shared,
# appends made in that thread are visible back here.
_push_queue_var = contextvars.ContextVar("scdc_push_queue", default=None)


def _flush_push_queue(jobs):
    """Runs after the response is delivered. Opens its own connection: the
    request's was already committed and closed by then."""
    try:
        with get_db() as db:
            for user_id, title, body, url, tag in jobs:
                try:
                    send_web_push_to_user(db, user_id, title, body, url, tag)
                except Exception as e:
                    log.warning("Deferred push to user %s failed: %s", user_id, e)
    except Exception as e:
        # A push problem must never be able to surface as a request failure —
        # by this point the user already has their response.
        log.warning("Deferred push flush failed: %s", e)


@app.middleware("http")
async def _deferred_push_middleware(request, call_next):
    queue = []
    token = _push_queue_var.set(queue)
    try:
        response = await call_next(request)
    finally:
        _push_queue_var.reset(token)
    if queue:
        response.background = BackgroundTask(_flush_push_queue, list(queue))
    return response


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


# One database connection PER REQUEST instead of per get_db() call.
#
# Every endpoint function opens `with get_db() as db`, which used to mean a
# brand new psycopg2.connect() — a full TCP + TLS handshake to Neon — each
# time. /app-data deliberately bundles ~12 endpoint functions into a single
# HTTP round-trip, but because each of them opened its own connection, one
# poll still paid for ~13 handshakes (roughly 50-150ms each, even against a
# pooled endpoint: pooling saves the Postgres-side backend startup, not the
# client-side TLS negotiation). That was the single largest source of latency
# in the app, and it repeated every poll cycle for every signed-in user.
#
# get_db() is now re-entrant: the outermost `with` opens the connection and
# owns commit/close, and any nested `with get_db()` inside the same thread
# hands back that same connection. Call sites are unchanged.
#
# threading.local (rather than a contextvar) is the right scope here because
# FastAPI runs every `def` (non-async) endpoint and dependency in a worker
# thread, and one request's whole call tree stays on that one thread.
#
# Side effect, and an improvement: a request is now a single transaction, so
# a failure partway through /tasks/{id}/self-assign no longer leaves half its
# writes committed.
_db_local = threading.local()


@contextmanager
def _new_db_connection():
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


@contextmanager
def get_db():
    existing = getattr(_db_local, "conn", None)
    if existing is not None:
        # Nested call — reuse the connection the outermost caller opened, and
        # leave commit/close to it.
        yield existing
        return
    with _new_db_connection() as conn:
        _db_local.conn = conn
        try:
            yield conn
        finally:
            # Always clear, even on error: worker threads are reused across
            # requests and must never inherit a closed connection.
            _db_local.conn = None


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


def iso_days_ago(days: int) -> str:
    return (datetime.utcnow() - timedelta(days=days)).isoformat(timespec="seconds") + "Z"


# How long a signed-in device stays signed in without being used. Sessions are
# pruned opportunistically at login (there is no scheduler), so this is an
# upper bound on how long an abandoned token can linger, not a precise expiry.
SESSION_MAX_AGE_DAYS = 90


def describe_user_agent(ua: Optional[str]) -> str:
    """A short, recognisable name for a signed-in device.

    The point is only that someone looking at their own device list can tell
    the warehouse tablet from their phone, so this stays deliberately coarse —
    no UA parsing library, no version numbers, no fingerprinting.
    """
    if not ua:
        return "อุปกรณ์ไม่ทราบชนิด"
    u = ua.lower()
    if "ipad" in u:
        os_name = "iPad"
    elif "iphone" in u:
        os_name = "iPhone"
    elif "android" in u:
        os_name = "Android"
    elif "windows" in u:
        os_name = "Windows"
    elif "mac os" in u or "macintosh" in u:
        os_name = "Mac"
    else:
        os_name = "อุปกรณ์อื่น"
    # Order matters: Edge and Chrome both claim Safari, Edge also claims Chrome.
    if "edg/" in u:
        browser = "Edge"
    elif "crios" in u or "chrome" in u:
        browser = "Chrome"
    elif "fxios" in u or "firefox" in u:
        browser = "Firefox"
    elif "safari" in u:
        browser = "Safari"
    else:
        browser = None
    return f"{os_name} · {browser}" if browser else os_name


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


# ---------------------------------------------------------------------------
# App configuration
# ---------------------------------------------------------------------------
# Anything an admin should be able to change without a redeploy lives here.
# Kept in app_config (a key/value table) rather than environment variables so
# it survives a deploy and can be edited from inside the product.
CONFIG_DEFAULTS = {
    # Rack location layout. The warehouse defines its own ranges, so nothing
    # here is hard-coded in the validator — it all reads from these values.
    "rack_aisle_from": "AAA",
    "rack_aisle_to": "ABH",
    "rack_bay_from": "01",
    "rack_bay_to": "40",
    "rack_level_from": "A",
    "rack_level_to": "G",
    "rack_slot_from": "1",
    "rack_slot_to": "5",
    # First letter of the aisle -> which racking it is. Extra letters can be
    # added here as new aisles appear, without touching code.
    "rack_prefix_map": "A=DOUBLE_DEEP,B=SELECTIVE_RACK",
}


def get_config(db, key: str) -> str:
    row = db.execute("SELECT value FROM app_config WHERE key=?", (key,)).fetchone()
    if row and row["value"] is not None:
        return row["value"]
    return CONFIG_DEFAULTS.get(key, "")


def set_config(db, key: str, value: str):
    db.execute(
        "INSERT INTO app_config (key, value) VALUES (?,?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, value),
    )


def _rack_prefix_map(db) -> dict:
    out = {}
    for part in (get_config(db, "rack_prefix_map") or "").split(","):
        if "=" in part:
            letter, zone = part.split("=", 1)
            letter = letter.strip().upper()
            if letter:
                out[letter] = zone.strip().upper()
    return out


def rack_layout(db) -> dict:
    """Everything the location validator and the UI need, in one shape."""
    return {
        "aisle_from": get_config(db, "rack_aisle_from").upper(),
        "aisle_to": get_config(db, "rack_aisle_to").upper(),
        "bay_from": int(get_config(db, "rack_bay_from") or 1),
        "bay_to": int(get_config(db, "rack_bay_to") or 40),
        "level_from": get_config(db, "rack_level_from").upper(),
        "level_to": get_config(db, "rack_level_to").upper(),
        "slot_from": int(get_config(db, "rack_slot_from") or 1),
        "slot_to": int(get_config(db, "rack_slot_to") or 5),
        "prefix_map": _rack_prefix_map(db),
    }


def parse_rack_code(code: str, layout: dict):
    """Validate a rack location against the configured layout.

    Format is fixed at Aisle(3) + Bay(2) + Level(1) + Slot(1), e.g. ABA03A2.
    The RANGES are not fixed — they come from settings — which is why this is
    a parser rather than one big regex: a wrong code can then say exactly
    which part is out of range instead of just "invalid format".

    Returns (parsed_dict, error_message). Exactly one is ever non-None.
    """
    code = (code or "").strip().upper()
    if len(code) != 7:
        return None, "รหัสตำแหน่งต้องมี 7 ตัวอักษร เช่น ABA03A2"
    aisle, bay_s, level, slot_s = code[0:3], code[3:5], code[5], code[6]

    if not aisle.isalpha():
        return None, f"Aisle ต้องเป็นตัวอักษร 3 ตัว (ได้รับ '{aisle}')"
    if not (layout["aisle_from"] <= aisle <= layout["aisle_to"]):
        return None, f"Aisle '{aisle}' อยู่นอกช่วง {layout['aisle_from']}-{layout['aisle_to']}"

    if not bay_s.isdigit():
        return None, f"Bay ต้องเป็นตัวเลข 2 หลัก (ได้รับ '{bay_s}')"
    bay = int(bay_s)
    if not (layout["bay_from"] <= bay <= layout["bay_to"]):
        return None, f"Bay {bay_s} อยู่นอกช่วง {layout['bay_from']:02d}-{layout['bay_to']:02d}"

    if not level.isalpha() or not (layout["level_from"] <= level <= layout["level_to"]):
        return None, f"Level '{level}' อยู่นอกช่วง {layout['level_from']}-{layout['level_to']}"

    if not slot_s.isdigit():
        return None, f"Slot ต้องเป็นตัวเลข (ได้รับ '{slot_s}')"
    slot = int(slot_s)
    if not (layout["slot_from"] <= slot <= layout["slot_to"]):
        return None, f"Slot {slot_s} อยู่นอกช่วง {layout['slot_from']}-{layout['slot_to']}"

    rack_zone = layout["prefix_map"].get(aisle[0])
    if not rack_zone:
        known = ", ".join(sorted(layout["prefix_map"])) or "ยังไม่ได้ตั้งค่า"
        return None, f"ยังไม่ได้กำหนดว่าตัวอักษร '{aisle[0]}' เป็นชั้นวางประเภทใด (ที่กำหนดไว้: {known})"

    return {"code": code, "aisle": aisle, "bay": bay, "level": level,
            "slot": slot, "zone_key": rack_zone}, None


def normalize_phone(phone: Optional[str]) -> Optional[str]:
    """Digits only, so 08-1234-5678 and 0812345678 are the same number for
    both uniqueness and login."""
    if not phone:
        return None
    digits = re.sub(r"\D", "", phone)
    return digits or None


def validate_and_normalize_contact(db, contact: Optional[str], exclude_user_id: Optional[int] = None) -> Optional[str]:
    """Phone numbers are now a login identifier, so they must be unique —
    otherwise the system can't tell which account a number belongs to."""
    if not contact:
        return None
    digits = normalize_phone(contact)
    if not digits or not (9 <= len(digits) <= 10):
        raise HTTPException(400, "เบอร์ติดต่อไม่ถูกต้อง (ต้องเป็นเบอร์โทรไทย 9-10 หลัก)")
    existing = db.execute(
        "SELECT id FROM users WHERE contact=? AND deleted_at IS NULL", (digits,)
    ).fetchone()
    if existing and existing["id"] != exclude_user_id:
        raise HTTPException(400, "เบอร์นี้มีผู้ใช้งานแล้ว หากเป็นเบอร์ของคุณโปรดติดต่อ Admin")
    return digits


def validate_and_normalize_email(db, email: Optional[str], exclude_user_id: Optional[int] = None) -> Optional[str]:
    """Shared by register/profile-edit/driver-creation/admin-edit so the
    @central.co.th + uniqueness rules are enforced identically everywhere an
    email can be set, not just at registration."""
    if not email:
        return None
    email = email.strip().lower()
    if not email.endswith("@central.co.th"):
        raise HTTPException(400, "อีเมลต้องเป็นโดเมน @central.co.th เท่านั้น")
    existing = db.execute(
        "SELECT id FROM users WHERE email=? AND deleted_at IS NULL", (email,)
    ).fetchone()
    if existing and existing["id"] != exclude_user_id:
        raise HTTPException(400, "อีเมลนี้มีผู้ใช้งานแล้ว")
    return email


# ---------------------------------------------------------------------------
# Schema
# ---------------------------------------------------------------------------

# A rack code is Aisle(3) + Bay(2) + Level(1) + Slot(1), and a requester may
# stop at any of those four steps: "AAA" names an aisle, "AAA01A1" a single
# slot. Which letters and digits are valid is deliberately not checked — the
# floor renames aisles faster than anyone updates a pattern, and a job refused
# over a letter that is perfectly real costs more than a typo does. Only the
# first letter stays pinned per zone, so a code can never contradict the zone
# it was filed under.
RACK_PATTERN_SELECTIVE = "^B[A-Z]{2}([0-9]{2}([A-Z][0-9]?)?)?$"
RACK_PATTERN_DOUBLE_DEEP = "^A[A-Z]{2}([0-9]{2}([A-Z][0-9]?)?)?$"

# The patterns those replace. A live zone is only rewritten while it still
# carries the original, so a pattern an admin has since edited is left alone.
RACK_PATTERN_SELECTIVE_SEED = "^B[AB][A-H](0[1-9]|[12][0-9]|3[0-9])([A-G][12])?$"
RACK_PATTERN_DOUBLE_DEEP_SEED = "^AA[A-F](0[1-9]|[1-6][0-9]|7[0-4])([A-G][12])?$"

COOLROOM_FIRST_ROOM = "RRE"


SCHEMA = """
CREATE TABLE IF NOT EXISTS vehicles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT UNIQUE NOT NULL,
    vehicle_type TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'Available',
    battery_level INTEGER,
    -- RT is one vehicle type with a sub-capability rather than two types:
    -- the licence doesn't distinguish them, only which racking they can reach.
    supports_double_deep INTEGER NOT NULL DEFAULT 0,
    -- Retired vehicles stay in the table so historical tasks and assignments
    -- still resolve; they're just hidden from anything operational.
    retired_at TEXT
);

CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id TEXT UNIQUE NOT NULL,
    email TEXT,
    full_name TEXT,
    first_name TEXT,
    last_name TEXT,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL CHECK(role IN ('USER','DRIVER','DISPATCH','ADMIN')),
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

CREATE TABLE IF NOT EXISTS driver_zones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    driver_id INTEGER NOT NULL REFERENCES users(id),
    -- The zone's immutable key, the same string tasks store in from_zone.
    zone_key TEXT NOT NULL,
    -- 1 = โซนหลัก, at most one per person. 2 = โซนรอง, as many as they like.
    priority INTEGER NOT NULL DEFAULT 2,
    created_at TEXT NOT NULL,
    UNIQUE(driver_id, zone_key)
);

CREATE TABLE IF NOT EXISTS vehicle_types (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    type_key TEXT UNIQUE NOT NULL,
    type_name_th TEXT NOT NULL,
    description TEXT,
    -- Registering with this code grants a permanent licence for this type.
    -- It lives here, rather than in two fixed config rows, so a type added
    -- later can have one too. NULL means nobody can sign themselves up for
    -- this type and an admin has to grant the licence.
    invite_code TEXT
);

CREATE TABLE IF NOT EXISTS driver_licenses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    driver_id INTEGER NOT NULL REFERENCES users(id),
    vehicle_type TEXT NOT NULL,
    expiry_date TEXT,
    -- The driver's usual vehicle of this type ("รถประจำ"). Boarded for them
    -- automatically when they open the app, so picking a vehicle stops being
    -- a daily chore. It lives on the licence row because a driver licensed
    -- for two types has one usual vehicle of each.
    default_vehicle_id INTEGER REFERENCES vehicles(id),
    UNIQUE(driver_id, vehicle_type)
);

CREATE TABLE IF NOT EXISTS zones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    -- Immutable. Every task ever created stores this string in from_zone /
    -- to_zone, and four zones have their own hand-built picker keyed on it,
    -- so renaming it would orphan history and silently drop those pickers.
    -- The editable name is zone_name_th.
    zone_key TEXT UNIQUE NOT NULL,
    zone_name_th TEXT NOT NULL,
    code_pattern TEXT NOT NULL,
    allow_multiple INTEGER NOT NULL DEFAULT 0,
    example_code TEXT,
    sort_order INTEGER NOT NULL DEFAULT 0,
    free_text INTEGER NOT NULL DEFAULT 0,
    -- Can only be a job's origin (empty pallets are collected, never delivered).
    origin_only INTEGER NOT NULL DEFAULT 0,
    -- The zone IS the location: picking it saves straight away with no detail
    -- to type (Sort Yard, empty pallets).
    no_location INTEGER NOT NULL DEFAULT 0,
    -- The zone's locations are a LIST (zone_locations), not a pattern:
    -- Coolroom and its rooms, where "which rooms exist" is the whole rule.
    from_master INTEGER NOT NULL DEFAULT 0,
    -- Retired zones stay here so old tasks still read properly; they just
    -- stop being offered when creating a job.
    active INTEGER NOT NULL DEFAULT 1
);

-- The named places inside a zone, for zones whose locations are a LIST rather
-- than a pattern (Coolroom and its rooms). A pattern cannot express "these
-- four rooms exist and nothing else", and a rack zone cannot be a list —
-- thousands of slots nobody would maintain — so the two live side by side and
-- zones.from_master picks which one applies.
CREATE TABLE IF NOT EXISTS zone_locations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    zone_key TEXT NOT NULL,
    -- What gets written onto the job. Unique per zone, case-insensitively.
    code TEXT NOT NULL,
    -- Optional longer name for the picker; the code is what the job stores.
    label TEXT,
    sort_order INTEGER NOT NULL DEFAULT 0,
    -- Retired rooms stay here so old jobs still read properly; they just stop
    -- being offered.
    active INTEGER NOT NULL DEFAULT 1,
    -- Added by a requester who needed it, not yet checked by an admin. It is
    -- usable at once — blocking the job to wait for approval would just move
    -- the problem onto the floor — and shows as "รอยืนยัน" until confirmed.
    pending INTEGER NOT NULL DEFAULT 0,
    created_at TEXT,
    UNIQUE(zone_key, code)
);

-- The cost centres jobs are booked against. Before this they were five loose
-- digits typed into a box, so "82899" meant nothing to anyone who did not
-- already know the number, and a typo created a cost centre nobody noticed.
CREATE TABLE IF NOT EXISTS cost_centers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE,
    -- What people call it ("Mgmt"). Blank until someone fills it in.
    name TEXT,
    -- Added by whoever needed it rather than by an admin. Usable immediately
    -- so nobody's work waits on approval; flagged so the list can be tidied.
    pending INTEGER NOT NULL DEFAULT 0,
    active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT
);

-- UNUSED since the vehicle-type/zone matrix was removed. Kept rather than
-- dropped: a DROP TABLE on the live database cannot be undone, and an
-- unread table costs nothing. Delete it in a later release if it is still
-- unused then.
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
    cost_center TEXT,
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

-- NOTE: is_active_for_driver is added by ensure_columns() for existing
-- databases; fresh ones get it here.
CREATE TABLE IF NOT EXISTS task_assignments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_id INTEGER NOT NULL REFERENCES tasks(id),
    driver_id INTEGER NOT NULL REFERENCES users(id),
    -- Nullable: a driver can be assigned before they have boarded a vehicle.
    -- assigned_vehicle_type then says which kind of vehicle the job expects,
    -- and the id is filled in the moment they check one out.
    vehicle_id INTEGER REFERENCES vehicles(id),
    assigned_vehicle_type TEXT,
    assigned_by INTEGER NOT NULL REFERENCES users(id),
    assigned_at TEXT NOT NULL,
    unassigned_at TEXT,
    assignment_status TEXT NOT NULL DEFAULT 'Active',
    reason TEXT,
    -- 1 = the job this driver has in hand, 0 = waiting in their queue.
    is_active_for_driver INTEGER NOT NULL DEFAULT 1
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
    kind TEXT NOT NULL DEFAULT 'user',
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
    created_at TEXT NOT NULL,
    -- Several devices can be signed in at once, so each session records what
    -- it was opened from. Without this the "signed in on these devices" list
    -- would be a row of identical timestamps nobody could act on.
    user_agent TEXT
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


# Indexes the query plans above actually need. Kept separate from SCHEMA and
# run on every boot because Postgres only executes SCHEMA on a FRESH database
# — an already-deployed production database would otherwise never receive
# them. CREATE INDEX IF NOT EXISTS is valid on both SQLite and Postgres and is
# a no-op once the index exists, so this is cheap to repeat.
#
# The sessions(token) one matters most: current_user() looks a token up on
# EVERY authenticated request and, with no index, that was a full scan of a
# table which grows with every login the system has ever seen.
INDEX_STATEMENTS = [
    "CREATE INDEX IF NOT EXISTS idx_driver_zones_driver ON driver_zones(driver_id)",
    "CREATE INDEX IF NOT EXISTS idx_sessions_token ON sessions(token)",
    "CREATE INDEX IF NOT EXISTS idx_tasks_status ON tasks(status)",
    "CREATE INDEX IF NOT EXISTS idx_tasks_created_at ON tasks(created_at)",
    "CREATE INDEX IF NOT EXISTS idx_tasks_requester ON tasks(requester_id)",
    "CREATE INDEX IF NOT EXISTS idx_ta_task_status ON task_assignments(task_id, assignment_status)",
    "CREATE INDEX IF NOT EXISTS idx_ta_driver ON task_assignments(driver_id)",
    "CREATE INDEX IF NOT EXISTS idx_notifications_user ON notifications(user_id, is_read)",
    "CREATE INDEX IF NOT EXISTS idx_push_subs_user ON push_subscriptions(user_id)",
    "CREATE INDEX IF NOT EXISTS idx_task_messages_task ON task_messages(task_id)",
    "CREATE INDEX IF NOT EXISTS idx_task_photos_task ON task_photos(task_id)",
    "CREATE INDEX IF NOT EXISTS idx_status_hist_task ON task_status_history(task_id)",
]


def ensure_indexes():
    try:
        with get_db() as db:
            for stmt in INDEX_STATEMENTS:
                try:
                    db.execute(stmt)
                except Exception as e:
                    # A single index failing (e.g. a column missing on a very
                    # old database) must not stop the rest or block startup.
                    log.warning("Could not create index (%s): %s", stmt, e)
    except Exception as e:
        log.warning("ensure_indexes skipped: %s", e)


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


# Columns this build needs that older databases won't have.
#
# migrate_db() below only ever runs on SQLite — on Postgres, init_db() applies
# SCHEMA only when the database is FRESH, so a live production database would
# never receive a new column. Everything here therefore runs on every boot for
# BOTH backends, checking first and adding only what's missing.
NEW_COLUMNS = [
    # (table, column, type, backfill SQL or None)
    ("users", "first_name", "TEXT", None),
    ("users", "last_name", "TEXT", None),
    ("vehicles", "supports_double_deep", "INTEGER NOT NULL DEFAULT 0", None),
    ("vehicles", "retired_at", "TEXT", None),
    # 'user' for a person's message, 'system' for automatic status updates.
    # sender_id stays the person whose action caused it, so history is never
    # anonymous even for automatic lines.
    ("task_messages", "kind", "TEXT NOT NULL DEFAULT 'user'", None),
    # Which role a driver was assigned AS. Normally the type of the vehicle
    # they are in — but a driver can now be assigned before boarding one, and
    # the job still has to know whether it is waiting on an RT or a PE.
    ("task_assignments", "assigned_vehicle_type", "TEXT", None),
    # Which device a session was opened from, for the device list in Profile.
    ("sessions", "user_agent", "TEXT", None),
    # The driver's usual vehicle for this licence type ("รถประจำ").
    ("driver_licenses", "default_vehicle_id", "INTEGER", None),
    # A driver can hold several jobs at once now. Exactly one of their Active
    # assignments is the job in hand (1); the rest are their queue (0).
    # Existing rows default to 1, which is correct: before queues, every
    # Active assignment WAS the job in hand.
    ("task_assignments", "is_active_for_driver", "INTEGER NOT NULL DEFAULT 1", None),
    # Per-type invite code, replacing the two fixed invite_code_pe/rt settings.
    ("vehicle_types", "invite_code", "TEXT", None),
    # Zone behaviour, moved out of hand-written branches into data so an admin
    # can maintain it without a deploy.
    ("zones", "origin_only", "INTEGER NOT NULL DEFAULT 0", None),
    ("zones", "no_location", "INTEGER NOT NULL DEFAULT 0", None),
    ("zones", "active", "INTEGER NOT NULL DEFAULT 1", None),
    # The zone's locations come from the zone_locations list, not its pattern.
    ("zones", "from_master", "INTEGER NOT NULL DEFAULT 0", None),
    ("zone_locations", "pending", "INTEGER NOT NULL DEFAULT 0", None),
    # The cost centre the job is booked against. It used to be read off the
    # requester's profile, which left nothing to charge against when an admin
    # or dispatcher raised the job on someone's behalf — they have no cost
    # centre of their own.
    ("tasks", "cost_center", "TEXT", None),
]


def _column_exists(db, table, column):
    if IS_POSTGRES:
        return bool(db.execute(
            "SELECT 1 FROM information_schema.columns WHERE table_name=? AND column_name=?",
            (table, column),
        ).fetchone())
    return column in [r["name"] for r in db.execute(f"PRAGMA table_info({table})").fetchall()]


# Tables added after the first release.
#
# Same trap as NEW_COLUMNS, one level up: on Postgres, init_db() applies SCHEMA
# only when the database is FRESH, so a table added to SCHEMA alone would never
# exist on the live one and every screen that touched it would 500. This runs
# on every boot for both backends instead.
NEW_TABLES = ("driver_zones", "zone_locations", "cost_centers")


def ensure_tables():
    try:
        with get_db() as db:
            for name in NEW_TABLES:
                m = re.search(r"CREATE TABLE IF NOT EXISTS " + name + r" \(.*?\n\);", SCHEMA, re.S)
                if not m:
                    log.warning("No CREATE statement in SCHEMA for %s", name)
                    continue
                ddl = m.group(0)
                if IS_POSTGRES:
                    ddl = (ddl.replace("INTEGER PRIMARY KEY AUTOINCREMENT", "SERIAL PRIMARY KEY")
                              .replace("BLOB", "BYTEA"))
                db.execute(ddl)
    except Exception as e:
        log.warning("ensure_tables skipped: %s", e)


def ensure_seed_zones():
    """Corrections to the seeded zones, applied to a database that already exists.

    init_db() seeds only a FRESH database, so a zone added or corrected there
    would never reach production. Every step here is conditional: a pattern is
    rewritten only while it still matches the one it replaces, and a zone or
    room is inserted only when missing — so an admin's own edits in
    จัดการระบบ survive this and every later deploy.
    """
    try:
        with get_db() as db:
            for zone_key, was, now, example in (
                ("SELECTIVE_RACK", RACK_PATTERN_SELECTIVE_SEED, RACK_PATTERN_SELECTIVE, "BAA01A1"),
                ("DOUBLE_DEEP", RACK_PATTERN_DOUBLE_DEEP_SEED, RACK_PATTERN_DOUBLE_DEEP, "AAA01A1"),
            ):
                cur = db.execute(
                    "UPDATE zones SET code_pattern=?, example_code=? WHERE zone_key=? AND code_pattern=?",
                    (now, example, zone_key, was))
                if cur.rowcount:
                    log.info("Relaxed %s location pattern", zone_key)

            if not db.execute("SELECT 1 FROM zones WHERE zone_key='COOLROOM'").fetchone():
                order = db.execute(
                    "SELECT COALESCE(MAX(sort_order),0) + 1 AS m FROM zones").fetchone()["m"]
                db.execute(
                    "INSERT INTO zones (zone_key, zone_name_th, code_pattern, allow_multiple, "
                    "example_code, sort_order, free_text, origin_only, no_location, active, from_master) "
                    "VALUES ('COOLROOM','Coolroom','^[A-Z0-9-]+$',0,?,?,0,0,0,1,1)",
                    (COOLROOM_FIRST_ROOM, order))
                log.info("Added COOLROOM zone")

            if not db.execute("SELECT 1 FROM zone_locations WHERE zone_key='COOLROOM'").fetchone():
                db.execute(
                    "INSERT INTO zone_locations (zone_key, code, label, sort_order, active, created_at) "
                    "VALUES ('COOLROOM',?,NULL,1,1,?)", (COOLROOM_FIRST_ROOM, now_iso()))
                log.info("Seeded COOLROOM rooms")
    except Exception as e:
        log.warning("ensure_seed_zones skipped: %s", e)


COST_CENTER_LEN = 5


def valid_cost_center(code: str) -> bool:
    code = (code or "").strip()
    return code.isdigit() and len(code) == COST_CENTER_LEN


def ensure_cost_centers():
    """Fill the cost-centre list from the codes already in use.

    The list is new; the codes are not. Every code an account or a job already
    carries goes in as an established cost centre — not pending, since it was
    in use long before anyone could add one — with its name left blank for
    someone to fill in.
    """
    try:
        with get_db() as db:
            known = {r["code"] for r in db.execute("SELECT code FROM cost_centers").fetchall()}
            found = set()
            for q in ("SELECT DISTINCT cost_center AS c FROM users WHERE cost_center IS NOT NULL",
                      "SELECT DISTINCT cost_center AS c FROM tasks WHERE cost_center IS NOT NULL"):
                for r in db.execute(q).fetchall():
                    if valid_cost_center(r["c"]):
                        found.add(r["c"].strip())
            fresh = sorted(found - known)
            for code in fresh:
                db.execute("INSERT INTO cost_centers (code, name, pending, active, created_at) "
                           "VALUES (?,NULL,0,1,?)", (code, now_iso()))
            if fresh:
                log.info("Seeded %d cost centres already in use", len(fresh))
    except Exception as e:
        log.warning("ensure_cost_centers skipped: %s", e)


def register_cost_center(db, code: str) -> bool:
    """Record a cost centre the moment someone uses one that is not listed.

    Refusing it would stop a job over bookkeeping, so it goes in as pending
    and ops are told; an admin names or retires it afterwards. Returns True
    when this call is what created it.
    """
    code = (code or "").strip()
    if not valid_cost_center(code):
        return False
    if db.execute("SELECT 1 FROM cost_centers WHERE code=?", (code,)).fetchone():
        return False
    db.execute("INSERT INTO cost_centers (code, name, pending, active, created_at) "
               "VALUES (?,NULL,1,1,?)", (code, now_iso()))
    notify_ops(db, f"มี cost center ใหม่ {code} ถูกใช้งาน — ตั้งชื่อหรือปิดได้ในหน้าจัดการระบบ")
    return True


def ensure_columns():
    """Add any missing columns, on either backend, then backfill them once."""
    try:
        with get_db() as db:
            added = []
            for table, col, coltype, _ in NEW_COLUMNS:
                try:
                    if not _column_exists(db, table, col):
                        db.execute(f"ALTER TABLE {table} ADD COLUMN {col} {coltype}")
                        added.append(f"{table}.{col}")
                except Exception as e:
                    log.warning("Could not add column %s.%s: %s", table, col, e)
            if added:
                log.info("Schema updated: added %s", ", ".join(added))

            # vehicle_id used to be NOT NULL. Assigning a driver who has not
            # boarded yet needs it nullable; on SQLite the constraint can only
            # be changed by rebuilding the table, and fresh SQLite databases
            # already get the new definition, so only Postgres needs this.
            if IS_POSTGRES:
                try:
                    db.execute("ALTER TABLE task_assignments ALTER COLUMN vehicle_id DROP NOT NULL")
                except Exception as ex:
                    log.warning("Could not relax task_assignments.vehicle_id: %s", ex)

            # Backfill the role for every assignment made before the column
            # existed: it is simply the type of the vehicle they were in.
            try:
                db.execute(
                    "UPDATE task_assignments SET assigned_vehicle_type = "
                    "(SELECT vehicle_type FROM vehicles WHERE vehicles.id = task_assignments.vehicle_id) "
                    "WHERE assigned_vehicle_type IS NULL AND vehicle_id IS NOT NULL")
            except Exception as ex:
                log.warning("Could not backfill assigned_vehicle_type: %s", ex)

            # Phone numbers were stored as typed/displayed ("081-234-5678").
            # Login and the uniqueness check compare digits only, so every
            # legacy number must be reduced to digits once, or phone login
            # silently fails for all pre-existing accounts.
            legacy = db.execute(
                "SELECT id, contact FROM users WHERE contact IS NOT NULL AND contact != ''"
            ).fetchall()
            fixed = 0
            for r in legacy:
                digits = normalize_phone(r["contact"])
                if digits != r["contact"]:
                    db.execute("UPDATE users SET contact=? WHERE id=?", (digits, r["id"]))
                    fixed += 1
            if fixed:
                log.info("Normalised %d legacy phone number(s) to digits", fixed)

            # Split legacy single-field names into first/last ONCE, at the
            # first space. Names with three words keep everything after the
            # first space as the surname, which is the least-wrong split for
            # Thai names and is editable in Profile afterwards.
            if _column_exists(db, "users", "first_name"):
                rows = db.execute(
                    "SELECT id, full_name FROM users WHERE first_name IS NULL AND full_name IS NOT NULL"
                ).fetchall()
                for r in rows:
                    parts = (r["full_name"] or "").strip().split(" ", 1)
                    first = parts[0] if parts else ""
                    last = parts[1].strip() if len(parts) > 1 else ""
                    db.execute("UPDATE users SET first_name=?, last_name=? WHERE id=?",
                               (first, last, r["id"]))
                if rows:
                    log.info("Backfilled first/last name for %d user(s)", len(rows))

            # Existing jobs keep the number the reports already showed for
            # them: whatever their requester's profile says. Without this,
            # every past job would export a blank cost centre.
            try:
                n = db.execute(
                    "UPDATE tasks SET cost_center = (SELECT cost_center FROM users "
                    "WHERE users.id = tasks.requester_id) WHERE cost_center IS NULL"
                ).rowcount
                if n:
                    log.info("Migrated: cost centre copied onto %s existing task(s)", n)
            except Exception as ex:
                log.warning("Could not backfill tasks.cost_center: %s", ex)

            # Zone behaviour for the zones that already exist. Done here, in
            # the path that runs on every boot for both backends, because
            # seed_master_data only fills an EMPTY zones table — a live
            # database would otherwise get the columns and never the values.
            try:
                for zone_key, flags in (
                    # Empty pallets are collected, never delivered, and the
                    # zone itself is the location.
                    ("EMPTY_PALLET", {"origin_only": 1, "no_location": 1}),
                    # The sort yard is one place; it had no detail to type
                    # anyway, via a hand-written branch that this replaces.
                    ("SORT_YARD", {"no_location": 1}),
                ):
                    row = db.execute("SELECT sort_order FROM zones WHERE zone_key=?", (zone_key,)).fetchone()
                    if not row:
                        continue
                    sets = ", ".join(f"{k}=?" for k in flags)
                    db.execute(f"UPDATE zones SET {sets} WHERE zone_key=?",
                               tuple(flags.values()) + (zone_key,))
                # Staging and empty pallets both landed on sort_order 7 on
                # databases built by the migration path, so their order in the
                # picker differed from a fresh install. Pin them — and pin
                # Coolroom ahead of them, since it is a real storage zone and
                # those two are the catch-alls at the end of the list.
                for zone_key, order in (("COOLROOM", 7), ("STAGING", 8), ("EMPTY_PALLET", 9)):
                    db.execute("UPDATE zones SET sort_order=? WHERE zone_key=?", (order, zone_key))
                log.info("Migrated: zone flags and ordering applied")
            except Exception as ex:
                log.warning("Could not apply zone flags: %s", ex)

            # Invite codes moved from two fixed app_config rows onto the
            # vehicle type they grant. This has to happen here, not in
            # migrate_db(), which only ever runs on SQLite — on an existing
            # Postgres database the codes would otherwise stay NULL and every
            # driver who had been given one would be turned away.
            try:
                moved = []
                for key, type_key, fallback in (("invite_code_rt", "RT", "driver_rt_82640"),
                                                ("invite_code_pe", "PE", "driver_pe_82640")):
                    vt = db.execute("SELECT invite_code FROM vehicle_types WHERE type_key=?",
                                    (type_key,)).fetchone()
                    if not vt or vt["invite_code"]:
                        continue    # type gone, or it already carries its own code
                    old_row = db.execute("SELECT value FROM app_config WHERE key=?", (key,)).fetchone()
                    code = (old_row["value"] if old_row else None) or fallback
                    db.execute("UPDATE vehicle_types SET invite_code=? WHERE type_key=?",
                               (code, type_key))
                    moved.append(f"{type_key}={code}")
                if moved:
                    log.info("Migrated: invite codes moved onto vehicle types (%s)", ", ".join(moved))
            except Exception as ex:
                log.warning("Could not move invite codes onto vehicle types: %s", ex)

    except Exception as e:
        log.warning("ensure_columns skipped: %s", e)


def widen_role_constraint():
    """Teach an EXISTING database that DISPATCH is a role.

    users.role carries CHECK(role IN (...)), and every database built before
    this version lists only USER/DRIVER/ADMIN — so promoting anyone to
    ผู้กระจายงาน fails at the database, not in the app. Like everything in
    ensure_columns(), this has to run on every boot for both backends:
    init_db() applies SCHEMA to a FRESH database only, and migrate_db() never
    runs on Postgres.

    Postgres can simply replace the constraint. SQLite cannot alter one at
    all, so the table is rebuilt by the documented procedure (copy, drop,
    rename) on its own autocommit connection — foreign_keys must be off
    while the old table is dropped, and PRAGMA is ignored inside a
    transaction.
    """
    try:
        if IS_POSTGRES:
            with get_db() as db:
                # Serialize concurrent cold starts: DROP/ADD CONSTRAINT takes
                # an exclusive lock, and two instances arriving together would
                # otherwise both try it.
                db.execute("SELECT pg_advisory_xact_lock(913472002)")
                row = db.execute(
                    "SELECT pg_get_constraintdef(oid) AS def FROM pg_constraint "
                    "WHERE conrelid = 'users'::regclass AND conname = 'users_role_check'"
                ).fetchone()
                if row and "DISPATCH" in (row["def"] or ""):
                    return    # already widened, or a fresh database
                db.execute("ALTER TABLE users DROP CONSTRAINT IF EXISTS users_role_check")
                db.execute("ALTER TABLE users ADD CONSTRAINT users_role_check "
                           "CHECK (role IN ('USER','DRIVER','DISPATCH','ADMIN'))")
            log.info("Migrated: users.role now accepts DISPATCH")
            return

        with get_db() as db:
            row = db.execute(
                "SELECT sql FROM sqlite_master WHERE type='table' AND name='users'").fetchone()
        ddl = (row["sql"] if row else "") or ""
        if not ddl or "CHECK(role" not in ddl.replace("CHECK (role", "CHECK(role"):
            return    # no constraint to widen
        if "DISPATCH" in ddl:
            return    # already current

        body = re.search(
            r"CREATE TABLE IF NOT EXISTS users \((.*?)\n\);", SCHEMA, re.S)
        if not body:
            log.warning("Could not find the users block in SCHEMA; role constraint left as is")
            return

        conn = sqlite3.connect(DB_PATH, isolation_level=None)
        conn.row_factory = sqlite3.Row
        try:
            live = [r["name"] for r in conn.execute("PRAGMA table_info(users)").fetchall()]
            conn.execute("PRAGMA foreign_keys = OFF")
            conn.execute("BEGIN")
            conn.execute(f"CREATE TABLE users_role_migration ({body.group(1)}\n)")
            fresh = [r["name"] for r in conn.execute(
                "PRAGMA table_info(users_role_migration)").fetchall()]
            # Only columns both shapes have: a column the live table never
            # got is left to ensure_columns, which runs before this.
            shared = ", ".join(c for c in fresh if c in live)
            conn.execute(
                f"INSERT INTO users_role_migration ({shared}) SELECT {shared} FROM users")
            conn.execute("DROP TABLE users")
            conn.execute("ALTER TABLE users_role_migration RENAME TO users")
            conn.execute("COMMIT")
        except Exception:
            try:
                conn.execute("ROLLBACK")
            except Exception:
                pass
            raise
        finally:
            try:
                conn.execute("PRAGMA foreign_keys = ON")
                # Dropping the table took its indexes with it; every
                # CREATE here is IF NOT EXISTS, so this only refills those.
                conn.executescript(SCHEMA)
            except Exception as ex:
                log.warning("Could not restore indexes after the role rebuild: %s", ex)
            conn.close()
        log.info("Migrated: users table rebuilt so role accepts DISPATCH")
    except Exception as e:
        log.warning("Could not widen the users.role constraint: %s", e)


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

        if not db.execute("SELECT 1 FROM zones WHERE zone_key='STAGING'").fetchone():
            db.execute(
                "INSERT INTO zones (zone_key, zone_name_th, code_pattern, allow_multiple, example_code, sort_order, free_text) "
                "VALUES ('STAGING','Staging (อื่นๆ)','^.+$',0,'ระบุจุดเอง',7,1)"
            )
            log.info("Migrated database: added STAGING zone")


# ---------------------------------------------------------------------------
# Master Forklift
# ---------------------------------------------------------------------------
# The real fleet. RT is a SINGLE vehicle type — the licence doesn't distinguish
# Selective from Double Deep — with supports_double_deep marking which units can
# reach the deeper racking. Anything not in this list is retired rather than
# deleted, because historical tasks and assignments still point at those rows.
MASTER_FLEET = (
    [(f"WH-PE-{i:03d}", "PE", 0) for i in range(1, 11)]        # WH-PE-001 .. WH-PE-010
    + [(f"WH-{i:02d}", "RT", 0) for i in range(1, 11)]          # WH-01 .. WH-10  (Selective)
    + [(f"WH-{i:02d}", "RT", 1) for i in range(11, 14)]         # WH-11 .. WH-13  (Double Deep)
)
MASTER_CODES = {code for code, _, _ in MASTER_FLEET}


def sync_master_fleet(db):
    """Make the vehicles table match MASTER_FLEET.

    Runs every boot and is idempotent. Deliberately does three separate things
    so a half-applied state can't happen:
      1. insert vehicles that don't exist yet
      2. correct the type / double-deep flag on ones that do
      3. retire everything else (never DELETE — see above)
    A retired vehicle that reappears in the master list is un-retired, so
    fixing a mistake is just a matter of correcting the list.
    """
    existing = {r["code"]: r for r in db.execute("SELECT * FROM vehicles").fetchall()}
    added = corrected = retired = 0

    for code, vtype, dd in MASTER_FLEET:
        row = existing.get(code)
        if row is None:
            db.execute(
                "INSERT INTO vehicles (code, vehicle_type, status, supports_double_deep, retired_at) "
                "VALUES (?,?,?,?,NULL)", (code, vtype, "Available", dd))
            added += 1
        elif (row["vehicle_type"] != vtype
              or (row["supports_double_deep"] or 0) != dd
              or row["retired_at"] is not None):
            db.execute(
                "UPDATE vehicles SET vehicle_type=?, supports_double_deep=?, retired_at=NULL WHERE id=?",
                (vtype, dd, row["id"]))
            corrected += 1

    for code, row in existing.items():
        if code not in MASTER_CODES and row["retired_at"] is None:
            # Don't strand a driver who is checked into a vehicle that just
            # left the fleet — they'd be stuck unable to check out.
            db.execute("UPDATE users SET checked_in_vehicle_id=NULL, driver_status='Not Checked-in' "
                       "WHERE checked_in_vehicle_id=?", (row["id"],))
            # ...and stop auto-boarding anyone into a vehicle that just left
            # the fleet, which would otherwise fail on every app open.
            db.execute("UPDATE driver_licenses SET default_vehicle_id=NULL WHERE default_vehicle_id=?",
                       (row["id"],))
            db.execute("UPDATE vehicles SET retired_at=?, status='Retired' WHERE id=?",
                       (now_iso(), row["id"]))
            retired += 1

    if added or corrected or retired:
        log.info("Master fleet synced: %d added, %d corrected, %d retired", added, corrected, retired)


def seed_master_data(db):
    """Master/reference data (vehicle types, zones, the type<->zone matrix,
    task rules). Split out from user seeding and made idempotent PER TABLE so
    it self-heals: if an earlier broken deploy left these tables empty (tables
    created but seed never completed), this refills whatever is missing on the
    next boot instead of staying empty forever. Safe to call every startup."""
    if not db.execute("SELECT 1 FROM vehicle_types LIMIT 1").fetchone():
        db.execute(
            "INSERT INTO vehicle_types (type_key, type_name_th, description, invite_code) VALUES "
            "('RT','RT - รถยก','ยกสินค้าขึ้นชั้นวาง (บางคันเข้า Double Deep ได้)','driver_rt_82640'),"
            "('PE','PE - รถลาก','ลากพาเลทระยะสั้น','driver_pe_82640')"
        )
        log.info("Seeded master data: vehicle_types")

    if not db.execute("SELECT 1 FROM zones LIMIT 1").fetchone():
        db.execute(
            "INSERT INTO zones (zone_key, zone_name_th, code_pattern, allow_multiple, example_code, "
            "sort_order, free_text, origin_only, no_location) VALUES "
            "('CONCRETE_YARD','ลานปูน','^ลานปูน-[1-7]$',0,'ลานปูน-3',1,0,0,0),"
            "('GATE','ประตู','^ประตู-([1-9]|[1-7][0-9]|8[0-5])$',1,'ประตู-12',2,0,0,0),"
            "('SORT_YARD','ลาน Sort','^ลาน Sort$',0,'ลาน Sort',3,0,0,1),"
            "('MEZZANINE','Mezzanine','^Mezzanine-[1-4]-(ซ้าย \\(Cross dock\\)|ขวา \\(Double Deep\\))$',0,'Mezzanine-2-ซ้าย (Cross dock)',4,0,0,0),"
            "('SELECTIVE_RACK','Selective Rack','" + RACK_PATTERN_SELECTIVE + "',1,'BAA01A1',5,0,0,0),"
            "('DOUBLE_DEEP','Double Deep','" + RACK_PATTERN_DOUBLE_DEEP + "',1,'AAA01A1',6,0,0,0),"
            "('COOLROOM','Coolroom','^[A-Z0-9-]+$',0,'" + COOLROOM_FIRST_ROOM + "',7,0,0,0),"
            "('STAGING','Staging (อื่นๆ)','^.+$',0,'ระบุจุดเอง',8,1,0,0),"
            "('EMPTY_PALLET','พาเลทเปล่า','^.+$',0,'พาเลทเปล่า',9,1,1,1)"
        )
        db.execute("UPDATE zones SET from_master=1 WHERE zone_key='COOLROOM'")
        log.info("Seeded master data: zones")


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

# Task types (Transfer / Putaway / ...) were removed from the product. The
# column is kept and filled with a constant rather than dropped: dropping a
# NOT NULL column on a live Postgres database mid-deploy is a far bigger risk
# than an unused field, and old rows keep their original value for reporting.
TASK_TYPE_PLACEHOLDER = "-"


def infer_task_type(db, from_zone: str, to_zone: str) -> str:
    return TASK_TYPE_PLACEHOLDER


def validate_location_string(db, zone_key: str, value: str, *, is_destination: bool = False) -> str:
    zone = db.execute("SELECT * FROM zones WHERE zone_key=?", (zone_key,)).fetchone()
    if not zone:
        raise HTTPException(400, f"ไม่พบโซน {zone_key}")
    if is_destination and zone["origin_only"]:
        raise HTTPException(400, f"โซน {zone['zone_name_th']} เลือกเป็นจุดปลายทางไม่ได้ "
                                 f"(ใช้เป็นจุดต้นทางเท่านั้น)")
    if not zone["active"]:
        raise HTTPException(400, f"โซน {zone['zone_name_th']} ถูกปิดใช้งานแล้ว กรุณาเลือกโซนอื่น")
    value = (value or "").strip()
    # The zone itself is the location, so there is nothing to type and nothing
    # to validate: the name goes in as the location.
    if zone["no_location"]:
        return zone["zone_name_th"]
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
    # A list-backed zone (Coolroom and its rooms) is checked against its own
    # places rather than a pattern. A pattern cannot say WHICH rooms exist,
    # and a room that has been retired has to stop being accepted even though
    # its shape is still perfectly valid.
    if zone["from_master"]:
        known = {r["code"].upper(): r["code"] for r in db.execute(
            "SELECT code FROM zone_locations WHERE zone_key=? AND active=1",
            (zone_key,)).fetchall()}
        chosen = []
        for p in parts:
            hit = known.get(p.upper())
            if not hit:
                raise HTTPException(400, f"ไม่พบจุด '{p}' ในโซน {zone['zone_name_th']} "
                                         f"— เลือกจากรายการที่มี หรือให้แอดมินเพิ่มจุดนี้ก่อน")
            chosen.append(hit)
        return ",".join(chosen)
    pattern = re.compile(zone["code_pattern"], re.IGNORECASE)
    for p in parts:
        if not pattern.match(p):
            raise HTTPException(400, f"รูปแบบตำแหน่งไม่ถูกต้อง: '{p}' (ตัวอย่างที่ถูกต้อง: {zone['example_code']})")
    return ",".join(parts)


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


def person_label(db, user_id) -> str:
    """How a person is named in the task chat: 'คุณ' + first name, falling back
    to the full name and then the employee id."""
    u = db.execute("SELECT first_name, full_name, employee_id, role FROM users WHERE id=?",
                   (user_id,)).fetchone()
    if not u:
        return "ระบบ"
    name = (u["first_name"] or "").strip() or (u["full_name"] or "").strip() or u["employee_id"]
    if u["role"] == "ADMIN":
        return "แอดมิน"
    if u["role"] == "DISPATCH":
        return f"ผู้กระจายงาน {name}"
    return f"คุณ{name}"


def person_name(row) -> str:
    """The name to print for a person in a notification.

    Employee IDs are what the tables join on, not what anyone on the floor
    recognises: "20100001 รับงาน RT-261026-004" costs the reader a lookup
    before it means anything. The full name is used wherever the row carries
    one, and the employee ID stays as the fallback so a row without a name
    still says who, rather than nothing at all.

    Takes a row the caller already holds, so naming someone costs no extra
    query. person_label() is the chat's version and stays as it is: there a
    person is addressed ("คุณสมชาย", "แอดมิน"), while a notification has to
    name them.
    """
    if row is None:
        return "ไม่ทราบชื่อ"

    def field(key):
        try:
            v = row[key]
        except (KeyError, IndexError, TypeError):
            return ""
        return str(v).strip() if v is not None else ""

    return field("full_name") or field("employee_id") or "ไม่ทราบชื่อ"


def post_system_message(db, task_id: int, actor_id: int, text: str):
    """An automatic line in the task chat. Deliberately does NOT notify anyone:
    every event that posts one already sends its own notification, and a second
    "new chat message" ping for the same event would just be noise."""
    db.execute(
        "INSERT INTO task_messages (task_id, sender_id, message, kind, created_at) VALUES (?,?,?,?,?)",
        (task_id, actor_id, text, "system", now_iso()),
    )


_PAUSE_SIDE_TH = {"user": "ฝั่งผู้ขอ", "vehicle": "ฝั่งรถ"}


def _status_change_text(db, from_status, to_status, actor_id, note):
    who = person_label(db, actor_id)
    if to_status == "In Progress" and from_status == "Pause":
        return f"{who} แก้ไขเหตุขัดข้องแล้ว — งานดำเนินการต่อ"
    if to_status == "In Progress":
        return f"{who} เริ่มงาน"
    if to_status == "Pause":
        text = f"{who} แจ้งปัญหา"
        if note:
            m = re.match(r"\[(\w+)\]\s*(.*)", note)
            if m:
                side, reason = _PAUSE_SIDE_TH.get(m.group(1), m.group(1)), m.group(2)
                text += f" ({side}): {reason}"
            else:
                text += f": {note}"
        return text
    if to_status == "Completed":
        return f"งานเสร็จสิ้นเรียบร้อย — ปิดงานโดย{who}"
    if to_status == "Cancelled":
        return f"{who} ยกเลิกงาน" + (f": {note}" if note else "")
    if to_status == "Waiting":
        return "ไม่มีคนขับเหลืออยู่ในงาน — งานกลับเข้าคิวรอมอบหมาย"
    return None


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
    # Every status change goes through here, so the chat timeline can't miss one.
    text = _status_change_text(db, from_status, to_status, changed_by, note)
    if text:
        post_system_message(db, task_row["id"], changed_by, text)


def notify(db, user_id: int, message: str, task_id: Optional[int] = None, kind: str = "task"):
    db.execute(
        "INSERT INTO notifications (user_id, message, task_id, kind, created_at) VALUES (?,?,?,?,?)",
        (user_id, message, task_id, kind, now_iso()),
    )
    # The kind rides along so a tapped chat notification opens the chat and a
    # job notification opens the job, the same split the in-app list uses.
    push_url = f"/?task={task_id}&kind={kind}" if (kind in ("task", "chat") and task_id) else "/"
    # Chat messages about one job collapse into one line; anything else about
    # that job replaces the previous line about it; everything else stands on
    # its own. See send_web_push_to_user() for why this matters.
    if task_id:
        tag = f"scdc-chat-{task_id}" if kind == "chat" else f"scdc-task-{task_id}"
    else:
        tag = f"scdc-{kind}-{secrets.token_hex(4)}"
    queue = _push_queue_var.get()
    if queue is None:
        # Outside an HTTP request (startup tasks, scripts) — send inline.
        send_web_push_to_user(db, user_id, "SCDC", message, push_url, tag)
    else:
        # Inside a request — hand off to the post-response flush so the caller
        # isn't blocked on push delivery.
        queue.append((user_id, "SCDC", message, push_url, tag))


def notify_ops(db, message: str, task_id: Optional[int] = None, kind: str = "task"):
    """Everyone who runs the floor — admins AND dispatchers.

    This used to be notify_admins() and selected role='ADMIN' only, which left
    the one role whose whole job is dispatching with no notification of a new
    job, a job waiting too long, or a breakdown. Deleted accounts are excluded:
    without that filter, someone who has left the company keeps receiving job
    notifications on their own phone, because a soft-deleted row still has live
    push subscriptions.
    """
    for a in db.execute(
        "SELECT id FROM users WHERE role IN ('ADMIN','DISPATCH') AND deleted_at IS NULL"
    ).fetchall():
        notify(db, a["id"], message, task_id, kind)



def audit(db, actor_id: Optional[int], action: str, details: str = ""):
    db.execute(
        "INSERT INTO audit_logs (actor_id, action, details, created_at) VALUES (?,?,?,?)",
        (actor_id, action, details, now_iso()),
    )


def free_up_driver_and_vehicle(db, driver_id: Optional[int], vehicle_id: Optional[int]):
    if driver_id:
        # recompute rather than assume free: with queues, a driver coming off
        # one job may still be holding others, and flattening them to
        # "Available" would offer them to a second job they can't take.
        recompute_driver_status(db, driver_id)
    if vehicle_id:
        row = db.execute("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
        if row and row["status"] != "Breakdown":
            db.execute("UPDATE vehicles SET status='Available' WHERE id=?", (vehicle_id,))


# ---------------------------------------------------------------------------
# Job queues
#
# A dispatcher can now hand a driver several jobs in one go instead of waiting
# by the screen to deal them out one at a time. Exactly one of a driver's
# Active assignments is the job in hand (is_active_for_driver = 1); the others
# are their queue. Finishing the job in hand promotes the next one on its own,
# so the driver never has to come back to the app to be given work that was
# already theirs.
#
# The flag lives on the assignment, not on the task, because a task can carry
# a crew: one member may be working it while another still has it queued
# behind something else, and a single task-level flag could not express that.
# ---------------------------------------------------------------------------

def driver_active_assignment(db, driver_id):
    """The assignment for the job this driver has in hand, or None."""
    return db.execute(
        "SELECT ta.* FROM task_assignments ta JOIN tasks t ON t.id = ta.task_id "
        "WHERE ta.driver_id=? AND ta.assignment_status='Active' AND ta.is_active_for_driver=1 "
        "AND t.status NOT IN ('Completed','Cancelled') LIMIT 1",
        (driver_id,),
    ).fetchone()


def driver_queue(db, driver_id):
    """This driver's queued assignments, the one waiting longest first."""
    return db.execute(
        "SELECT ta.*, t.task_code FROM task_assignments ta JOIN tasks t ON t.id = ta.task_id "
        "WHERE ta.driver_id=? AND ta.assignment_status='Active' AND ta.is_active_for_driver=0 "
        "AND t.status NOT IN ('Completed','Cancelled') ORDER BY t.created_at, ta.assigned_at",
        (driver_id,),
    ).fetchall()


def task_has_active_worker(db, task_id) -> bool:
    return bool(db.execute(
        "SELECT 1 FROM task_assignments WHERE task_id=? AND assignment_status='Active' "
        "AND is_active_for_driver=1", (task_id,),
    ).fetchone())


def recompute_driver_status(db, driver_id):
    """Set a driver's status from what they are actually holding.

    Four states, in order of precedence: a reported breakdown outranks
    everything; working a job is Busy; holding only queued work is Assigned
    (spoken for, so no job offers them to anyone else); otherwise free, which
    is Available with a vehicle and Not Checked-in without one.
    """
    row = db.execute("SELECT driver_status, checked_in_vehicle_id FROM users WHERE id=?",
                     (driver_id,)).fetchone()
    if not row or row["driver_status"] == "Breakdown":
        return
    active = driver_active_assignment(db, driver_id)
    if active and active["vehicle_id"]:
        status = "Busy"
    elif active or driver_queue(db, driver_id):
        # Spoken for, but not out driving yet: either the job in hand is still
        # waiting on a vehicle, or they are only holding queued work. Either
        # way no other job may be offered to them.
        status = "Assigned"
    else:
        status = "Available" if row["checked_in_vehicle_id"] else "Not Checked-in"
    db.execute("UPDATE users SET driver_status=? WHERE id=?", (status, driver_id))


def start_assignment(db, assignment, actor_id, *, announce=True):
    """Make a queued assignment the job this driver has in hand."""
    db.execute("UPDATE task_assignments SET is_active_for_driver=1 WHERE id=?", (assignment["id"],))
    row = db.execute("SELECT * FROM tasks WHERE id=?", (assignment["task_id"],)).fetchone()
    stamp = now_iso()
    if row["status"] == "Assigned":
        db.execute("UPDATE tasks SET status='In Progress', started_at=COALESCE(started_at,?) WHERE id=?",
                   (stamp, row["id"]))
        db.execute(
            "INSERT INTO task_status_history (task_id, from_status, to_status, changed_by, changed_at) "
            "VALUES (?,?,?,?,?)", (row["id"], "Assigned", "In Progress", actor_id, stamp),
        )
    db.execute("UPDATE tasks SET current_driver_id=?, current_vehicle_id=? WHERE id=?",
               (assignment["driver_id"], assignment["vehicle_id"], row["id"]))
    # The vehicle is back out on a job. Finishing the previous one released it
    # a moment ago, so without this the fleet page and the dashboard would
    # show a vehicle as free while the driver is away using it.
    if assignment["vehicle_id"]:
        db.execute("UPDATE vehicles SET status='Busy' WHERE id=? AND status!='Breakdown'",
                   (assignment["vehicle_id"],))
    recompute_driver_status(db, assignment["driver_id"])
    if announce:
        post_system_message(db, row["id"], assignment["driver_id"],
                            f"{person_label(db, assignment['driver_id'])} เริ่มงานนี้แล้ว")
    return row


def promote_next_in_queue(db, driver_id, actor_id):
    """Start the next queued job for a driver who just freed up.

    This is the whole point of queues: the dispatcher hands out a batch, and
    each job begins by itself as the one before it ends, with no one waiting
    in the app to press anything.
    """
    if driver_active_assignment(db, driver_id):
        return None
    q = driver_queue(db, driver_id)
    if not q:
        return None
    nxt = q[0]
    row = start_assignment(db, nxt, actor_id, announce=False)
    post_system_message(db, row["id"], driver_id,
                        f"{person_label(db, driver_id)} เริ่มงานนี้ต่อจากงานที่เพิ่งเสร็จ")
    notify(db, driver_id, f"เริ่มงานถัดไปในคิว: {row['task_code']} "
                          f"{row['from_location']} → {row['to_location']}", row["id"])
    notify(db, row["requester_id"], f"งาน {row['task_code']} เริ่มดำเนินการแล้ว", row["id"])
    return row


def check_waiting_too_long(db):
    cutoff = (datetime.utcnow() - timedelta(minutes=WAITING_TOO_LONG_MINUTES)).isoformat(timespec="seconds") + "Z"
    rows = db.execute(
        "SELECT * FROM tasks WHERE status='Waiting' AND waiting_alert_sent=0 AND created_at<?", (cutoff,)
    ).fetchall()
    for t in rows:
        notify_ops(db, f"งาน {t['task_code']} รอนานเกิน {WAITING_TOO_LONG_MINUTES} นาทีแล้ว", t["id"])
        db.execute("UPDATE tasks SET waiting_alert_sent=1 WHERE id=?", (t["id"],))


def task_chat_allowed(db, task_row, user) -> bool:
    # Dispatchers coordinate every job, so they are in every job's chat.
    if user["role"] in ("ADMIN", "DISPATCH"):
        return True
    if task_row["requester_id"] == user["id"]:
        return True
    # A colleague booked to the same cost centre can see the job in their unit
    # view, so they can ask about it there too — otherwise the job appears in
    # their list with a chat button that always refuses them.
    if (user["role"] == "USER" and task_row["cost_center"]
            and (user["cost_center"] or "").strip() == task_row["cost_center"]):
        return True
    ever_driver = db.execute(
        "SELECT 1 FROM task_assignments WHERE task_id=? AND driver_id=?", (task_row["id"], user["id"])
    ).fetchone()
    return bool(ever_driver)


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

class LoginBody(BaseModel):
    # Accepts an employee ID, a phone number, or a @central.co.th email.
    # The field keeps its old name so existing clients keep working.
    employee_id: str
    password: str


class RegisterBody(BaseModel):
    employee_id: str
    email: Optional[str] = None
    full_name: str
    first_name: Optional[str] = None
    last_name: Optional[str] = None
    password: str
    cost_center: str = Field(..., min_length=5, max_length=5)
    contact: str
    # Optional invite code. A valid one registers the person as a driver with
    # the matching licence straight away; an invalid one is a hard error rather
    # than being ignored, so a typo can't quietly create a plain user who then
    # wonders why they have no driver menu.
    invite_code: Optional[str] = None


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
        raise HTTPException(401, "ไม่พบ token กรุณาเข้าสู่ระบบใหม่")
    token = authorization.removeprefix("Bearer ").strip()
    with get_db() as db:
        row = db.execute(
            "SELECT users.* FROM sessions JOIN users ON users.id = sessions.user_id WHERE token=?",
            (token,),
        ).fetchone()
    if not row or row["deleted_at"]:
        raise HTTPException(401, "เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่")
    return dict(row)


# ---------------------------------------------------------------------------
# Roles
#
# DISPATCH is an admin minus "จัดการระบบ": they run the floor — hand out work,
# manage staff and vehicles, answer every chat — but they do not change the
# system's own configuration (zones, vehicle types, rack layout, settings) and
# they cannot read the audit log.
#
# The one rule that is not about screens, and the reason it is written here
# rather than per endpoint: a dispatcher must never be able to reach an ADMIN
# account or hand out the ADMIN role. Without that, "no system management" is
# decoration — they would simply reset an admin's password and sign in as them.
# ---------------------------------------------------------------------------

ROLES = ("USER", "DRIVER", "DISPATCH", "ADMIN")

# Who can be put on a job.
#
# Dispatchers run the floor, and one who also drives is crew like anybody
# else — but only once they hold a licence. Without that condition every
# dispatcher who never touches a forklift would show up in the driver list,
# in "คนขับว่าง", and in the assignment suggestions.
CREW_ROLES = ("DRIVER", "DISPATCH")

# Who each role may set someone else's role to.
ROLES_ASSIGNABLE_BY = {
    "ADMIN": ("USER", "DRIVER", "DISPATCH", "ADMIN"),
    "DISPATCH": ("USER", "DRIVER", "DISPATCH"),
}

# Anything that acts on another person's account. Only an admin may aim it at
# an admin.
def require_can_manage(actor, target_row):
    if target_row["role"] == "ADMIN" and actor["role"] != "ADMIN":
        raise HTTPException(403, "บัญชีนี้เป็นแอดมิน เฉพาะแอดมินเท่านั้นที่จัดการได้")


def require_can_grant_role(actor, role):
    allowed = ROLES_ASSIGNABLE_BY.get(actor["role"], ())
    if role not in allowed:
        raise HTTPException(
            403, f"คุณตั้งสิทธิ์เป็น {role} ให้ใครไม่ได้ "
                 f"(ตั้งได้: {', '.join(allowed) or 'ไม่มี'})")


def _guard_target(actor, user_id):
    """require_can_manage for an endpoint that only has the target's id."""
    with get_db() as db:
        row = db.execute("SELECT role FROM users WHERE id=?", (user_id,)).fetchone()
    if not row:
        raise HTTPException(404, "ไม่พบผู้ใช้")
    require_can_manage(actor, row)


def require_role(*roles):
    def checker(user=Depends(current_user)):
        if user["role"] not in roles:
            raise HTTPException(403, "คุณไม่มีสิทธิ์เข้าถึงส่วนนี้")
        return user
    return checker


# Shorthand for "the floor", used by everything a dispatcher shares with an
# admin — spelled out once so a new endpoint cannot quietly land on the wrong
# side of the line.
def require_ops():
    return require_role("ADMIN", "DISPATCH")


def require_crew():
    """Anyone who may be on a job: a driver, or a dispatcher who also drives."""
    return require_role(*CREW_ROLES)


@app.post("/auth/login")
def login(body: LoginBody, user_agent: Optional[str] = Header(None)):
    identifier = body.employee_id.strip()
    with get_db() as db:
        # One field, three ways in. Phone is matched on digits only so the
        # formatting someone types never decides whether they can log in.
        phone = normalize_phone(identifier)
        # Only add the phone clause when there's actually a number to match.
        # (An earlier version passed a NUL character as a "matches nothing"
        # placeholder; SQLite accepts that, but PostgreSQL rejects any string
        # containing NUL outright, so every login without digits — including
        # "admin" — crashed with a 500 in production.)
        sql = ("SELECT * FROM users WHERE deleted_at IS NULL AND "
               "(employee_id=? OR LOWER(email)=LOWER(?)")
        params = [identifier, identifier]
        if phone:
            sql += " OR contact=?"
            params.append(phone)
        sql += ")"
        row = db.execute(sql, tuple(params)).fetchone()
        if not row:
            # 404 specifically means "no such account" — the frontend uses this
            # to offer registration with what they typed already filled in.
            # A wrong PASSWORD returns 401 instead and tells them to contact an
            # admin, so the two cases can never be confused.
            raise HTTPException(404, "ยังไม่มีบัญชีนี้ โปรดสมัครสมาชิก")
        if row["password_hash"] == PENDING_PASSWORD:
            raise HTTPException(428, "เข้าสู่ระบบครั้งแรก กรุณาตั้งรหัสผ่าน")
        if not verify_password(body.password, row["password_hash"]):
            raise HTTPException(401, "รหัสผ่านไม่ถูกต้อง หากลืมรหัสโปรดติดต่อ Admin")
        # Signing in no longer touches the vehicle. It used to clear
        # checked_in_vehicle_id on every login, which is what made drivers
        # re-pick a vehicle so often: anything that ended a session — a login
        # on a second device, cleared site data — cost them their vehicle even
        # though they were still sitting in it. A vehicle now changes hands
        # only when somebody actually takes it (see /vehicles/check-in).
        #
        # Several devices may be signed in at once, too. The old rule kept one
        # session per account, so a driver using both their phone and a shared
        # tablet knocked themselves out all day and retyped their password
        # every time. Profile lists the signed-in devices and can sign them
        # all out, which is the control that replaces that rule.
        token = secrets.token_urlsafe(32)
        db.execute(
            "INSERT INTO sessions (token, user_id, created_at, user_agent) VALUES (?,?,?,?)",
            (token, row["id"], now_iso(), (user_agent or "")[:300] or None),
        )
        # Housekeeping, done here because there is no scheduler: tokens this
        # old belong to devices nobody signed out from months ago.
        db.execute("DELETE FROM sessions WHERE user_id=? AND created_at < ?",
                   (row["id"], iso_days_ago(SESSION_MAX_AGE_DAYS)))
        return {"token": token, "role": row["role"], "employee_id": row["employee_id"]}


@app.get("/auth/sessions")
def list_my_sessions(authorization: Optional[str] = Header(None), user=Depends(current_user)):
    """Devices this account is signed in on, newest first."""
    mine = (authorization or "").removeprefix("Bearer ").strip()
    with get_db() as db:
        rows = db.execute(
            "SELECT token, created_at, user_agent FROM sessions WHERE user_id=? ORDER BY created_at DESC",
            (user["id"],),
        ).fetchall()
    return [{
        "created_at": r["created_at"],
        "device": describe_user_agent(r["user_agent"]),
        "current": r["token"] == mine,
    } for r in rows]


@app.post("/auth/logout")
def logout(authorization: Optional[str] = Header(None), user=Depends(current_user)):
    """Sign this device out. Without this the token stayed valid forever: the
    frontend only dropped its copy, and nothing on the server ever expired
    it."""
    token = (authorization or "").removeprefix("Bearer ").strip()
    with get_db() as db:
        db.execute("DELETE FROM sessions WHERE token=? AND user_id=?", (token, user["id"]))
    return {"ok": True}


@app.post("/auth/logout-all")
def logout_all(user=Depends(current_user)):
    """Sign out everywhere, this device included — what to reach for when a
    phone goes missing or a shared tablet was left signed in."""
    with get_db() as db:
        n = db.execute("SELECT COUNT(*) c FROM sessions WHERE user_id=?", (user["id"],)).fetchone()["c"]
        db.execute("DELETE FROM sessions WHERE user_id=?", (user["id"],))
        # The reason to press this is usually a phone that has gone missing, so
        # the notifications have to stop going to it too.
        db.execute("DELETE FROM push_subscriptions WHERE user_id=?", (user["id"],))
    return {"ok": True, "signed_out": n}


def find_deleted_account(db, employee_id):
    """Returns the soft-deleted row for this employee id, or None.
    Raises if a LIVE account already holds it."""
    row = db.execute("SELECT * FROM users WHERE employee_id=?", (employee_id,)).fetchone()
    if row and not row["deleted_at"]:
        return "live"
    return row    # a deleted row, or None


def reactivate_account(db, row, *, role, full_name, password_hash, email,
                       cost_center=None, contact=None):
    """Bring a soft-deleted account back. Everything operational is reset —
    an account that was deleted mid-shift could still be pointing at a vehicle
    or carrying a stale Busy status, and inheriting that would wedge the new
    session."""
    db.execute(
        "UPDATE users SET deleted_at=NULL, role=?, full_name=?, password_hash=?, email=?, "
        "cost_center=?, contact=?, driver_status=?, checked_in_vehicle_id=NULL WHERE id=?",
        (role, full_name, password_hash, email, cost_center, contact,
         "Not Checked-in" if role == "DRIVER" else None, row["id"]),
    )
    # Any session from before the deletion is long gone (delete clears them),
    # but clear again in case the row was revived by another path.
    db.execute("DELETE FROM sessions WHERE user_id=?", (row["id"],))
    return row["id"]


@app.post("/auth/register")
def register(body: RegisterBody):
    with get_db() as db:
        existing = find_deleted_account(db, body.employee_id)
        if existing == "live":
            raise HTTPException(400, "รหัสพนักงานนี้มีผู้ใช้งานแล้ว")
        email = validate_and_normalize_email(db, body.email)
        if not body.cost_center.isdigit():
            raise HTTPException(400, "Cost center ต้องเป็นตัวเลข 5 หลัก")
        if not body.full_name.strip():
            raise HTTPException(400, "กรุณากรอกชื่อ-นามสกุล")
        contact = validate_and_normalize_contact(db, body.contact)
        if len(body.password) < MIN_PASSWORD_LENGTH:
            raise HTTPException(400, f"รหัสผ่านต้องมีอย่างน้อย {MIN_PASSWORD_LENGTH} ตัวอักษร")

        # The invite code decides the role. Resolved before anything is
        # written, so a bad code costs nothing.
        role, licence = "USER", None
        if body.invite_code:
            # Every vehicle type carries its own code, so a type added later
            # can hand out its own licence without a code change in here.
            code = body.invite_code.strip()
            match = db.execute(
                "SELECT type_key FROM vehicle_types WHERE invite_code IS NOT NULL "
                "AND invite_code != '' AND invite_code=?", (code,),
            ).fetchone() if code else None
            if not match:
                raise HTTPException(400, "รหัสเชิญไม่ถูกต้อง")
            role, licence = "DRIVER", match["type_key"]

        full_name = body.full_name.strip()
        first = (body.first_name or "").strip() or full_name.split(" ", 1)[0]
        last = (body.last_name or "").strip() or (
            full_name.split(" ", 1)[1].strip() if " " in full_name else "")

        if existing:
            user_id = reactivate_account(
                db, existing, role=role, full_name=full_name,
                password_hash=hash_password(body.password), email=email,
                cost_center=body.cost_center, contact=contact)
        else:
            cur = db.execute(
                "INSERT INTO users (employee_id, email, full_name, first_name, last_name, "
                "password_hash, role, driver_status, cost_center, contact, created_at) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?) RETURNING id",
                (body.employee_id, email, full_name, first, last,
                 hash_password(body.password), role,
                 "Not Checked-in" if role == "DRIVER" else None,
                 body.cost_center, contact, now_iso()),
            )
            user_id = cur.fetchone()["id"]
        db.execute("UPDATE users SET first_name=?, last_name=? WHERE id=?", (first, last, user_id))
        register_cost_center(db, body.cost_center)

        if licence:
            # Permanent licence (no expiry), as specified for invite-code drivers.
            db.execute("INSERT INTO driver_licenses (driver_id, vehicle_type) VALUES (?,?) "
                       "ON CONFLICT DO NOTHING", (user_id, licence))
            # Anyone who knows the code can make themselves a driver, so every
            # use is announced instead of happening silently.
            notify_ops(db, f"{full_name or body.employee_id} สมัครเป็นคนขับด้วยรหัสเชิญ (ใบอนุญาต {licence})")
            audit(db, user_id, "register_with_invite_code", f"{body.employee_id} -> DRIVER {licence}")
        return {"ok": True, "reactivated": bool(existing), "role": role, "license": licence}


@app.post("/auth/set-password")
def set_password(body: SetPasswordBody):
    if len(body.new_password) < MIN_PASSWORD_LENGTH:
        raise HTTPException(400, f"รหัสผ่านต้องมีอย่างน้อย {MIN_PASSWORD_LENGTH} ตัวอักษร")
    with get_db() as db:
        row = db.execute("SELECT * FROM users WHERE employee_id=?", (body.employee_id,)).fetchone()
        if not row or row["password_hash"] != PENDING_PASSWORD:
            raise HTTPException(400, "บัญชีนี้ตั้งรหัสผ่านไว้แล้ว")
        db.execute("UPDATE users SET password_hash=? WHERE id=?", (hash_password(body.new_password), row["id"]))
        return {"ok": True}


@app.post("/auth/change-password")
def change_password(body: ChangePasswordBody, user=Depends(current_user)):
    if len(body.new_password) < MIN_PASSWORD_LENGTH:
        raise HTTPException(400, f"รหัสผ่านต้องมีอย่างน้อย {MIN_PASSWORD_LENGTH} ตัวอักษร")
    with get_db() as db:
        row = db.execute("SELECT * FROM users WHERE id=?", (user["id"],)).fetchone()
        if not verify_password(body.old_password, row["password_hash"]):
            raise HTTPException(400, "รหัสผ่านปัจจุบันไม่ถูกต้อง")
        db.execute("UPDATE users SET password_hash=? WHERE id=?", (hash_password(body.new_password), user["id"]))
        return {"ok": True}


@app.get("/me")
def me(user=Depends(current_user)):
    out = {"employee_id": user["employee_id"], "role": user["role"], "id": user["id"],
           "full_name": user["full_name"], "contact": user["contact"], "cost_center": user["cost_center"],
           "email": user["email"],
           "driver_status": user["driver_status"], "checked_in_vehicle_id": user["checked_in_vehicle_id"]}
    if user["role"] in CREW_ROLES:
        with get_db() as db:
            out["home_zones"] = read_home_zones(db, user["id"])
        # Read-only on their own profile; only an admin changes these. A
        # dispatcher gets them too: holding one is what makes them crew.
        with get_db() as db:
            today = datetime.utcnow().strftime("%Y-%m-%d")
            out["licenses"] = [
                {"vehicle_type": r["vehicle_type"], "expiry_date": r["expiry_date"],
                 "valid": (not r["expiry_date"]) or r["expiry_date"] >= today,
                 # Their usual vehicle for this type, so the profile can state
                 # it without the client having to join against the fleet.
                 "default_vehicle_code": r["default_code"]}
                for r in db.execute(
                    "SELECT dl.vehicle_type, dl.expiry_date, v.code AS default_code "
                    "FROM driver_licenses dl LEFT JOIN vehicles v "
                    "ON v.id = dl.default_vehicle_id AND v.retired_at IS NULL "
                    "WHERE dl.driver_id=? ORDER BY dl.vehicle_type", (user["id"],)).fetchall()
            ]
    return out


@app.put("/me")
def update_me(body: ProfileUpdateBody, user=Depends(current_user)):
    """Everyone can edit their own name, contact, email and cost centre.
    Employee ID and role are never editable here.

    Cost centre used to be USER-only. Now that an admin or a dispatcher raises
    jobs too — and a job takes the creator's cost centre as its default — they
    need somewhere to set their own, or they retype it on every job."""
    fields = {}
    if body.full_name is not None:
        if not body.full_name.strip():
            raise HTTPException(400, "ชื่อ-นามสกุลต้องไม่เว้นว่าง")
        fields["full_name"] = body.full_name.strip()
    # contact is validated inside the DB block below — it needs a uniqueness
    # check now that phone numbers are a login identifier.
    if body.cost_center is not None:
        cc = body.cost_center.strip()
        # Clearing it is allowed: a driver has no cost centre of their own.
        if cc and (not cc.isdigit() or len(cc) != 5):
            raise HTTPException(400, "Cost center ต้องเป็นตัวเลข 5 หลัก")
        fields["cost_center"] = cc or None
    if not fields and body.email is None and body.contact is None:
        return {"ok": True}
    with get_db() as db:
        if body.email is not None:
            fields["email"] = validate_and_normalize_email(db, body.email, exclude_user_id=user["id"])
        if body.contact is not None:
            # Digits only, the same form register and login use. Storing the
            # dashed display form here is what made phone login fail.
            fields["contact"] = validate_and_normalize_contact(db, body.contact, exclude_user_id=user["id"])
        if not fields:
            return {"ok": True}
        if fields.get("cost_center"):
            register_cost_center(db, fields["cost_center"])
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
def create_driver(body: NewDriverBody, user=Depends(require_ops())):
    with get_db() as db:
        existing = find_deleted_account(db, body.employee_id)
        if existing == "live":
            raise HTTPException(400, "รหัสพนักงานนี้มีอยู่ในระบบแล้ว")
        if existing is not None and existing != "live":
            # Bringing back a removed account: if it used to be an admin, this
            # would quietly turn it into a driver, so only an admin may.
            require_can_manage(user, existing)
        email = validate_and_normalize_email(db, body.email)
        reactivated = False
        if existing:
            driver_id = reactivate_account(
                db, existing, role="DRIVER", full_name=body.full_name,
                password_hash=PENDING_PASSWORD, email=email)
            reactivated = True
        else:
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
        audit(db, user["id"], "reactivate_driver" if reactivated else "create_driver", body.employee_id)
        return {"ok": True, "reactivated": reactivated}


@app.get("/admin/users")
def list_all_users(user=Depends(require_ops())):
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
def admin_update_user(user_id: int, body: UserAdminUpdateBody, user=Depends(require_ops())):
    with get_db() as db:
        row = db.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบผู้ใช้")
        fields = body.dict(exclude_unset=True, exclude={"email"})
        # Personal details belong to the person. An admin can see them, change
        # someone's role, reset their password and remove the account — but
        # name, phone, email and cost centre are edited by their owner only,
        # so there is one source of truth for who someone is.
        personal = [k for k in ("full_name", "contact", "cost_center") if k in fields]
        if personal or body.email is not None:
            raise HTTPException(
                403, "แอดมินแก้ไขข้อมูลส่วนตัวของผู้ใช้ไม่ได้ — เจ้าของบัญชีแก้ไขเองในเมนูโปรไฟล์")
        require_can_manage(user, row)
        if "role" in fields:
            if fields["role"] not in ROLES:
                raise HTTPException(400, "role ไม่ถูกต้อง")
            require_can_grant_role(user, fields["role"])
            # Moving someone off the floor mid-job would leave the job with a
            # driver the system no longer treats as one.
            if row["role"] in CREW_ROLES and fields["role"] not in CREW_ROLES:
                active = db.execute(
                    "SELECT tasks.task_code FROM task_assignments "
                    "JOIN tasks ON tasks.id = task_assignments.task_id "
                    "WHERE task_assignments.driver_id=? AND task_assignments.assignment_status='Active' "
                    "AND tasks.status NOT IN ('Completed','Cancelled')", (user_id,)).fetchall()
                if active:
                    codes = ", ".join(a["task_code"] for a in active)
                    raise HTTPException(
                        400, f"เปลี่ยนสิทธิ์ไม่ได้ เพราะยังมีงานที่ทำอยู่ ({codes}) — "
                             "กรุณาจบงานหรือมอบหมายใหม่ก่อน")
        if not fields:
            return {"ok": True}
        set_clause = ", ".join(f"{k}=?" for k in fields.keys())
        db.execute(f"UPDATE users SET {set_clause} WHERE id=?", list(fields.values()) + [user_id])
        new_role = fields.get("role")
        if new_role in CREW_ROLES and row["role"] != new_role and row["driver_status"] is None:
            db.execute("UPDATE users SET driver_status='Not Checked-in' WHERE id=?", (user_id,))
        # Someone moved off the floor should not keep a vehicle or a status
        # that makes them look available for work.
        if new_role and new_role not in CREW_ROLES and row["role"] in CREW_ROLES:
            db.execute("UPDATE vehicles SET status='Available' WHERE id=? AND status!='Breakdown'",
                       (row["checked_in_vehicle_id"],))
            db.execute("UPDATE users SET driver_status=NULL, checked_in_vehicle_id=NULL WHERE id=?",
                       (user_id,))
        audit(db, user["id"], "admin_update_user", f"{row['employee_id']}: {fields}")
        return {"ok": True}


@app.delete("/admin/users/{user_id}")
def admin_delete_user(user_id: int, user=Depends(require_ops())):
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
        require_can_manage(user, row)
        if row["role"] == "ADMIN":
            other_admins = db.execute(
                "SELECT COUNT(*) n FROM users WHERE role='ADMIN' AND deleted_at IS NULL AND id!=?", (user_id,)
            ).fetchone()["n"]
            if other_admins == 0:
                raise HTTPException(400, "ลบไม่ได้ เพราะเป็น Admin คนสุดท้ายในระบบ")
        if row["role"] in CREW_ROLES:
            active = db.execute(
                "SELECT tasks.task_code FROM task_assignments "
                "JOIN tasks ON tasks.id = task_assignments.task_id "
                "WHERE task_assignments.driver_id=? AND task_assignments.assignment_status='Active' "
                "AND tasks.status NOT IN ('Completed','Cancelled')",
                (user_id,),
            ).fetchall()
            if active:
                codes = ", ".join(a["task_code"] for a in active)
                raise HTTPException(
                    400, f"ลบไม่ได้ เพราะยังมีงานที่ทำอยู่ ({codes}) — กรุณาจบงานหรือมอบหมายใหม่ก่อน")
            # Hand the vehicle back so it doesn't stay tied to someone who is
            # no longer in the system.
            if row["checked_in_vehicle_id"]:
                db.execute("UPDATE vehicles SET status='Available' WHERE id=? AND status!='Breakdown'",
                           (row["checked_in_vehicle_id"],))
        db.execute("UPDATE users SET deleted_at=?, checked_in_vehicle_id=NULL, "
                   "driver_status='Not Checked-in' WHERE id=?", (now_iso(), user_id))
        db.execute("DELETE FROM sessions WHERE user_id=?", (user_id,))  # force logout everywhere
        # And stop pushing to their phone. Clearing sessions alone only ended
        # the login: the push subscription outlived it, so someone who had left
        # the company kept getting job notifications on their own device.
        db.execute("DELETE FROM push_subscriptions WHERE user_id=?", (user_id,))
        audit(db, user["id"], "delete_user", row["employee_id"])
        return {"ok": True}


@app.post("/admin/users/{user_id}/reset-password")
def admin_reset_password(user_id: int, user=Depends(require_ops())):
    """'Forgot password' to trigger on someone else's behalf: puts the account
    back into the same first-login state new accounts start in, so the user
    logs in with any password and is prompted to set a new one (reuses the
    existing /auth/set-password flow, no new mechanism needed)."""
    with get_db() as db:
        row = db.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบผู้ใช้")
        # Otherwise a dispatcher resets an admin's password and signs in as
        # them, which would make every admin-only screen reachable anyway.
        require_can_manage(user, row)
        db.execute("UPDATE users SET password_hash=? WHERE id=?", (PENDING_PASSWORD, user_id))
        db.execute("DELETE FROM sessions WHERE user_id=?", (user_id,))  # force logout everywhere
        audit(db, user["id"], "reset_password", row["employee_id"])
        return {"ok": True}


@app.get("/admin/users/{user_id}/licenses")
def get_driver_licenses(user_id: int, user=Depends(require_ops())):
    with get_db() as db:
        rows = db.execute("SELECT * FROM driver_licenses WHERE driver_id=?", (user_id,)).fetchall()
        return [dict(r) for r in rows]


class LicenseBody(BaseModel):
    vehicle_type: str
    expiry_date: Optional[str] = None  # None = permanent, never expires
    # The driver's usual vehicle of this type. "" / None clears it.
    default_vehicle_code: Optional[str] = None


# ---------------------------------------------------------------------------
# Home zones ("โซนประจำ")
# ---------------------------------------------------------------------------
# A driver works an area, not a vehicle type, so this lives on the person
# rather than on the licence the way รถประจำ does. One โซนหลัก, any number of
# โซนรอง (the zones they also help out in), or none at all — which is the
# default and means they see every zone's work equally, exactly as before.
#
# Nothing here ever BLOCKS a job: zones only change the order work is shown in
# and add a label. A zone with one person off sick must not become a zone whose
# jobs nobody can see.
MAX_HOME_ZONES = 20


def read_home_zones(db, user_id: int):
    rows = db.execute(
        "SELECT zone_key, priority FROM driver_zones WHERE driver_id=? ORDER BY priority, zone_key",
        (user_id,),
    ).fetchall()
    primary = next((r["zone_key"] for r in rows if r["priority"] == 1), None)
    return {"primary": primary,
            "secondary": [r["zone_key"] for r in rows if r["priority"] != 1]}


def write_home_zones(db, user_id: int, primary: Optional[str], secondary):
    """Replaces whatever they had. An empty primary and an empty list is a
    valid answer — 'ไม่ระบุโซนประจำ' is a choice, not a missing value."""
    primary = (primary or "").strip() or None
    wanted = []
    for key in ([primary] if primary else []) + list(secondary or []):
        key = (key or "").strip()
        # The primary wins if it was also ticked as a secondary.
        if key and key not in wanted:
            wanted.append(key)
    if len(wanted) > MAX_HOME_ZONES:
        raise HTTPException(400, f"เลือกโซนได้ไม่เกิน {MAX_HOME_ZONES} โซน")
    for key in wanted:
        # Checked against zone_key, not the Thai name: the name is editable and
        # the key is what every task stores.
        if not db.execute("SELECT 1 FROM zones WHERE zone_key=?", (key,)).fetchone():
            raise HTTPException(400, f"ไม่พบโซน {key}")
    db.execute("DELETE FROM driver_zones WHERE driver_id=?", (user_id,))
    now = now_iso()
    for key in wanted:
        db.execute(
            "INSERT INTO driver_zones (driver_id, zone_key, priority, created_at) VALUES (?,?,?,?)",
            (user_id, key, 1 if key == primary else 2, now))
    return read_home_zones(db, user_id)


class HomeZonesBody(BaseModel):
    primary_zone: Optional[str] = None
    secondary_zones: List[str] = []


@app.put("/me/zones")
def update_my_zones(body: HomeZonesBody, user=Depends(require_crew())):
    """A driver sets their own. Admins and dispatchers can set it for them
    through the staff sheet — same rules, same validation."""
    with get_db() as db:
        return write_home_zones(db, user["id"], body.primary_zone, body.secondary_zones)


@app.get("/admin/users/{user_id}/zones")
def admin_get_zones(user_id: int, user=Depends(require_ops())):
    with get_db() as db:
        return read_home_zones(db, user_id)


@app.put("/admin/users/{user_id}/zones")
def admin_set_zones(user_id: int, body: HomeZonesBody, user=Depends(require_ops())):
    _guard_target(user, user_id)
    with get_db() as db:
        target = db.execute("SELECT role, employee_id FROM users WHERE id=?", (user_id,)).fetchone()
        if not target:
            raise HTTPException(404, "ไม่พบผู้ใช้")
        if target["role"] not in CREW_ROLES:
            raise HTTPException(400, "ตั้งโซนประจำได้เฉพาะคนขับและผู้กระจายงานที่ขับรถ")
        out = write_home_zones(db, user_id, body.primary_zone, body.secondary_zones)
        audit(db, user["id"], "set_home_zones", f"{target['employee_id']}: {out}")
        return out


@app.post("/admin/users/{user_id}/licenses")
def set_driver_license(user_id: int, body: LicenseBody, user=Depends(require_ops())):
    _guard_target(user, user_id)
    with get_db() as db:
        target = db.execute("SELECT role, driver_status FROM users WHERE id=?", (user_id,)).fetchone()
    if not target:
        raise HTTPException(404, "ไม่พบผู้ใช้")
    if target["role"] not in CREW_ROLES:
        raise HTTPException(
            400, "ออกใบอนุญาตได้เฉพาะคนขับและผู้กระจายงาน — "
                 "ถ้าต้องการให้คนนี้ขับรถ กรุณาเปลี่ยนสิทธิ์ก่อน")
    if body.expiry_date:
        try:
            datetime.strptime(body.expiry_date, "%Y-%m-%d")
        except ValueError:
            raise HTTPException(400, "วันหมดอายุต้องอยู่ในรูปแบบ YYYY-MM-DD (เช่น 2027-12-31)")
    with get_db() as db:
        default_id = None
        code = (body.default_vehicle_code or "").strip()
        if code:
            v = db.execute("SELECT * FROM vehicles WHERE code=? AND retired_at IS NULL", (code,)).fetchone()
            if not v:
                raise HTTPException(404, f"ไม่พบรถรหัส {code}")
            # A usual vehicle of a type the licence isn't for could never be
            # boarded, so refuse it here rather than failing silently later.
            if v["vehicle_type"] != body.vehicle_type:
                raise HTTPException(
                    400, f"รถ {code} เป็นประเภท {v['vehicle_type']} "
                         f"จึงตั้งเป็นรถประจำของใบอนุญาต {body.vehicle_type} ไม่ได้")
            default_id = v["id"]
        # A dispatcher's row has no driver_status until the day they are given
        # a licence — give them one now, or they would be crew with no state
        # and never appear as available to anybody.
        if target["driver_status"] is None:
            db.execute("UPDATE users SET driver_status='Not Checked-in' WHERE id=?", (user_id,))
        db.execute(
            "INSERT INTO driver_licenses (driver_id, vehicle_type, expiry_date, default_vehicle_id) "
            "VALUES (?,?,?,?) ON CONFLICT(driver_id, vehicle_type) DO UPDATE SET "
            "expiry_date=excluded.expiry_date, default_vehicle_id=excluded.default_vehicle_id",
            (user_id, body.vehicle_type, body.expiry_date, default_id),
        )
        audit(db, user["id"], "set_license",
              f"driver {user_id}: {body.vehicle_type} exp={body.expiry_date} default={code or '-'}")
        return {"ok": True}


@app.delete("/admin/users/{user_id}/licenses/{vehicle_type}")
def delete_driver_license(user_id: int, vehicle_type: str, user=Depends(require_ops())):
    _guard_target(user, user_id)
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
    origin_only: bool = False
    no_location: bool = False
    from_master: bool = False


@app.get("/zones")
def list_zones(user=Depends(current_user)):
    """Zones for picking a location.

    A retired zone is withheld from everyone but an admin: the pickers must
    not offer it, while the admin page has to list it to switch it back on.
    """
    with get_db() as db:
        rows = db.execute("SELECT * FROM zones ORDER BY sort_order, id").fetchall()
        # A list-backed zone carries its places with it. The picker is built
        # synchronously as the person changes zone, so fetching them per zone
        # would put a round trip in the middle of a tap.
        places = {}
        for r in db.execute(
            "SELECT * FROM zone_locations WHERE active=1 ORDER BY sort_order, code"
        ).fetchall():
            places.setdefault(r["zone_key"], []).append(
                {"code": r["code"], "label": r["label"], "pending": r["pending"]})
    out = []
    for r in rows:
        if user["role"] != "ADMIN" and not r["active"]:
            continue
        d = dict(r)
        d["locations"] = places.get(r["zone_key"], []) if r["from_master"] else []
        out.append(d)
    return out


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
            "INSERT INTO zones (zone_key, zone_name_th, code_pattern, allow_multiple, example_code, "
            "sort_order, free_text, origin_only, no_location, from_master) VALUES (?,?,?,?,?,?,?,?,?,?)",
            (body.zone_key, body.zone_name_th, body.code_pattern, int(body.allow_multiple),
             body.example_code, max_order + 1, int(body.free_text),
             int(body.origin_only), int(body.no_location), int(body.from_master)),
        )
        audit(db, user["id"], "add_zone", body.zone_key)
        return {"ok": True}


class ZoneUpdateBody(BaseModel):
    """Everything about a zone except its key, which is immutable."""
    zone_name_th: Optional[str] = None
    code_pattern: Optional[str] = None
    example_code: Optional[str] = None
    allow_multiple: Optional[bool] = None
    free_text: Optional[bool] = None
    origin_only: Optional[bool] = None
    no_location: Optional[bool] = None
    from_master: Optional[bool] = None
    active: Optional[bool] = None
    sort_order: Optional[int] = None


@app.put("/zones/{zone_key}")
def update_zone(zone_key: str, body: ZoneUpdateBody, user=Depends(require_role("ADMIN"))):
    """Change a zone in place.

    There was no way to do this at all: the only routes were add and delete,
    so fixing a name meant deleting the zone — which took its settings with it
    and left every past task pointing at a key that no longer existed. The key
    stays fixed for exactly that reason; everything else is editable.
    """
    with get_db() as db:
        row = db.execute("SELECT * FROM zones WHERE zone_key=?", (zone_key,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบโซนนี้")
        fields = {}
        if body.zone_name_th is not None:
            name = body.zone_name_th.strip()
            if not name:
                raise HTTPException(400, "ชื่อโซนต้องไม่เว้นว่าง")
            fields["zone_name_th"] = name
        if body.code_pattern is not None:
            pattern = body.code_pattern.strip() or "^.+$"
            try:
                re.compile(pattern)
            except re.error as e:
                raise HTTPException(400, f"รูปแบบรหัสใช้ไม่ได้: {e}")
            fields["code_pattern"] = pattern
        if body.example_code is not None:
            fields["example_code"] = body.example_code.strip() or None
        for flag in ("allow_multiple", "free_text", "origin_only", "no_location",
                     "from_master", "active"):
            val = getattr(body, flag)
            if val is not None:
                fields[flag] = 1 if val else 0
        if body.sort_order is not None:
            fields["sort_order"] = int(body.sort_order)
        if not fields:
            return {"ok": True, "unchanged": True}
        # A pattern that rejects its own example would make the zone
        # unusable, and the person would only find out when someone tried to
        # create a job.
        final_pattern = fields.get("code_pattern", row["code_pattern"])
        final_example = fields.get("example_code", row["example_code"])
        free_text = fields.get("free_text", row["free_text"])
        no_location = fields.get("no_location", row["no_location"])
        # A list-backed zone's pattern is not what validates a location, so
        # its example has nothing to agree with.
        from_master = fields.get("from_master", row["from_master"])
        if final_example and not free_text and not no_location and not from_master:
            if not re.compile(final_pattern, re.IGNORECASE).match(final_example):
                raise HTTPException(400, f"ตัวอย่าง '{final_example}' ไม่ผ่านรูปแบบที่ตั้งไว้ "
                                         f"— แก้ตัวอย่างหรือรูปแบบให้ตรงกันก่อน")
        db.execute(f"UPDATE zones SET {', '.join(f'{k}=?' for k in fields)} WHERE zone_key=?",
                   tuple(fields.values()) + (zone_key,))
        audit(db, user["id"], "update_zone", f"{zone_key}: {', '.join(sorted(fields))}")
        return {"ok": True}


@app.delete("/zones/{zone_key}")
def delete_zone(zone_key: str, user=Depends(require_role("ADMIN"))):
    """Retire a zone.

    A zone any task has ever used is switched off rather than removed, so the
    history of those tasks still reads properly — the task table stores the
    key as plain text, and deleting the row turned old jobs into raw keys
    nobody could interpret. Only a zone nothing refers to is deleted for real.
    """
    with get_db() as db:
        row = db.execute("SELECT * FROM zones WHERE zone_key=?", (zone_key,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบโซนนี้")
        used_by_tasks = db.execute(
            "SELECT COUNT(*) n FROM tasks WHERE from_zone=? OR to_zone=?", (zone_key, zone_key)
        ).fetchone()["n"]
        used_by_rules = db.execute(
            "SELECT 1 FROM task_rules WHERE from_zone=? OR to_zone=?", (zone_key, zone_key)
        ).fetchone()
        if used_by_tasks or used_by_rules:
            db.execute("UPDATE zones SET active=0 WHERE zone_key=?", (zone_key,))
            audit(db, user["id"], "deactivate_zone", f"{zone_key} (tasks={used_by_tasks})")
            return {"ok": True, "deactivated": True, "task_count": used_by_tasks}
        db.execute("DELETE FROM zones WHERE zone_key=?", (zone_key,))
        audit(db, user["id"], "delete_zone", zone_key)
        return {"ok": True, "deleted": True}


class CostCenterBody(BaseModel):
    code: Optional[str] = None
    name: Optional[str] = None
    active: Optional[bool] = None
    pending: Optional[bool] = None


@app.get("/cost-centers")
def list_cost_centers(user=Depends(current_user)):
    """The cost centres a job can be booked against.

    Admin also sees the retired ones, to switch one back on; everyone else
    only gets what they may still pick.
    """
    with get_db() as db:
        rows = db.execute("SELECT * FROM cost_centers ORDER BY code").fetchall()
    if user["role"] == "ADMIN":
        return [dict(r) for r in rows]
    return [dict(r) for r in rows if r["active"]]


@app.post("/cost-centers")
def add_cost_center(body: CostCenterBody, user=Depends(current_user)):
    """Add a cost centre.

    Open to everyone on purpose: the person raising a job is the one who knows
    which account it belongs to, and making them wait for an admin would stop
    the work. One an admin adds is established straight away; one anybody else
    adds is marked pending for an admin to name or retire.
    """
    code = (body.code or "").strip()
    if not valid_cost_center(code):
        raise HTTPException(400, f"Cost center ต้องเป็นตัวเลข {COST_CENTER_LEN} หลัก")
    name = (body.name or "").strip() or None
    with get_db() as db:
        row = db.execute("SELECT * FROM cost_centers WHERE code=?", (code,)).fetchone()
        if row:
            # Already there: give it back rather than refusing. The caller is
            # picking a cost centre, not administering the list.
            if not row["active"]:
                raise HTTPException(400, f"Cost center {code} ถูกปิดใช้งานอยู่ กรุณาติดต่อแอดมิน")
            if name and not row["name"]:
                db.execute("UPDATE cost_centers SET name=? WHERE code=?", (name, code))
            return {"ok": True, "code": code, "existing": True}
        is_admin = user["role"] == "ADMIN"
        db.execute("INSERT INTO cost_centers (code, name, pending, active, created_at) "
                   "VALUES (?,?,?,1,?)", (code, name, 0 if is_admin else 1, now_iso()))
        audit(db, user["id"], "add_cost_center", f"{code}" + (f" ({name})" if name else ""))
        if not is_admin:
            notify_ops(db, f"{person_name(user)} เพิ่ม cost center {code}"
                           + (f" ({name})" if name else "") + " — ตรวจได้ในหน้าจัดการระบบ")
        return {"ok": True, "code": code, "pending": not is_admin}


@app.put("/cost-centers/{code}")
def update_cost_center(code: str, body: CostCenterBody, user=Depends(require_role("ADMIN"))):
    """Name, confirm or retire a cost centre. The CODE itself never changes:
    every account and job that uses it stores the digits as plain text."""
    with get_db() as db:
        row = db.execute("SELECT * FROM cost_centers WHERE code=?", (code,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบ cost center นี้")
        fields = {}
        if body.name is not None:
            fields["name"] = body.name.strip() or None
        if body.active is not None:
            fields["active"] = 1 if body.active else 0
        if body.pending is not None:
            fields["pending"] = 1 if body.pending else 0
        elif body.name is not None:
            # Naming one is the act of checking it.
            fields["pending"] = 0
        if not fields:
            return {"ok": True, "unchanged": True}
        db.execute(f"UPDATE cost_centers SET {', '.join(f'{k}=?' for k in fields)} WHERE code=?",
                   tuple(fields.values()) + (code,))
        audit(db, user["id"], "update_cost_center", f"{code}: {', '.join(sorted(fields))}")
        return {"ok": True}


@app.delete("/cost-centers/{code}")
def delete_cost_center(code: str, user=Depends(require_role("ADMIN"))):
    """Retire a cost centre anything refers to; delete one nothing uses."""
    with get_db() as db:
        row = db.execute("SELECT * FROM cost_centers WHERE code=?", (code,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบ cost center นี้")
        used = (db.execute("SELECT COUNT(*) n FROM tasks WHERE cost_center=?", (code,)).fetchone()["n"]
                + db.execute("SELECT COUNT(*) n FROM users WHERE cost_center=?", (code,)).fetchone()["n"])
        if used:
            db.execute("UPDATE cost_centers SET active=0 WHERE code=?", (code,))
            audit(db, user["id"], "deactivate_cost_center", f"{code} (refs={used})")
            return {"ok": True, "deactivated": True, "ref_count": used}
        db.execute("DELETE FROM cost_centers WHERE code=?", (code,))
        audit(db, user["id"], "delete_cost_center", code)
        return {"ok": True, "deleted": True}


class ZoneLocationBody(BaseModel):
    code: Optional[str] = None
    label: Optional[str] = None
    active: Optional[bool] = None
    pending: Optional[bool] = None


@app.get("/zones/{zone_key}/locations")
def list_zone_locations(zone_key: str, user=Depends(current_user)):
    """The places a list-backed zone offers.

    Admin also sees the retired ones, because the management screen has to
    list them to switch one back on; a picker must never offer them.
    """
    with get_db() as db:
        rows = db.execute(
            "SELECT * FROM zone_locations WHERE zone_key=? ORDER BY sort_order, code",
            (zone_key,)).fetchall()
    if user["role"] == "ADMIN":
        return [dict(r) for r in rows]
    return [dict(r) for r in rows if r["active"]]


@app.post("/zones/{zone_key}/locations")
def add_zone_location(zone_key: str, body: ZoneLocationBody, user=Depends(require_role("ADMIN"))):
    code = (body.code or "").strip().upper()
    if not code:
        raise HTTPException(400, "กรุณาระบุรหัสจุด")
    with get_db() as db:
        if not db.execute("SELECT 1 FROM zones WHERE zone_key=?", (zone_key,)).fetchone():
            raise HTTPException(404, "ไม่พบโซนนี้")
        if db.execute("SELECT 1 FROM zone_locations WHERE zone_key=? AND UPPER(code)=?",
                      (zone_key, code)).fetchone():
            raise HTTPException(400, f"จุด {code} มีอยู่แล้วในโซนนี้")
        order = db.execute(
            "SELECT COALESCE(MAX(sort_order),0) + 1 AS m FROM zone_locations WHERE zone_key=?",
            (zone_key,)).fetchone()["m"]
        db.execute(
            "INSERT INTO zone_locations (zone_key, code, label, sort_order, active, created_at) "
            "VALUES (?,?,?,?,1,?)",
            (zone_key, code, (body.label or "").strip() or None, order, now_iso()))
        audit(db, user["id"], "add_zone_location", f"{zone_key}: {code}")
        return {"ok": True, "code": code}


@app.post("/zones/{zone_key}/locations/request")
def request_zone_location(zone_key: str, body: ZoneLocationBody, user=Depends(current_user)):
    """Add a place to a list-backed zone from the job form.

    The person standing in front of a room that is not on the list is the one
    who knows it exists. Refusing the job until an admin adds it moves the
    problem onto the floor, so the place is created and usable at once, marked
    pending, and ops are told. An admin names, merges or retires it after.
    """
    code = (body.code or "").strip().upper()
    if not code or len(code) > 40:
        raise HTTPException(400, "กรุณาระบุรหัสจุด (ไม่เกิน 40 ตัวอักษร)")
    with get_db() as db:
        zone = db.execute("SELECT * FROM zones WHERE zone_key=?", (zone_key,)).fetchone()
        if not zone:
            raise HTTPException(404, "ไม่พบโซนนี้")
        if not zone["from_master"]:
            raise HTTPException(400, "โซนนี้ไม่ได้ใช้รายการจุด")
        if not zone["active"]:
            raise HTTPException(400, f"โซน {zone['zone_name_th']} ถูกปิดใช้งานแล้ว")
        # The zone's pattern is a shape check here, not the rule: it keeps a
        # sentence out of a field that holds a code.
        if not re.compile(zone["code_pattern"], re.IGNORECASE).match(code):
            raise HTTPException(400, f"รูปแบบรหัสจุดไม่ถูกต้อง (ตัวอย่างที่ถูกต้อง: {zone['example_code']})")
        row = db.execute("SELECT * FROM zone_locations WHERE zone_key=? AND UPPER(code)=?",
                         (zone_key, code)).fetchone()
        if row:
            if not row["active"]:
                raise HTTPException(400, f"จุด {row['code']} ถูกปิดใช้งานอยู่ กรุณาติดต่อแอดมิน")
            return {"ok": True, "code": row["code"], "existing": True}
        order = db.execute(
            "SELECT COALESCE(MAX(sort_order),0) + 1 AS m FROM zone_locations WHERE zone_key=?",
            (zone_key,)).fetchone()["m"]
        db.execute(
            "INSERT INTO zone_locations (zone_key, code, label, sort_order, active, pending, created_at) "
            "VALUES (?,?,?,?,1,1,?)",
            (zone_key, code, (body.label or "").strip() or None, order, now_iso()))
        audit(db, user["id"], "request_zone_location", f"{zone_key}: {code}")
        notify_ops(db, f"{person_name(user)} เพิ่มจุด {code} ในโซน {zone['zone_name_th']} "
                       f"— ตรวจได้ในหน้าจัดการระบบ")
        return {"ok": True, "code": code, "pending": True}


@app.put("/zones/{zone_key}/locations/{loc_id}")
def update_zone_location(zone_key: str, loc_id: int, body: ZoneLocationBody,
                         user=Depends(require_role("ADMIN"))):
    """Rename or retire one place.

    The CODE itself is not editable: every job that ever used this place
    stores that text, and changing it here would leave those jobs pointing at
    a place that no longer exists by that name. Add the new code and retire
    the old one instead.
    """
    with get_db() as db:
        row = db.execute("SELECT * FROM zone_locations WHERE id=? AND zone_key=?",
                         (loc_id, zone_key)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบจุดนี้")
        fields = {}
        if body.label is not None:
            fields["label"] = body.label.strip() or None
        if body.active is not None:
            fields["active"] = 1 if body.active else 0
        if body.pending is not None:
            fields["pending"] = 1 if body.pending else 0
        elif body.label is not None:
            # Naming a place is the act of checking it.
            fields["pending"] = 0
        if not fields:
            return {"ok": True, "unchanged": True}
        db.execute(f"UPDATE zone_locations SET {', '.join(f'{k}=?' for k in fields)} WHERE id=?",
                   tuple(fields.values()) + (loc_id,))
        audit(db, user["id"], "update_zone_location",
              f"{zone_key}: {row['code']} ({', '.join(sorted(fields))})")
        return {"ok": True}


@app.delete("/zones/{zone_key}/locations/{loc_id}")
def delete_zone_location(zone_key: str, loc_id: int, user=Depends(require_role("ADMIN"))):
    """Retire a place that jobs have used; delete one nothing refers to.

    Same rule as retiring a zone, and for the same reason: a job stores the
    location as plain text, so removing the row cannot corrupt history — but
    keeping it means the management screen can still explain what an old job
    was talking about, and switch the place back on.
    """
    with get_db() as db:
        row = db.execute("SELECT * FROM zone_locations WHERE id=? AND zone_key=?",
                         (loc_id, zone_key)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบจุดนี้")
        used = db.execute(
            "SELECT COUNT(*) n FROM tasks WHERE (from_zone=? AND from_location=?) "
            "OR (to_zone=? AND to_location=?)",
            (zone_key, row["code"], zone_key, row["code"])).fetchone()["n"]
        if used:
            db.execute("UPDATE zone_locations SET active=0 WHERE id=?", (loc_id,))
            audit(db, user["id"], "deactivate_zone_location", f"{zone_key}: {row['code']} (tasks={used})")
            return {"ok": True, "deactivated": True, "task_count": used}
        db.execute("DELETE FROM zone_locations WHERE id=?", (loc_id,))
        audit(db, user["id"], "delete_zone_location", f"{zone_key}: {row['code']}")
        return {"ok": True, "deleted": True}


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
    # Which cost centre the job is booked against. Left out, it comes from the
    # person raising it — which is why an admin or dispatcher has to state one.
    cost_center: Optional[str] = None


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


def _placeholders(n: int) -> str:
    """`?,?,?` for an IN (...) clause. Works on both backends — the Postgres
    wrapper rewrites `?` to `%s` for us."""
    return ",".join(["?"] * n)


def tasks_to_dicts(rows, db, viewer_role: Optional[str] = None) -> list:
    """Serialise many task rows with a FIXED number of queries (4), instead of
    4-6 queries per task.

    The previous per-task version meant a list of 200 tasks fired roughly a
    thousand queries — every single poll, per signed-in admin — and got worse
    every week as the task table grew. This prefetches the related users,
    vehicles and active assignments in one query each and joins them in
    memory. Output shape is byte-for-byte what task_to_dict produced.
    """
    rows = list(rows)
    if not rows:
        return []
    dicts = [dict(r) for r in rows]
    task_ids = [d["id"] for d in dicts]

    # 1. Active assignments for every task at once (was: one query per task).
    assignments_by_task = {}
    arows = db.execute(
        "SELECT task_assignments.task_id, task_assignments.vehicle_id, "
        "task_assignments.assigned_vehicle_type, task_assignments.is_active_for_driver, "
        "task_assignments.assigned_at, users.employee_id, users.full_name "
        "FROM task_assignments JOIN users ON users.id = task_assignments.driver_id "
        f"WHERE task_assignments.task_id IN ({_placeholders(len(task_ids))}) "
        "AND task_assignments.assignment_status='Active' "
        "ORDER BY task_assignments.assigned_at",
        tuple(task_ids),
    ).fetchall()
    for a in arows:
        assignments_by_task.setdefault(a["task_id"], []).append(a)

    # 1a2. Where each waiting request sits in the queue. One pass over the
    #      waiting tasks rather than a position query per card.
    #
    #      Only "ขอทันที" requests count. A booking for three o'clock is also
    #      "Waiting", and counting it would tell someone they are eighth in
    #      line when only two requests are actually competing with them now.
    queue_pos = {}
    if any(d["status"] == "Waiting" for d in dicts):
        waiting = db.execute(
            "SELECT id FROM tasks WHERE status='Waiting' AND request_type='Request Now' "
            "ORDER BY created_at, id"
        ).fetchall()
        for i, r in enumerate(waiting):
            queue_pos[r["id"]] = i + 1

    # 1b. For paused tasks, the most recent hold: who raised it, which side,
    #     and what they said. One query for all of them, never per card.
    pause_by_task = {}
    paused_ids = [d["id"] for d in dicts if d.get("status") == "Pause"]
    if paused_ids:
        prows = db.execute(
            "SELECT h.task_id, h.note, h.changed_at, u.employee_id, u.full_name, u.first_name, u.role "
            "FROM task_status_history h LEFT JOIN users u ON u.id = h.changed_by "
            f"WHERE h.task_id IN ({_placeholders(len(paused_ids))}) AND h.to_status='Pause' "
            "ORDER BY h.id",
            tuple(paused_ids),
        ).fetchall()
        for p in prows:
            note = p["note"] or ""
            m = re.match(r"\[(\w+)\]\s*(.*)", note)
            side, reason = (m.group(1), m.group(2)) if m else (None, note)
            pause_by_task[p["task_id"]] = {      # later rows overwrite: latest wins
                "side": side,                     # 'user' or 'vehicle'
                "side_label": _PAUSE_SIDE_TH.get(side, side or ""),
                "reason": reason,
                "by_role": p["role"],
                "by_name": (p["first_name"] or p["full_name"] or p["employee_id"] or ""),
                "by_employee_id": p["employee_id"],
                "at": p["changed_at"],
            }

    # 2. Every user referenced by any task (driver + requester) in one query.
    user_ids = set()
    for d in dicts:
        if d.get("current_driver_id"):
            user_ids.add(d["current_driver_id"])
        if d.get("requester_id"):
            user_ids.add(d["requester_id"])
    users_by_id = {}
    if user_ids:
        ids = list(user_ids)
        for u in db.execute(
            f"SELECT id, employee_id, full_name, contact FROM users WHERE id IN ({_placeholders(len(ids))})",
            tuple(ids),
        ).fetchall():
            users_by_id[u["id"]] = u

    # 3. Every vehicle referenced by a task OR by one of its active
    #    assignments (the old code re-queried the vehicle for each driver).
    vehicle_ids = set()
    for d in dicts:
        if d.get("current_vehicle_id"):
            vehicle_ids.add(d["current_vehicle_id"])
    for alist in assignments_by_task.values():
        for a in alist:
            if a["vehicle_id"]:
                vehicle_ids.add(a["vehicle_id"])
    vehicle_code_by_id = {}
    if vehicle_ids:
        ids = list(vehicle_ids)
        for v in db.execute(
            f"SELECT id, code FROM vehicles WHERE id IN ({_placeholders(len(ids))})",
            tuple(ids),
        ).fetchall():
            vehicle_code_by_id[v["id"]] = v["code"]

    out = []
    for d in dicts:
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
            u = users_by_id.get(d["current_driver_id"])
            if u:
                d["driver_employee_id"] = u["employee_id"]
                d["driver_name"] = u["full_name"]
        if d.get("current_vehicle_id"):
            d["vehicle_code"] = vehicle_code_by_id.get(d["current_vehicle_id"])
        req = users_by_id.get(d.get("requester_id"))
        if req:
            d["requester_employee_id"] = req["employee_id"]
            d["requester_name"] = req["full_name"]
            d["requester_contact"] = req["contact"]
        # Multi-driver: everyone currently Active on this task (primary = current_driver_id, first joined).
        drivers = [
            {
                "employee_id": a["employee_id"], "full_name": a["full_name"],
                "vehicle_code": vehicle_code_by_id.get(a["vehicle_id"]),
                # The role they were assigned as. Set even before they board,
                # so the card can say "RT 1 · PE 2" from the moment work is
                # handed out rather than once everyone has a vehicle.
                "vehicle_type": a["assigned_vehicle_type"],
                "awaiting_vehicle": a["vehicle_id"] is None,
                # False when this job is sitting in that driver's queue behind
                # another one. The driver's own screen uses it to separate the
                # job in hand from what is waiting.
                "working_on_it": bool(a["is_active_for_driver"]),
            }
            for a in assignments_by_task.get(d["id"], [])
        ]
        d["drivers"] = drivers
        d["other_driver_count"] = max(0, len(drivers) - 1)
        counts = {}
        for a in drivers:
            if a["vehicle_type"]:
                counts[a["vehicle_type"]] = counts.get(a["vehicle_type"], 0) + 1
        d["driver_counts"] = counts
        # 1 = next to be handed out. Absent once the job is no longer waiting,
        # so the card simply stops showing a position rather than freezing on
        # the last one it saw.
        d["queue_position"] = queue_pos.get(d["id"])
        d["queue_total"] = len(queue_pos) or None
        d["pause_info"] = pause_by_task.get(d["id"])
        out.append(d)
    return out


def task_to_dict(row, db, viewer_role: Optional[str] = None) -> dict:
    """Single-task convenience wrapper. Delegates to the batch version so the
    two can never drift apart."""
    return tasks_to_dicts([row], db, viewer_role=viewer_role)[0]


@app.post("/tasks")
def create_task(body: TaskCreate, user=Depends(require_role("USER", "ADMIN", "DISPATCH"))):
    """Raise a job.

    Open to admins and dispatchers as well as requesters: plenty of work is
    phoned or radioed in, and the dispatcher raising it themselves beats
    telling the caller to go and use the app. Whoever raises it is the
    requester on the job — they are the one the drivers will talk to in the
    chat and the one who gets told when it is done.
    """
    if body.request_type == "Booking":
        if not BOOKING_ENABLED:
            raise HTTPException(400, "ระบบจองล่วงหน้าปิดใช้งานอยู่ — ทุกงานเป็นการขอทันที")
        if not body.scheduled_at:
            raise HTTPException(400, "การจองล่วงหน้าต้องระบุวันเวลา")
    if body.request_type not in ("Request Now", "Booking"):
        raise HTTPException(400, "ประเภทคำขอไม่ถูกต้อง")
    if body.bu not in BU_LIST:
        raise HTTPException(400, f"BU ไม่ถูกต้อง ต้องเป็นหนึ่งใน: {', '.join(BU_LIST)}")
    # The job's own cost centre, defaulting to the requester's. An admin or
    # dispatcher has none on their profile, so for them it is required —
    # otherwise the job would be charged to nobody and the report would show a
    # blank the month after, with no way to work out whose it was.
    cost_center = (body.cost_center or "").strip() or (user["cost_center"] or "").strip()
    if not cost_center:
        raise HTTPException(400, "กรุณาระบุ cost center ของงานนี้")
    if not (cost_center.isdigit() and len(cost_center) == 5):
        raise HTTPException(400, "cost center ต้องเป็นตัวเลข 5 หลัก")
    with get_db() as db:
        # A cost centre nobody has listed yet is recorded rather than refused:
        # the person raising the job knows which account it belongs to, and
        # holding the job for an admin would stop the work.
        register_cost_center(db, cost_center)
        from_location = validate_location_string(db, body.from_zone, body.from_location)
        to_location = validate_location_string(db, body.to_zone, body.to_location, is_destination=True)
        task_type = infer_task_type(db, body.from_zone, body.to_zone)
        code = generate_task_code(db)
        cur = db.execute(
            "INSERT INTO tasks (task_code, request_type, scheduled_at, bu, cost_center, from_zone, "
            "from_location, to_zone, to_location, task_type, pallet_qty, priority, remark, "
            "barcode_pallet_id, requester_id, status, created_at) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) RETURNING id",
            (code, body.request_type, body.scheduled_at, body.bu, cost_center, body.from_zone,
             from_location, body.to_zone, to_location, task_type, body.pallet_qty, body.priority,
             body.remark, body.barcode_pallet_id, user["id"], "Waiting", now_iso()),
        )
        task_id = cur.fetchone()["id"]
        db.execute(
            "INSERT INTO task_status_history (task_id, from_status, to_status, changed_by, changed_at) "
            "VALUES (?,?,?,?,?)", (task_id, None, "Waiting", user["id"], now_iso()),
        )
        notify_ops(db, f"งานใหม่ {code} สร้างโดย {person_name(user)}", task_id)
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        result = task_to_dict(row, db, viewer_role=user["role"])
        available = db.execute(
            "SELECT COUNT(*) n FROM users WHERE (users.role='DRIVER' OR (users.role='DISPATCH' AND EXISTS (SELECT 1 FROM driver_licenses dl WHERE dl.driver_id = users.id))) AND driver_status='Available' AND deleted_at IS NULL"
        ).fetchone()["n"]
        if available == 0:
            result["warning"] = "ตอนนี้ยังไม่มีคนขับว่าง งานอาจต้องใช้เวลาสักหน่อยกว่าจะมีคนขับรับงาน"
        return result


@app.put("/tasks/{task_id}")
def update_task(task_id: int, body: TaskUpdate, user=Depends(require_role("USER"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
        if row["requester_id"] != user["id"]:
            raise HTTPException(403, "งานนี้ไม่ใช่ของคุณ")
        if row["status"] != "Waiting":
            raise HTTPException(400, "แก้ไขงานได้เฉพาะตอนที่ยังไม่เริ่มงาน (สถานะรอรับงาน)")
        fields = body.dict(exclude_unset=True)
        if not fields:
            return task_to_dict(row, db, viewer_role=user["role"])
        from_zone = fields.get("from_zone", row["from_zone"])
        to_zone = fields.get("to_zone", row["to_zone"])
        if "from_location" in fields or "from_zone" in fields:
            fields["from_location"] = validate_location_string(db, from_zone, fields.get("from_location", row["from_location"]))
            fields["from_zone"] = from_zone
        if "to_location" in fields or "to_zone" in fields:
            fields["to_location"] = validate_location_string(
                db, to_zone, fields.get("to_location", row["to_location"]), is_destination=True)
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
            raise HTTPException(404, "ไม่พบงานนี้")
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
            raise HTTPException(404, "ไม่พบรูปภาพนี้")
        return Response(content=bytes(row["data"]), media_type=row["content_type"])


@app.post("/tasks/{task_id}/priority")
def change_priority(task_id: int, body: PriorityBody, user=Depends(require_ops())):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
        if row["status"] in ("Completed", "Cancelled"):
            raise HTTPException(400, f"เปลี่ยนความสำคัญของงานสถานะ {TASK_STATUS_LABEL_TH.get(row['status'], row['status'])} ไม่ได้")
        old = row["priority"]
        db.execute("UPDATE tasks SET priority=? WHERE id=?", (body.priority, task_id))
        audit(db, user["id"], "change_priority", f"task {row['task_code']}: {old} -> {body.priority}")
        notify(db, row["requester_id"], f"งาน {row['task_code']} เปลี่ยนความสำคัญเป็น {PRIORITY_LABEL_TH.get(body.priority, body.priority)}", task_id)
        return {"ok": True}


# How far back finished work stays in the list every refresh pulls down.
# Live (non-terminal) tasks are ALWAYS included regardless of age — this only
# bounds the history tail. Without it an admin's poll fetched every task ever
# created, so the payload and query cost grew without limit for the lifetime
# of the deployment. Anything older than this is still in the database and
# still in the dashboard exports; it just isn't re-sent every few seconds.
HISTORY_WINDOW_DAYS = 14


def _history_cutoff_iso() -> str:
    return (datetime.utcnow() - timedelta(days=HISTORY_WINDOW_DAYS)).isoformat(timespec="seconds") + "Z"


@app.get("/tasks")
def list_tasks(status_filter: Optional[str] = None, scope: Optional[str] = None,
               user=Depends(current_user)):
    with get_db() as db:
        if user["role"] in ("ADMIN", "DISPATCH"):
            check_waiting_too_long(db)
        cutoff = _history_cutoff_iso()
        if user["role"] == "USER" and scope == "unit" and (user["cost_center"] or "").strip():
            # Everything booked to the same cost centre as this person, their
            # own jobs included. The match is on the JOB's cost centre rather
            # than the requester's, because a job can be raised against
            # another unit's account and it is that account's work.
            #
            # The rating clause stays personal: a colleague's unrated job is
            # not this person's to score, so it does not linger in their list.
            q = ("SELECT * FROM tasks WHERE cost_center=? AND ("
                 "status NOT IN ('Completed','Cancelled') "
                 "OR created_at >= ? "
                 "OR (status='Completed' AND rating IS NULL AND requester_id=?))")
            params = [(user["cost_center"] or "").strip(), cutoff, user["id"]]
        elif user["role"] == "USER":
            # Live tasks, recent history, and — regardless of age — anything
            # still awaiting this requester's rating, so the "ให้คะแนน" button
            # can never vanish just because the task got old.
            q = ("SELECT * FROM tasks WHERE requester_id=? AND ("
                 "status NOT IN ('Completed','Cancelled') "
                 "OR created_at >= ? "
                 "OR (status='Completed' AND rating IS NULL))")
            params = [user["id"], cutoff]
        elif user["role"] == "DRIVER":
            # Drivers see: any task they've had an assignment on (recent
            # history — including as a joiner, not just primary driver), plus
            # every non-terminal task (so they can self-assign/join/take-over).
            q = ("SELECT DISTINCT tasks.* FROM tasks "
                 "LEFT JOIN task_assignments ON task_assignments.task_id = tasks.id AND task_assignments.driver_id=? "
                 "WHERE (tasks.status NOT IN ('Completed','Cancelled') "
                 "OR (task_assignments.driver_id IS NOT NULL AND tasks.created_at >= ?))")
            params = [user["id"], cutoff]
        else:
            q = ("SELECT * FROM tasks WHERE (status NOT IN ('Completed','Cancelled') "
                 "OR created_at >= ?)")
            params = [cutoff]
        if status_filter:
            q += " AND status=?"
            params.append(status_filter)
        q += " ORDER BY created_at DESC"
        rows = db.execute(q, params).fetchall()
        return tasks_to_dicts(rows, db, viewer_role=user["role"])


@app.get("/tasks/{task_id}")
def get_task(task_id: int, user=Depends(current_user)):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
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
            raise HTTPException(404, "ไม่พบงานนี้")
        if user["role"] == "USER":
            if row["requester_id"] != user["id"] or row["status"] not in ("Waiting", "Assigned"):
                raise HTTPException(403, "ยกเลิกงานนี้ไม่ได้")
        elif user["role"] not in ("ADMIN", "DISPATCH"):
            raise HTTPException(403, "ไม่มีสิทธิ์ดำเนินการนี้")
        transition(db, row, "Cancelled", user["id"], note=reason)
        db.execute("UPDATE tasks SET cancel_reason=? WHERE id=?", (reason, task_id))
        free_up_driver_and_vehicle(db, row["current_driver_id"], row["current_vehicle_id"])
        db.execute(
            "UPDATE task_assignments SET unassigned_at=?, assignment_status='Cancelled' "
            "WHERE task_id=? AND assignment_status='Active'", (now_iso(), task_id),
        )
        audit(db, user["id"], "cancel_task", f"task {row['task_code']}: {reason or ''}")
        notify(db, row["requester_id"], f"งาน {row['task_code']} ถูกยกเลิก", task_id)
        if row["current_driver_id"]:
            notify(db, row["current_driver_id"], f"งาน {row['task_code']} ถูกยกเลิก", task_id)
        return {"ok": True}


class SendBackBody(BaseModel):
    reason: str


@app.post("/tasks/{task_id}/send-back")
def send_back_task(task_id: int, body: SendBackBody, user=Depends(require_ops())):
    """Admin kicks a Waiting task back to the User to fix (e.g. not enough PE
    vehicles right now) instead of leaving it stuck or cancelling outright."""
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
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
            raise HTTPException(404, "ไม่พบงานนี้")
        if row["requester_id"] != user["id"]:
            raise HTTPException(403, "งานนี้ไม่ใช่ของคุณ")
        if row["status"] != "Waiting" or not row["blocked_reason"]:
            raise HTTPException(400, "งานนี้ไม่ได้ถูกตีกลับ")
        db.execute(
            "UPDATE tasks SET blocked_reason=NULL, blocked_at=NULL, waiting_alert_sent=0 WHERE id=?", (task_id,)
        )
        audit(db, user["id"], "resubmit_task", f"task {row['task_code']}")
        notify_ops(db, f"งาน {row['task_code']} ถูกส่งกลับมาใหม่โดย {person_name(user)}", task_id)
        return {"ok": True}


class AssignBody(BaseModel):
    # Single driver (the original field) OR several at once. The admin picks a
    # whole crew on one screen now, so the first id becomes the primary driver
    # and the rest are added to the same task as additional active assignments.
    driver_employee_id: Optional[str] = None
    driver_employee_ids: Optional[List[str]] = None
    reason: Optional[str] = None
    estimated_minutes: Optional[int] = None
    force: bool = False

    # Richer form: which role each driver takes. Needed because a driver can
    # be assigned before boarding, when their type can't be read off a vehicle.
    # [{"employee_id": "D01", "vehicle_type": "RT"}, ...]
    crew: Optional[List[dict]] = None

    def drivers(self) -> List[str]:
        """Normalised, de-duplicated, order-preserving list of employee ids."""
        raw = [str(c.get("employee_id") or "") for c in (self.crew or [])]
        raw += list(self.driver_employee_ids or [])
        if self.driver_employee_id:
            raw.insert(0, self.driver_employee_id)
        seen, out = set(), []
        for e in raw:
            e = (e or "").strip()
            if e and e not in seen:
                seen.add(e)
                out.append(e)
        return out

    def type_for(self, employee_id: str) -> Optional[str]:
        """The role this driver was picked for, when the caller said."""
        for c in (self.crew or []):
            if str(c.get("employee_id") or "").strip() == employee_id:
                t = (c.get("vehicle_type") or "").strip().upper()
                return t or None
        return None


_STATUS_TH = {
    "Available": "ว่าง", "Busy": "ไม่ว่าง", "Pause": "พัก", "Breakdown": "ขัดข้อง", "Not Checked-in": "ยังไม่ขึ้นรถ",
    # Spoken for by a job, but still has to board a vehicle before working.
    "Assigned": "รอเลือกรถ",
}


def eligibility_issues(db, driver, vehicle, task_row, vehicle_type=None) -> List[str]:
    """Problems with putting this driver on this task.

    `vehicle` may be None: a driver can be assigned before boarding one. The
    checks that need an actual vehicle (its status) are then skipped — they
    are enforced again at check-in — but the licence and zone checks still
    run against the TYPE the driver is being assigned as, so a job can never
    be given to someone who could not do it whatever vehicle they pick.
    """
    issues = []
    vtype = vehicle["vehicle_type"] if vehicle else vehicle_type
    if not vtype:
        return ["ไม่ทราบประเภทรถที่จะใช้"]
    # Busy and Assigned are no longer problems: handing work to a driver who
    # already has some is exactly what a queue is for, and refusing it was
    # what forced the dispatcher to sit in the app dealing jobs out one at a
    # time. Only a reported breakdown or a paused driver is a real obstacle.
    if driver["driver_status"] in ("Breakdown", "Pause"):
        issues.append(f"คนขับมีสถานะ '{_STATUS_TH.get(driver['driver_status'], driver['driver_status'])}'")
    if vehicle and vehicle["status"] == "Breakdown":
        issues.append("รถคันนี้ถูกแจ้งว่าเสีย")

    # The vehicle-type/zone matrix used to be consulted here. It is gone: the
    # type is how the dispatcher organises work, not a restriction on the job,
    # and in practice the matrix mostly produced refusals nobody wanted — a
    # type added later could work nowhere until someone filled in a grid.
    return issues


def resolve_crew(db, body, task_row):
    """Turn the requested crew into (driver, vehicle_or_None, type, issues).

    A driver who has not boarded yet is allowed: the caller says which role
    they take, or it is inferred from their licences when that is
    unambiguous. Nothing is written until every member has been resolved, so
    a bad name in the list can't leave half a crew assigned.
    """
    out, problems = [], []
    for emp in body.drivers():
        drv = db.execute("SELECT * FROM users WHERE employee_id=? AND (users.role='DRIVER' OR (users.role='DISPATCH' AND EXISTS (SELECT 1 FROM driver_licenses dl WHERE dl.driver_id = users.id)))", (emp,)).fetchone()
        if not drv:
            raise HTTPException(404, f"ไม่พบคนขับ {emp}")
        veh = None
        if drv["checked_in_vehicle_id"]:
            veh = db.execute("SELECT * FROM vehicles WHERE id=?", (drv["checked_in_vehicle_id"],)).fetchone()
        vtype = veh["vehicle_type"] if veh else body.type_for(emp)
        if not vtype:
            # No vehicle and no stated role: only safe to guess when the
            # driver holds exactly one licence.
            lic = sorted(valid_license_types(db, drv["id"]))
            if len(lic) == 1:
                vtype = lic[0]
            elif not lic:
                raise HTTPException(
                    400, f"คนขับ {drv['full_name'] or emp} ไม่มีใบอนุญาตขับรถ จึงมอบหมายงานไม่ได้")
            else:
                raise HTTPException(
                    400, f"คนขับ {drv['full_name'] or emp} มีใบอนุญาตหลายประเภท ({', '.join(lic)}) "
                         f"กรุณาระบุว่าจะให้ไปในฐานะรถประเภทใด")
        stated = body.type_for(emp)
        extra = []
        if veh and stated and stated != veh["vehicle_type"]:
            # They were put down as one type while sitting in another. The job
            # does not demand a type, so this is allowed; it goes through as
            # the stated type with the vehicle left behind, and the dispatcher
            # is told so they can see what they asked for.
            extra.append(f"ขณะนี้อยู่บนรถ {veh['vehicle_type']} {veh['code']} "
                         f"— มอบหมายในฐานะ {stated} จะต้องเปลี่ยนไปขึ้นรถ {stated} ก่อนเริ่มงาน")
            vtype = stated
            veh = None
        # Checked against the role they will actually take, which is only
        # settled above. A hard stop, deliberately not one of the forceable
        # issues: check-in refuses an unlicensed vehicle at 403, so forcing
        # this through would only create a job the driver could never start.
        licensed = valid_license_types(db, drv["id"])
        if vtype not in licensed:
            raise HTTPException(
                409, f"คนขับ {drv['full_name'] or emp} ไม่มีใบอนุญาตขับรถประเภท {vtype} "
                     f"(ใบอนุญาตที่มี: {', '.join(sorted(licensed)) or 'ไม่มี'}) "
                     f"จึงมอบหมายในฐานะนี้ไม่ได้")
        iss = extra + eligibility_issues(db, drv, veh, task_row, vehicle_type=vtype)
        out.append((drv, veh, vtype, iss))
        if iss:
            problems.append(f"{drv['full_name'] or emp}: " + "; ".join(iss))
    return out, problems


def _attach_extra_driver(db, task_row, driver, vehicle, actor_id, reason, vehicle_type=None):
    """Add one more Active driver to a task that already has a primary.
    Shared by multi-driver assign and the admin 'add driver' action so the two
    can't drift apart."""
    if db.execute(
        "SELECT 1 FROM task_assignments WHERE task_id=? AND driver_id=? AND assignment_status='Active'",
        (task_row["id"], driver["id"]),
    ).fetchone():
        return False
    vtype = vehicle["vehicle_type"] if vehicle else vehicle_type
    takes_it_up = not driver_active_assignment(db, driver["id"])
    try:
        db.execute(
            "INSERT INTO task_assignments (task_id, driver_id, vehicle_id, assigned_vehicle_type, assigned_by, "
            "assigned_at, assignment_status, reason, is_active_for_driver) VALUES (?,?,?,?,?,?,?,?,?)",
            (task_row["id"], driver["id"], vehicle["id"] if vehicle else None, vtype, actor_id, now_iso(),
             "Active", reason or "มอบหมายพร้อมกันหลายคน", 1 if takes_it_up else 0),
        )
    except IntegrityError:
        raise HTTPException(409, f"คนขับ {driver['full_name'] or driver['employee_id']} เพิ่งถูกเพิ่มเข้างานนี้ไปแล้ว")
    if vehicle and takes_it_up:
        db.execute("UPDATE users SET driver_status='Busy' WHERE id=?", (driver["id"],))
        db.execute("UPDATE vehicles SET status='Busy' WHERE id=?", (vehicle["id"],))
    else:
        recompute_driver_status(db, driver["id"])
    notify(db, driver["id"], f"คุณถูกเพิ่มเข้างาน {task_row['task_code']}"
                             + ("" if takes_it_up else " (อยู่ในคิวงานของคุณ)"), task_row["id"])
    post_system_message(db, task_row["id"], driver["id"],
                        f"{person_label(db, driver['id'])} ได้รับงานนี้เรียบร้อย "
                        + (f"({vehicle['vehicle_type']} {vehicle['code']})" if vehicle
                           else f"({vtype} — รอเลือกรถ)")
                        + ("" if takes_it_up else " — อยู่ในคิวงานของคนขับ"))
    return True


def _do_assign(db, task_row, driver, vehicle, assigned_by_id, reason, estimated_minutes, is_reassign,
               vehicle_type=None):
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
            notify(db, d["driver_id"], f"คุณถูกถอดออกจากงาน {task_row['task_code']}", task_row["id"])
            free_up_driver_and_vehicle(db, d["driver_id"], d["vehicle_id"])
        if task_row["status"] == "Assigned":
            db.execute("UPDATE tasks SET accepted_at=NULL WHERE id=?", (task_row["id"],))

    vtype = vehicle["vehicle_type"] if vehicle else vehicle_type
    # If this driver already has a job in hand, the new one joins their queue
    # rather than displacing it — including on a reassign. Handing a live job
    # to someone who is mid-job elsewhere cannot make them start it this
    # second, and marking it as though they had would have given that driver
    # two jobs in hand at once, which the rest of the queue logic assumes can
    # never happen. The caller puts the task back to "Assigned" when this
    # leaves nobody actually on it.
    takes_it_up = not driver_active_assignment(db, driver["id"])
    try:
        db.execute(
            "INSERT INTO task_assignments (task_id, driver_id, vehicle_id, assigned_vehicle_type, assigned_by, "
            "assigned_at, assignment_status, reason, is_active_for_driver) VALUES (?,?,?,?,?,?,?,?,?)",
            (task_row["id"], driver["id"], vehicle["id"] if vehicle else None, vtype,
             assigned_by_id, now_iso(), "Active", reason, 1 if takes_it_up else 0),
        )
    except IntegrityError:
        raise HTTPException(409, "คนขับคนนี้มีการมอบหมายที่ Active อยู่กับงานนี้อยู่แล้ว (อาจเกิดจากกดซ้ำ หรือมีคนอื่นทำพร้อมกัน)")
    eta = add_minutes_iso(estimated_minutes) if estimated_minutes else task_row["estimated_arrival"]
    db.execute(
        "UPDATE tasks SET current_driver_id=?, current_vehicle_id=?, estimated_arrival=? WHERE id=?",
        (driver["id"], vehicle["id"] if vehicle else None, eta, task_row["id"]),
    )
    # A driver with no vehicle yet is not "Busy" in the fleet sense — they
    # still have to board one — but they are spoken for, so they must not be
    # offered to another job either. 'Assigned' says exactly that, and it also
    # covers a driver holding nothing but queued work.
    if vehicle and takes_it_up:
        db.execute("UPDATE users SET driver_status='Busy' WHERE id=?", (driver["id"],))
        db.execute("UPDATE vehicles SET status='Busy' WHERE id=?", (vehicle["id"],))
    else:
        recompute_driver_status(db, driver["id"])
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
        # Work begins the moment it is handed out — there is no separate
        # "accept" tap any more. accepted_at is still stamped because reports,
        # the driver stats endpoint and the CSV export all read it as "when did
        # this driver take the job"; leaving it NULL would silently break them.
        #
        # Unless it went into the driver's queue: then the job is theirs but
        # nobody is on it yet, which is what 'Assigned' means. It becomes
        # In Progress when they reach it (promote_next_in_queue, or they pick
        # it themselves).
        stamp = now_iso()
        if takes_it_up:
            db.execute(
                "UPDATE tasks SET status='In Progress', accepted_at=?, started_at=? WHERE id=?",
                (stamp, stamp, task_row["id"]),
            )
            db.execute(
                "INSERT INTO task_status_history (task_id, from_status, to_status, changed_by, changed_at) "
                "VALUES (?,?,?,?,?)",
                (task_row["id"], "Assigned", "In Progress", assigned_by_id, stamp),
            )
        else:
            db.execute("UPDATE tasks SET accepted_at=? WHERE id=?", (stamp, task_row["id"]))
        post_system_message(db, task_row["id"], driver["id"],
                            f"{person_label(db, driver['id'])} ได้รับงานนี้เรียบร้อย "
                            + (f"({vehicle['vehicle_type']} {vehicle['code']})" if vehicle
                               else f"({vtype} — รอเลือกรถ)")
                            + ("" if takes_it_up else " — อยู่ในคิวงานของคนขับ"))

    queue_pos = 0 if takes_it_up else len(driver_queue(db, driver["id"]))
    notify(db, driver["id"],
           f"คุณได้รับมอบหมายงาน {task_row['task_code']}"
           + (f" (อยู่ในคิวลำดับที่ {queue_pos})" if queue_pos else ""), task_row["id"])
    notify(db, task_row["requester_id"], f"งาน {task_row['task_code']} มอบหมายให้คนขับแล้ว", task_row["id"])
    if is_reassign:
        notify_ops(db, f"งาน {task_row['task_code']} มอบหมายใหม่ให้ {person_name(driver)}", task_row["id"])


@app.post("/tasks/{task_id}/assign")
def assign_task(task_id: int, body: AssignBody, user=Depends(require_ops())):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
        if row["status"] not in ("Waiting",):
            raise HTTPException(400, "งานนี้ไม่ได้อยู่ในสถานะรอรับงานแล้ว กรุณาใช้การมอบหมายใหม่แทน")
        emp_ids = body.drivers()
        if not emp_ids:
            raise HTTPException(400, "กรุณาเลือกคนขับอย่างน้อย 1 คน")

        resolved, all_issues = resolve_crew(db, body, row)
        if all_issues and not body.force:
            raise HTTPException(409, "เงื่อนไขไม่ผ่าน — " + " | ".join(all_issues) +
                                 " (ยืนยันด้วย force=true เพื่อมอบหมายทั้งที่ไม่ตรงเงื่อนไข)")
        if all_issues and body.force:
            audit(db, user["id"], "manual_override_assign", f"task {row['task_code']}: {' | '.join(all_issues)}")

        primary_drv, primary_veh, primary_type, _ = resolved[0]
        _do_assign(db, row, primary_drv, primary_veh, user["id"], body.reason,
                   body.estimated_minutes, is_reassign=False, vehicle_type=primary_type)
        # The rest join the same task. _do_assign already moved it to
        # In Progress, so these are additional crew on live work.
        for drv, veh, vtype, _ in resolved[1:]:
            _attach_extra_driver(db, row, drv, veh, user["id"], body.reason, vehicle_type=vtype)
        return {"ok": True, "assigned": [d["employee_id"] for d, _, _, _ in resolved],
                "overridden_issues": all_issues or None}


@app.post("/tasks/{task_id}/reassign")
def reassign_task(task_id: int, body: AssignBody, user=Depends(require_ops())):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
        if row["status"] not in ("Assigned", "In Progress", "Pause"):
            raise HTTPException(400, f"มอบหมายงานใหม่ไม่ได้ เพราะงานอยู่ในสถานะ {TASK_STATUS_LABEL_TH.get(row['status'], row['status'])}")
        if not body.drivers():
            raise HTTPException(400, "กรุณาเลือกคนขับอย่างน้อย 1 คน")
        resolved, issues = resolve_crew(db, body, row)
        if issues and not body.force:
            raise HTTPException(409, "เงื่อนไขไม่ผ่าน: " + " | ".join(issues) +
                                 " (ยืนยันด้วย force=true เพื่อมอบหมายทั้งที่ไม่ตรงเงื่อนไข)")
        if issues and body.force:
            audit(db, user["id"], "manual_override_reassign", f"task {row['task_code']}: {' | '.join(issues)}")
        drv0, veh0, type0, _ = resolved[0]
        _do_assign(db, row, drv0, veh0, user["id"], body.reason, body.estimated_minutes,
                   is_reassign=True, vehicle_type=type0)
        for drv, veh, vtype, _ in resolved[1:]:
            _attach_extra_driver(db, row, drv, veh, user["id"], body.reason, vehicle_type=vtype)
        # If everyone it was handed to is mid-job elsewhere, it is theirs but
        # nobody is on it yet. Say so rather than leaving it reading "กำลัง
        # ดำเนินการ" with no one working it.
        queued_only = not task_has_active_worker(db, task_id)
        if queued_only and row["status"] == "In Progress":
            db.execute("UPDATE tasks SET status='Assigned' WHERE id=?", (task_id,))
            db.execute(
                "INSERT INTO task_status_history (task_id, from_status, to_status, changed_by, changed_at) "
                "VALUES (?,?,?,?,?)", (task_id, "In Progress", "Assigned", user["id"], now_iso()),
            )
            post_system_message(db, task_id, user["id"],
                                "มอบหมายใหม่แล้ว — อยู่ในคิวของคนขับ จะเริ่มเมื่อเขาทำงานปัจจุบันเสร็จ")
        return {"ok": True, "overridden_issues": issues if issues else None,
                "queued": queued_only}


@app.post("/tasks/{task_id}/make-active")
def make_task_active(task_id: int, user=Depends(require_crew())):
    """Work this queued job now instead of the one the system picked.

    The queue is served oldest-first by default, but a driver who is already
    beside a different bay should not have to drive back for the sake of the
    order. The job they were on goes back to their queue, untouched — it only
    ever got there automatically, so nothing is lost.
    """
    with get_db() as db:
        mine = db.execute(
            "SELECT * FROM task_assignments WHERE task_id=? AND driver_id=? AND assignment_status='Active'",
            (task_id, user["id"]),
        ).fetchone()
        if not mine:
            raise HTTPException(403, "งานนี้ไม่ใช่ของคุณ")
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if row["status"] not in ("Assigned", "In Progress"):
            raise HTTPException(400, "งานนี้ไม่ได้อยู่ในสถานะที่เริ่มทำได้")
        if mine["is_active_for_driver"]:
            return {"ok": True, "already_active": True}
        current = driver_active_assignment(db, user["id"])
        if current:
            db.execute("UPDATE task_assignments SET is_active_for_driver=0 WHERE id=?", (current["id"],))
            prev = db.execute("SELECT * FROM tasks WHERE id=?", (current["task_id"],)).fetchone()
            # A crewed job keeps running on whoever is left; one nobody is on
            # goes back to "handed out but not started".
            if not task_has_active_worker(db, prev["id"]) and prev["status"] == "In Progress":
                db.execute("UPDATE tasks SET status='Assigned' WHERE id=?", (prev["id"],))
                db.execute(
                    "INSERT INTO task_status_history (task_id, from_status, to_status, changed_by, changed_at) "
                    "VALUES (?,?,?,?,?)", (prev["id"], "In Progress", "Assigned", user["id"], now_iso()),
                )
            post_system_message(db, prev["id"], user["id"],
                                f"{person_label(db, user['id'])} กลับไปทำงาน {row['task_code']} ก่อน "
                                f"— งานนี้รออยู่ในคิว")
        start_assignment(db, mine, user["id"])
        notify(db, row["requester_id"], f"งาน {row['task_code']} เริ่มดำเนินการแล้ว", task_id)
        return {"ok": True}


class SelfAssignBody(BaseModel):
    battery_level: Optional[int] = None


@app.post("/tasks/{task_id}/self-assign")
def self_assign_task(task_id: int, body: SelfAssignBody, user=Depends(require_crew())):
    """A driver takes a Waiting task themselves, using the vehicle they already checked in with."""
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
        if row["status"] != "Waiting":
            raise HTTPException(400, "งานนี้ถูกรับไปแล้ว")
        driver = db.execute("SELECT * FROM users WHERE id=?", (user["id"],)).fetchone()
        if not driver["checked_in_vehicle_id"]:
            raise HTTPException(400, "กรุณาเช็คอินรถก่อนรับงาน")
        # Already holding work is no reason to refuse — it joins their queue.
        if driver["driver_status"] in ("Breakdown", "Pause"):
            raise HTTPException(400, f"สถานะของคุณตอนนี้คือ {DRIVER_STATUS_LABEL_TH.get(driver['driver_status'], driver['driver_status'])} จึงรับงานไม่ได้")
        vehicle = db.execute("SELECT * FROM vehicles WHERE id=?", (driver["checked_in_vehicle_id"],)).fetchone()
        issues = eligibility_issues(db, driver, vehicle, row)
        if issues:
            raise HTTPException(409, "รับงานนี้ไม่ได้: " + "; ".join(issues))
        _do_assign(db, row, driver, vehicle, user["id"], "Self-assigned by driver", None, is_reassign=False)
        db.execute("UPDATE tasks SET accepted_at=?, accept_battery_level=? WHERE id=?",
                   (now_iso(), body.battery_level, task_id))
        if body.battery_level is not None:
            db.execute("UPDATE vehicles SET battery_level=? WHERE id=?", (body.battery_level, vehicle["id"]))
        notify_ops(db, f"คนขับ {person_name(user)} รับงาน {row['task_code']} ด้วยตัวเอง", task_id)
        return {"ok": True}


@app.post("/tasks/{task_id}/take-over")
def take_over_task(task_id: int, body: SelfAssignBody, user=Depends(require_crew())):
    """A driver takes over a task another driver is already working, using their
    own checked-in vehicle. Functionally a driver-initiated reassign."""
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
        if row["status"] not in ("Assigned", "In Progress", "Pause"):
            raise HTTPException(400, "ตอนนี้ยังรับช่วงงานนี้ไม่ได้")
        if row["current_driver_id"] == user["id"]:
            raise HTTPException(400, "นี่เป็นงานของคุณอยู่แล้ว")
        driver = db.execute("SELECT * FROM users WHERE id=?", (user["id"],)).fetchone()
        if not driver["checked_in_vehicle_id"]:
            raise HTTPException(400, "กรุณาเช็คอินรถก่อนรับงาน")
        # Take-over stays strict: this seizes a job someone else is working,
        # and doing it while holding your own would leave that one stranded.
        if driver["driver_status"] != "Available":
            raise HTTPException(400, f"สถานะของคุณตอนนี้คือ {DRIVER_STATUS_LABEL_TH.get(driver['driver_status'], driver['driver_status'])} จึงรับงานไม่ได้")
        vehicle = db.execute("SELECT * FROM vehicles WHERE id=?", (driver["checked_in_vehicle_id"],)).fetchone()
        issues = eligibility_issues(db, driver, vehicle, row)
        if issues:
            raise HTTPException(409, "รับช่วงงานนี้ไม่ได้: " + "; ".join(issues))
        _do_assign(db, row, driver, vehicle, user["id"], "Taken over by another driver", None, is_reassign=True)
        db.execute("UPDATE tasks SET accepted_at=?, accept_battery_level=? WHERE id=?",
                   (now_iso(), body.battery_level, task_id))
        if body.battery_level is not None:
            db.execute("UPDATE vehicles SET battery_level=? WHERE id=?", (body.battery_level, vehicle["id"]))
        notify_ops(db, f"คนขับ {person_name(user)} รับช่วงงาน {row['task_code']}", task_id)
        return {"ok": True}


class AcceptBody(BaseModel):
    battery_level: Optional[int] = None


@app.post("/tasks/{task_id}/accept")
def accept_task(task_id: int, body: AcceptBody, user=Depends(require_crew())):
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
            raise HTTPException(409, "งานนี้ถูกกดรับไปแล้ว หรือไม่ได้มอบหมายให้คุณ")
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if row["current_vehicle_id"]:
            if body.battery_level is not None:
                db.execute("UPDATE vehicles SET battery_level=? WHERE id=?",
                           (body.battery_level, row["current_vehicle_id"]))
        notify(db, row["requester_id"], f"คนขับกดรับงาน {row['task_code']} แล้ว", task_id)
        notify_ops(db, f"คนขับ {person_name(user)} กดรับงาน {row['task_code']}", task_id)
        return {"ok": True}


@app.post("/tasks/{task_id}/start")
def start_task(task_id: int, user=Depends(require_crew())):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
        if row["current_driver_id"] != user["id"]:
            raise HTTPException(403, "งานนี้ไม่ใช่ของคุณ")
        if row["status"] != "Assigned":
            raise HTTPException(400, "งานนี้ไม่ได้อยู่ในสถานะรอเริ่มงาน")
        transition(db, row, "In Progress", user["id"])
        db.execute("UPDATE tasks SET started_at=? WHERE id=?", (now_iso(), task_id))
        notify(db, row["requester_id"], f"คนขับเริ่มงาน {row['task_code']} แล้ว", task_id)
        return {"ok": True}


@app.post("/tasks/{task_id}/complete")
def complete_task(task_id: int, user=Depends(require_crew())):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
        my_assignment = db.execute(
            "SELECT * FROM task_assignments WHERE task_id=? AND driver_id=? AND assignment_status='Active'",
            (task_id, user["id"]),
        ).fetchone()
        if not my_assignment:
            raise HTTPException(403, "งานนี้ไม่ใช่ของคุณ")
        if row["status"] not in ("In Progress",):
            raise HTTPException(400, "งานนี้จบไปแล้ว หรือยังไม่ได้เริ่มทำ")
        db.execute(
            "UPDATE task_assignments SET unassigned_at=?, assignment_status='Completed' WHERE id=?",
            (now_iso(), my_assignment["id"]),
        )
        free_up_driver_and_vehicle(db, user["id"], my_assignment["vehicle_id"])
        # Whatever else happens to this task, this driver's next queued job
        # begins now. Done before the task bookkeeping below so the driver's
        # status settles on the new job rather than flickering through free.
        promote_next_in_queue(db, user["id"], user["id"])
        remaining = db.execute(
            "SELECT COUNT(*) n FROM task_assignments WHERE task_id=? AND assignment_status='Active'", (task_id,)
        ).fetchone()["n"]
        if remaining == 0:
            transition(db, row, "Completed", user["id"])
            db.execute("UPDATE tasks SET completed_at=? WHERE id=?", (now_iso(), task_id))
            notify(db, row["requester_id"], f"งาน {row['task_code']} เสร็จสิ้นแล้ว", task_id)
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
            post_system_message(db, task_id, user["id"],
                                f"{person_label(db, user['id'])} ทำส่วนของตนเสร็จแล้ว "
                                f"(ยังมีคนขับทำอยู่อีก {remaining} คน)")
            notify(db, row["requester_id"],
                   f"คนขับ {person_name(user)} ทำงาน {row['task_code']} ส่วนของตนเสร็จแล้ว "
                   f"(ยังมีคนขับอีก {remaining} คนทำอยู่)", task_id)
        return {"ok": True}


def _close_task_for_everyone(db, row, actor, note):
    """Finish a task on behalf of everyone still working on it: close every
    active assignment and hand the drivers and vehicles back. Shared by the
    admin override and the requester's own "done" button so the two can't
    leave the fleet in different states."""
    if row["status"] not in ("In Progress", "Pause"):
        raise HTTPException(
            400, f"ยืนยันงานเสร็จไม่ได้ในสถานะ {TASK_STATUS_LABEL_TH.get(row['status'], row['status'])} "
                 f"(งานที่ยังไม่เริ่ม ใช้ 'ยกเลิก' แทน)")
    active_rows = db.execute(
        "SELECT * FROM task_assignments WHERE task_id=? AND assignment_status='Active'", (row["id"],)
    ).fetchall()
    for a in active_rows:
        db.execute("UPDATE task_assignments SET unassigned_at=?, assignment_status='Completed' WHERE id=?",
                   (now_iso(), a["id"]))
        free_up_driver_and_vehicle(db, a["driver_id"], a["vehicle_id"])
    transition(db, row, "Completed", actor["id"], note=note)
    db.execute("UPDATE tasks SET completed_at=? WHERE id=?", (now_iso(), row["id"]))
    return [a["driver_id"] for a in active_rows]


@app.post("/tasks/{task_id}/requester-complete")
def requester_complete_task(task_id: int, user=Depends(require_role("USER"))):
    """The person who asked for the work confirms it's done.

    They're the one who can see the pallets arrived, so waiting for every
    driver to tap "done" separately shouldn't hold the job open. Drivers are
    told, and their vehicles are freed straight away.
    """
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
        if row["requester_id"] != user["id"]:
            raise HTTPException(403, "งานนี้ไม่ใช่ของคุณ")
        drivers = _close_task_for_everyone(db, row, user, note="Confirmed complete by requester")
        for d in drivers:
            notify(db, d, f"ผู้ขอยืนยันงาน {row['task_code']} เสร็จแล้ว", task_id)
        notify_ops(db, f"ผู้ขอยืนยันงาน {row['task_code']} เสร็จแล้ว", task_id)
        return {"ok": True}


@app.post("/tasks/{task_id}/admin-complete")
def admin_complete_task(task_id: int, user=Depends(require_ops())):
    """Admin manually force-completes a stuck/orphaned task, bypassing the normal
    driver-must-complete-their-own-part flow."""
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
        _close_task_for_everyone(db, row, user, note="Force-completed by admin")
        audit(db, user["id"], "admin_complete_task", f"task {row['task_code']}")
        notify(db, row["requester_id"], f"งาน {row['task_code']} ถูกปิดงานโดยแอดมิน", task_id)
        return {"ok": True}


class JoinTaskBody(BaseModel):
    battery_level: Optional[int] = None


@app.post("/tasks/{task_id}/join")
def join_task(task_id: int, body: JoinTaskBody, user=Depends(require_crew())):
    """A driver adds themself to a task someone else is already working, alongside
    them (not replacing) — for jobs that genuinely need more than one person."""
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
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
        # Already holding work is no reason to refuse — it joins their queue.
        if driver["driver_status"] in ("Breakdown", "Pause"):
            raise HTTPException(400, f"คุณมีสถานะ {_STATUS_TH.get(driver['driver_status'], driver['driver_status'])} ไม่ใช่ 'ว่าง'")
        vehicle = db.execute("SELECT * FROM vehicles WHERE id=?", (driver["checked_in_vehicle_id"],)).fetchone()
        issues = eligibility_issues(db, driver, vehicle, row)
        if issues:
            raise HTTPException(409, "เข้าร่วมไม่ได้: " + "; ".join(issues))
        takes_it_up = not driver_active_assignment(db, driver["id"])
        try:
            db.execute(
                "INSERT INTO task_assignments (task_id, driver_id, vehicle_id, assigned_vehicle_type, "
                "assigned_by, assigned_at, assignment_status, reason, is_active_for_driver) "
                "VALUES (?,?,?,?,?,?,?,?,?)",
                (task_id, driver["id"], vehicle["id"], vehicle["vehicle_type"], user["id"], now_iso(),
                 "Active", "Joined by driver", 1 if takes_it_up else 0),
            )
        except IntegrityError:
            raise HTTPException(409, "คุณเพิ่งเข้าร่วมงานนี้ไปแล้ว (อาจเกิดจากกดซ้ำ)")
        recompute_driver_status(db, driver["id"])
        if body.battery_level is not None:
            db.execute("UPDATE vehicles SET battery_level=? WHERE id=?", (body.battery_level, vehicle["id"]))
        if takes_it_up:
            db.execute("UPDATE vehicles SET status='Busy' WHERE id=?", (vehicle["id"],))
        post_system_message(db, task_id, user["id"],
                            f"{person_label(db, user['id'])} เข้าร่วมงาน"
                            + ("" if takes_it_up else " (อยู่ในคิวงานของเขา)"))
        notify(db, row["requester_id"], f"คนขับ {person_name(user)} เข้าร่วมงาน {row['task_code']}", task_id)
        notify_ops(db, f"คนขับ {person_name(user)} เข้าร่วมงาน {row['task_code']}", task_id)
        return {"ok": True}


class AddDriverBody(BaseModel):
    driver_employee_id: str


@app.post("/tasks/{task_id}/add-driver")
def admin_add_driver(task_id: int, body: AddDriverBody, user=Depends(require_ops())):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
        if row["status"] not in ("Assigned", "In Progress", "Pause"):
            raise HTTPException(400, "เพิ่มคนขับได้เฉพาะงานที่กำลังดำเนินการเท่านั้น")
        driver = db.execute(
            "SELECT * FROM users WHERE employee_id=? AND (users.role='DRIVER' OR (users.role='DISPATCH' AND EXISTS (SELECT 1 FROM driver_licenses dl WHERE dl.driver_id = users.id)))", (body.driver_employee_id,)
        ).fetchone()
        if not driver:
            raise HTTPException(404, "ไม่พบคนขับคนนี้")
        if db.execute(
            "SELECT 1 FROM task_assignments WHERE task_id=? AND driver_id=? AND assignment_status='Active'",
            (task_id, driver["id"]),
        ).fetchone():
            raise HTTPException(400, "คนขับคนนี้อยู่ในงานนี้อยู่แล้ว")
        vehicle = None
        if driver["checked_in_vehicle_id"]:
            vehicle = db.execute("SELECT * FROM vehicles WHERE id=?",
                                 (driver["checked_in_vehicle_id"],)).fetchone()
        # No vehicle is fine: they are put down for a type and board one when
        # they get to the job, same as any other assignment.
        licensed = sorted(valid_license_types(db, driver["id"]))
        vtype = vehicle["vehicle_type"] if vehicle else (licensed[0] if len(licensed) == 1 else None)
        if not vtype:
            raise HTTPException(
                409, f"คนขับ {driver['full_name'] or driver['employee_id']} "
                     + ("มีใบอนุญาตหลายประเภท กรุณาใช้ 'แก้ไขการมอบหมาย' เพื่อระบุประเภทรถ"
                        if licensed else "ไม่มีใบอนุญาตขับรถ จึงเพิ่มเข้างานไม่ได้"))
        issues = eligibility_issues(db, driver, vehicle, row, vehicle_type=vtype)
        if issues:
            raise HTTPException(409, "เพิ่มไม่ได้: " + "; ".join(issues))
        _attach_extra_driver(db, row, driver, vehicle, user["id"], "Added by admin", vehicle_type=vtype)
        notify(db, row["requester_id"], f"เพิ่มคนขับ {person_name(driver)} เข้างาน {row['task_code']}", task_id)
        return {"ok": True}


@app.post("/tasks/{task_id}/leave")
def leave_task(task_id: int, user=Depends(require_crew())):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
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
        # Stepping off the job in hand starts the next one they are holding,
        # the same as finishing it would.
        promote_next_in_queue(db, user["id"], user["id"])
        if active_count <= 1:
            # Last driver, task hadn't started yet -> send it back to the pool.
            db.execute(
                "UPDATE tasks SET current_driver_id=NULL, current_vehicle_id=NULL, accepted_at=NULL WHERE id=?",
                (task_id,),
            )
            transition(db, row, "Waiting", user["id"], note="Last driver left before starting")
            notify_ops(db, f"งาน {row['task_code']} ไม่มีคนขับแล้ว กลับไปรอมอบหมายใหม่", task_id)
            return {"ok": True}
        if row["current_driver_id"] == user["id"]:
            nxt = db.execute(
                "SELECT * FROM task_assignments WHERE task_id=? AND assignment_status='Active' "
                "ORDER BY assigned_at LIMIT 1", (task_id,),
            ).fetchone()
            if nxt:
                db.execute("UPDATE tasks SET current_driver_id=?, current_vehicle_id=? WHERE id=?",
                           (nxt["driver_id"], nxt["vehicle_id"], task_id))
        post_system_message(db, task_id, user["id"], f"{person_label(db, user['id'])} ออกจากงาน")
        notify(db, row["requester_id"], f"คนขับ {person_name(user)} ออกจากงาน {row['task_code']}", task_id)
        notify_ops(db, f"คนขับ {person_name(user)} ออกจากงาน {row['task_code']}", task_id)
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
            raise HTTPException(404, "ไม่พบงานนี้")
        if row["requester_id"] != user["id"]:
            raise HTTPException(403, "งานนี้ไม่ใช่ของคุณ")
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
def pause_task(task_id: int, body: PauseBody, user=Depends(current_user)):
    """Report a problem on a live job.

    Open to the drivers on it, the requester who asked for it, and admins —
    a hold-up is as often the requester's to report (goods not ready, bay
    blocked) as the driver's, and routing it through the driver just delayed
    everyone finding out.
    """
    if body.category not in ("user", "vehicle"):
        raise HTTPException(400, "category must be 'user' or 'vehicle'")
    on_the_job = False
    if user["role"] in CREW_ROLES:
        with get_db() as db:
            on_the_job = bool(db.execute(
                "SELECT 1 FROM task_assignments WHERE task_id=? AND driver_id=? "
                "AND assignment_status='Active'", (task_id, user["id"])).fetchone())
    if not on_the_job:
        # Only someone driving this job can report a VEHICLE fault, since that
        # pulls their own vehicle out of service.
        if body.category != "user":
            raise HTTPException(400, "เฉพาะคนขับที่อยู่บนงานนี้เท่านั้นที่แจ้งปัญหารถได้")
        if user["role"] == "DRIVER":
            # Unchanged for drivers: a driver who is not on the job has no
            # business pausing it.
            raise HTTPException(403, "งานนี้ไม่ใช่ของคุณ")
        with get_db() as db:
            row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
            if not row:
                raise HTTPException(404, "ไม่พบงานนี้")
            if user["role"] == "USER" and row["requester_id"] != user["id"]:
                raise HTTPException(403, "งานนี้ไม่ใช่ของคุณ")
            # Only a job in progress can be held: the state machine has no
            # Assigned -> Pause edge, and inventing one would let a resume
            # mark a job as started that nobody has started. A job that has
            # not begun is cancelled or talked about in its chat instead.
            if row["status"] != "In Progress":
                raise HTTPException(400, "แจ้งปัญหาได้เฉพาะงานที่กำลังดำเนินการอยู่ "
                                         "(งานที่ยังไม่เริ่ม ใช้แชทในงานหรือยกเลิกแทน)")
            transition(db, row, "Pause", user["id"], note=f"[user] {body.reason}")
            for a in db.execute("SELECT driver_id FROM task_assignments WHERE task_id=? "
                                "AND assignment_status='Active'", (task_id,)).fetchall():
                notify(db, a["driver_id"], f"งาน {row['task_code']} มีเหตุขัดข้อง: {body.reason}", task_id)
            notify_ops(db, f"งาน {row['task_code']} ขัดข้องฝั่งผู้ขอ: {body.reason}", task_id)
            return {"ok": True}
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
        my_assignment = db.execute(
            "SELECT * FROM task_assignments WHERE task_id=? AND driver_id=? AND assignment_status='Active'",
            (task_id, user["id"]),
        ).fetchone()
        if not my_assignment:
            raise HTTPException(403, "งานนี้ไม่ใช่ของคุณ")
        # Same rule as above, said here rather than letting transition() throw
        # "Invalid transition: Assigned -> Pause" at a driver on the floor.
        if row["status"] != "In Progress":
            raise HTTPException(400, "แจ้งปัญหาได้เฉพาะงานที่กำลังดำเนินการอยู่ "
                                     "(งานที่ยังไม่เริ่ม ใช้แชทในงานหรือยกเลิกแทน)")
        note = f"[{body.category}] {body.reason}"

        # A problem on the requester's side (location/goods not ready) blocks
        # everyone regardless of how many vehicles are on the job — pause the
        # whole task.
        if body.category == "user":
            transition(db, row, "Pause", user["id"], note=note)
            notify(db, row["requester_id"], f"งาน {row['task_code']} หยุดชั่วคราว: {body.reason}", task_id)
            notify_ops(db, f"งาน {row['task_code']} หยุดชั่วคราว (ฝั่งผู้ขอ): {body.reason}", task_id)
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
            notify(db, row["requester_id"], f"งาน {row['task_code']} หยุดชั่วคราว: {body.reason}", task_id)
            notify_ops(db, f"งาน {row['task_code']} หยุดชั่วคราว (ฝั่งรถ): {body.reason}", task_id)
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
        # The job keeps running without them, so there is no status change —
        # but a driver dropping out is exactly what the requester needs to see
        # in the timeline, or they're left wondering where one of the crew went.
        post_system_message(
            db, task_id, user["id"],
            f"{person_label(db, user['id'])} แจ้งปัญหารถ"
            f"{' ' + vehicle_row['code'] if vehicle_row else ''}: {body.reason} "
            f"— ออกจากงาน งานดำเนินต่อโดยคนขับที่เหลือ ({others} คน)")
        if row["current_driver_id"] == user["id"]:
            nxt = db.execute(
                "SELECT * FROM task_assignments WHERE task_id=? AND assignment_status='Active' "
                "ORDER BY assigned_at LIMIT 1", (task_id,),
            ).fetchone()
            if nxt:
                db.execute("UPDATE tasks SET current_driver_id=?, current_vehicle_id=? WHERE id=?",
                           (nxt["driver_id"], nxt["vehicle_id"], task_id))
        notify(db, row["requester_id"],
               f"คนขับ {person_name(user)} แจ้งปัญหารถในงาน {row['task_code']} "
               f"({body.reason}) — งานยังดำเนินต่อโดยคนขับที่เหลือ", task_id)
        notify_ops(db, f"แจ้งปัญหารถในงาน {row['task_code']} โดย {person_name(user)}: "
                          f"{body.reason} (งานยังดำเนินการอยู่ เหลือคนขับ {others} คน)", task_id)
        return {"ok": True}


@app.post("/tasks/{task_id}/resume")
def resume_task(task_id: int, user=Depends(current_user)):
    """Clear a problem and carry on.

    Either side of the job can do this: a driver on the task, the requester who
    raised it, or an admin. A problem is often the requester's to fix (goods
    not ready, bay blocked), and making them wait for the driver to tap it was
    the bottleneck.
    """
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
        if row["status"] != "Pause":
            raise HTTPException(400, "งานนี้ไม่ได้อยู่ในสถานะขัดข้อง")
        allowed = user["role"] in ("ADMIN", "DISPATCH") or (
            user["role"] == "USER" and row["requester_id"] == user["id"]
        ) or bool(db.execute(
            "SELECT 1 FROM task_assignments WHERE task_id=? AND driver_id=? AND assignment_status='Active'",
            (task_id, user["id"]),
        ).fetchone())
        if not allowed:
            raise HTTPException(403, "งานนี้ไม่ใช่ของคุณ")
        transition(db, row, "In Progress", user["id"])
        # The drivers need to know they can carry on, whoever cleared it.
        for a in db.execute(
            "SELECT driver_id FROM task_assignments WHERE task_id=? AND assignment_status='Active'", (task_id,)
        ).fetchall():
            if a["driver_id"] != user["id"]:
                notify(db, a["driver_id"], f"งาน {row['task_code']} แก้ไขเหตุขัดข้องแล้ว ทำงานต่อได้", task_id)
        if row["requester_id"] != user["id"]:
            notify(db, row["requester_id"], f"งาน {row['task_code']} แก้ไขเหตุขัดข้องแล้ว", task_id)
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
            raise HTTPException(404, "ไม่พบงานนี้")
        if not task_chat_allowed(db, row, user):
            raise HTTPException(403, "คุณไม่ได้อยู่ในห้องแชทของงานนี้")
        # sender_name was never selected here, even though the chat renderer
        # already read it — so it was always undefined and every line fell
        # back to the employee id. The team chat had it; this one didn't.
        rows = db.execute(
            "SELECT task_messages.*, users.employee_id AS sender_employee_id, "
            "users.role AS sender_role, users.full_name AS sender_full_name, "
            "users.first_name AS sender_first_name "
            "FROM task_messages JOIN users ON users.id = task_messages.sender_id "
            "WHERE task_id=? ORDER BY created_at, task_messages.id", (task_id,)
        ).fetchall()
        out = []
        for r in rows:
            d = dict(r)
            d["sender_name"] = ((r["sender_full_name"] or "").strip()
                                or (r["sender_first_name"] or "").strip()
                                or r["sender_employee_id"])
            out.append(d)
        return out


@app.get("/tasks/{task_id}/people")
def get_task_people(task_id: int, user=Depends(current_user)):
    """Everyone in this job's chat, with the details needed to reach them.

    Loaded once when the chat opens rather than attached to every message:
    a chat of fifty lines would otherwise repeat the same phone number fifty
    times, on every poll.
    """
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
        if not task_chat_allowed(db, row, user):
            raise HTTPException(403, "คุณไม่ได้อยู่ในห้องแชทของงานนี้")
        ids = set()
        if row["requester_id"]:
            ids.add(row["requester_id"])
        for r in db.execute(
            "SELECT DISTINCT driver_id FROM task_assignments WHERE task_id=?", (task_id,)
        ).fetchall():
            ids.add(r["driver_id"])
        # Anyone who has spoken in the chat — an admin who stepped in, a
        # driver who has since left the job — so tapping their name works.
        for r in db.execute(
            "SELECT DISTINCT sender_id FROM task_messages WHERE task_id=?", (task_id,)
        ).fetchall():
            ids.add(r["sender_id"])
        ids.discard(None)
        if not ids:
            return []
        rows = db.execute(
            "SELECT id, employee_id, full_name, first_name, role, cost_center, contact, email "
            f"FROM users WHERE id IN ({_placeholders(len(ids))})", tuple(ids),
        ).fetchall()
        crew = {r["driver_id"] for r in db.execute(
            "SELECT driver_id FROM task_assignments WHERE task_id=? AND assignment_status='Active'",
            (task_id,)).fetchall()}
        out = []
        for r in rows:
            out.append({
                "employee_id": r["employee_id"],
                "full_name": (r["full_name"] or "").strip() or r["employee_id"],
                "role": r["role"],
                "cost_center": r["cost_center"],
                "contact": r["contact"],
                "email": r["email"],
                "is_requester": r["id"] == row["requester_id"],
                "is_crew": r["id"] in crew,
            })
        return out


@app.post("/tasks/{task_id}/messages")
def post_task_message(task_id: int, body: ChatMessageBody, user=Depends(current_user)):
    with get_db() as db:
        row = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบงานนี้")
        if not task_chat_allowed(db, row, user):
            raise HTTPException(403, "คุณไม่ได้อยู่ในห้องแชทของงานนี้")
        if not body.message.strip():
            raise HTTPException(400, "ข้อความต้องไม่เว้นว่าง")
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
        # Every job conversation reaches ops, spoken in or not. A dispatcher
        # answers for every job on the floor, and hearing only the threads they
        # had already joined meant a problem raised in a chat no dispatcher had
        # touched sat unseen until someone happened to open that job. The
        # service worker tags a notification per job, so a burst of messages in
        # one job replaces itself rather than stacking up.
        for r in db.execute(
            "SELECT id FROM users WHERE role IN ('ADMIN','DISPATCH') AND deleted_at IS NULL"
        ).fetchall():
            if r["id"] != user["id"]:
                others.add(r["id"])
        # Anyone else who has already spoken in THIS job's chat — a colleague
        # in the same cost centre, say, who asked about it and is owed the
        # answer.
        for r in db.execute(
            "SELECT DISTINCT sender_id FROM task_messages WHERE task_id=? AND sender_id IS NOT NULL "
            "AND kind='user'", (task_id,)
        ).fetchall():
            if r["sender_id"] != user["id"]:
                others.add(r["sender_id"])
        # Never notify an account that has been closed.
        if others:
            live = {r["id"] for r in db.execute(
                f"SELECT id FROM users WHERE deleted_at IS NULL AND id IN ({','.join(['?'] * len(others))})",
                tuple(others)).fetchall()}
            others &= live
        for uid in others:
            notify(db, uid, f"มีข้อความใหม่ในงาน {row['task_code']}", task_id, kind="chat")
        return {"ok": True}


# ---------------------------------------------------------------------------
# Team chat (Admin <-> Driver only, not visible to USER) with @mentions
# ---------------------------------------------------------------------------

class TeamChatBody(BaseModel):
    message: str


@app.get("/team-chat/participants")
def team_chat_participants(user=Depends(require_role("ADMIN", "DISPATCH", "DRIVER"))):
    with get_db() as db:
        rows = db.execute(
            "SELECT employee_id, full_name, role FROM users WHERE role IN ('ADMIN','DISPATCH','DRIVER') "
            "AND deleted_at IS NULL ORDER BY role, employee_id"
        ).fetchall()
        return [dict(r) for r in rows]


@app.get("/team-chat/messages")
def get_team_chat_messages(user=Depends(require_role("ADMIN", "DISPATCH", "DRIVER"))):
    with get_db() as db:
        rows = db.execute(
            "SELECT team_chat_messages.*, users.employee_id AS sender_employee_id, users.full_name AS sender_name, "
            "users.role AS sender_role FROM team_chat_messages JOIN users ON users.id = team_chat_messages.sender_id "
            "ORDER BY team_chat_messages.created_at LIMIT 200"
        ).fetchall()
        return [dict(r) for r in rows]


@app.post("/team-chat/messages")
def post_team_chat_message(body: TeamChatBody, user=Depends(require_role("ADMIN", "DISPATCH", "DRIVER"))):
    message = body.message.strip()
    if not message:
        raise HTTPException(400, "ข้อความต้องไม่เว้นว่าง")
    with get_db() as db:
        db.execute(
            "INSERT INTO team_chat_messages (sender_id, message, created_at) VALUES (?,?,?)",
            (user["id"], message, now_iso()),
        )
        # Parse @employee_id mentions and notify each matched, real member of
        # this chat. DISPATCH used to be missing from this list, so mentioning
        # a dispatcher silently did nothing — the one role most likely to be
        # called on by name.
        mentioned_ids = set(re.findall(r"@(\w+)", message))
        for emp_id in mentioned_ids:
            target = db.execute(
                "SELECT id FROM users WHERE employee_id=? AND role IN ('ADMIN','DISPATCH','DRIVER') "
                "AND deleted_at IS NULL", (emp_id,)
            ).fetchone()
            if target and target["id"] != user["id"]:
                notify(db, target["id"], f"{person_name(user)} กล่าวถึงคุณในแชททีม", kind="team_chat")
        return {"ok": True}


# ---------------------------------------------------------------------------
# Drivers / Vehicles / Check-in / Dispatch suggestion
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Dispatch board
# ---------------------------------------------------------------------------

class ConfigBody(BaseModel):
    values: dict


@app.get("/admin/config")
def get_admin_config(user=Depends(require_role("ADMIN"))):
    """Everything editable from จัดการระบบ, plus a duplicate report.

    The duplicate report matters because phone and email only became login
    identifiers in this build. Any pre-existing account that shares one with
    another account is now ambiguous at login, and the admin needs to see that
    list BEFORE it bites someone — the app can't guess which account is right.
    """
    with get_db() as db:
        values = {k: get_config(db, k) for k in CONFIG_DEFAULTS}
        dup_phone = db.execute(
            "SELECT contact, COUNT(*) n FROM users "
            "WHERE deleted_at IS NULL AND contact IS NOT NULL AND contact != '' "
            "GROUP BY contact HAVING COUNT(*) > 1"
        ).fetchall()
        dup_email = db.execute(
            "SELECT LOWER(email) e, COUNT(*) n FROM users "
            "WHERE deleted_at IS NULL AND email IS NOT NULL AND email != '' "
            "GROUP BY LOWER(email) HAVING COUNT(*) > 1"
        ).fetchall()

        def who(col, val):
            return [dict(r) for r in db.execute(
                f"SELECT employee_id, full_name FROM users WHERE {col}=? AND deleted_at IS NULL",
                (val,)).fetchall()]

        return {
            "values": values,
            "defaults": CONFIG_DEFAULTS,
            "rack_layout": rack_layout(db),
            "duplicates": {
                "phone": [{"value": r["contact"], "count": r["n"], "users": who("contact", r["contact"])}
                          for r in dup_phone],
                "email": [{"value": r["e"], "count": r["n"], "users": who("LOWER(email)", r["e"])}
                          for r in dup_email],
            },
        }


@app.put("/admin/config")
def put_admin_config(body: ConfigBody, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        unknown = [k for k in body.values if k not in CONFIG_DEFAULTS]
        if unknown:
            raise HTTPException(400, "ตั้งค่าที่ไม่รู้จัก: " + ", ".join(unknown))
        # Invite codes used to live here as two fixed keys; they are now a
        # column on vehicle_types, edited through /vehicle-types.
        for k, v in body.values.items():
            set_config(db, k, str(v).strip())
        # Fail loudly if the new layout can't validate its own example, rather
        # than letting every location entry break silently afterwards.
        layout = rack_layout(db)
        sample = (layout["aisle_from"] + f"{layout['bay_from']:02d}"
                  + layout["level_from"] + str(layout["slot_from"]))
        parsed, err = parse_rack_code(sample, layout)
        if err:
            raise HTTPException(400, f"ค่าที่ตั้งใช้ไม่ได้ — ตัวอย่าง {sample} ไม่ผ่าน: {err}")
        audit(db, user["id"], "update_config", ", ".join(sorted(body.values)))
        return {"ok": True, "rack_layout": layout, "sample": sample}


@app.get("/rack/validate")
def rack_validate(code: str, user=Depends(current_user)):
    """Used by the location picker so the person sees the same verdict the
    server will give, before they submit."""
    with get_db() as db:
        parsed, err = parse_rack_code(code, rack_layout(db))
        return {"ok": err is None, "error": err, "parsed": parsed}


@app.get("/admin/diagnostics")
def admin_diagnostics(user=Depends(require_role("ADMIN"))):
    """Where this deployment runs, and how far it is from its database.

    The region NAMES are only a label — what actually matters is the measured
    round-trip, which is why that is taken here rather than inferred from the
    connection string. Bangkok to a database in the same region is single-digit
    to low-tens of milliseconds; anything in the hundreds means the function
    and the database are on different continents, and no amount of query
    tuning will make up the difference.
    """
    import time as _t

    # Connect + query separately: a slow connect with a fast query points at
    # network distance, while both slow points at a loaded database.
    t0 = _t.perf_counter()
    with _new_db_connection() as probe:
        connect_ms = (_t.perf_counter() - t0) * 1000
        t1 = _t.perf_counter()
        probe.execute("SELECT 1").fetchone()
        query_ms = (_t.perf_counter() - t1) * 1000

    db_host = None
    if DATABASE_URL:
        try:
            db_host = DATABASE_URL.split("@", 1)[1].split("/", 1)[0].split("?")[0]
        except Exception:
            db_host = None

    total = connect_ms + query_ms
    verdict = ("อยู่ภูมิภาคเดียวกัน" if total < 50
               else "คนละภูมิภาคแต่ใกล้กัน" if total < 150
               else "คนละทวีป — ควรย้ายให้อยู่ภูมิภาคเดียวกัน")

    return {
        "version": APP_VERSION,
        "function_region": os.environ.get("VERCEL_REGION") or os.environ.get("AWS_REGION"),
        "database_host": db_host,          # host only — never the credentials
        "database_backend": "postgres" if IS_POSTGRES else "sqlite",
        "connect_ms": round(connect_ms, 1),
        "query_ms": round(query_ms, 1),
        "total_ms": round(total, 1),
        "verdict": verdict,
        # So an admin can confirm push will reach iPhones without needing to
        # read server logs: Apple refuses anything but a real contact address.
        "vapid_subject": VAPID_SUBJECT,
        "vapid_subject_source": VAPID_SUBJECT_SOURCE,
        "vapid_subject_ios_ok": subject_is_usable(VAPID_SUBJECT),
    }


@app.get("/dispatch/board")
def dispatch_board(user=Depends(require_ops())):
    """Everything the admin needs to hand out work, in one request.

    The assign sheet used to list every driver in table order and let the admin
    guess; eligibility_issues() only ran on submit, so a wrong guess cost a full
    round-trip and a re-pick. This computes the same rules UP FRONT for every
    waiting task against every driver, so the UI can simply not offer someone
    who cannot do the job.

    Cost is fixed — four queries — regardless of how many tasks or drivers there
    are: the zone/licence tables are pulled once and the matching itself is done
    in memory. Doing it per task would have meant a request per card.
    """
    with get_db() as db:
        # --- reference data, once ---------------------------------------
        licences = {}            # driver_id -> set(vehicle_type)
        for r in db.execute("SELECT driver_id, vehicle_type FROM driver_licenses").fetchall():
            licences.setdefault(r["driver_id"], set()).add(r["vehicle_type"])

        # --- drivers, with what they are doing right now ------------------
        driver_rows = db.execute(
            "SELECT users.id, users.employee_id, users.full_name, users.driver_status, "
            "users.checked_in_vehicle_id, "
            "vehicles.code AS vehicle_code, vehicles.vehicle_type, vehicles.status AS vehicle_status, "
            "tasks.task_code AS current_task_code, tasks.id AS current_task_id, "
            "tasks.status AS current_task_status, tasks.estimated_arrival, "
            "task_assignments.assigned_at AS busy_since "
            "FROM users "
            "LEFT JOIN vehicles ON vehicles.id = users.checked_in_vehicle_id AND vehicles.retired_at IS NULL "
            "LEFT JOIN task_assignments ON task_assignments.driver_id = users.id "
            "   AND task_assignments.assignment_status='Active' "
            "LEFT JOIN tasks ON tasks.id = task_assignments.task_id "
            "   AND tasks.status NOT IN ('Completed','Cancelled') "
            "WHERE (users.role='DRIVER' OR (users.role='DISPATCH' AND EXISTS (SELECT 1 FROM driver_licenses dl WHERE dl.driver_id = users.id))) AND users.deleted_at IS NULL "
            "ORDER BY users.full_name, users.employee_id"
        ).fetchall()

        drivers = []
        for r in driver_rows:
            drivers.append({
                "id": r["id"],
                "employee_id": r["employee_id"],
                "full_name": r["full_name"],
                "driver_status": r["driver_status"],
                "vehicle_code": r["vehicle_code"],
                "vehicle_type": r["vehicle_type"],
                "vehicle_status": r["vehicle_status"],
                "licenses": sorted(licences.get(r["id"], [])),
                # "Who is free, and when will the rest be free?" — the badge
                # alone never answered the second half.
                "current_task_id": r["current_task_id"],
                "current_task_code": r["current_task_code"],
                "current_task_status": r["current_task_status"],
                "busy_since": r["busy_since"],
                "estimated_arrival": r["estimated_arrival"],
            })

        # --- the queue: oldest first, because that is the order it should
        #     be handed out in (the ops list is newest-first for browsing) ---
        tasks = db.execute(
            "SELECT * FROM tasks WHERE status='Waiting' ORDER BY created_at ASC"
        ).fetchall()
        task_dicts = tasks_to_dicts(tasks, db, viewer_role="ADMIN")

        def issues_for(d, t):
            """Same rules as eligibility_issues(), evaluated from the prefetched
            tables. Kept deliberately parallel to that function — if the two ever
            disagree, the server-side check still wins at assign time and the
            admin would see a 409, so this can only ever be optimistic, never
            permissive in a way that lets a bad assignment through."""
            out = []
            if not d["vehicle_code"]:
                out.append("ยังไม่ได้เช็คอินรถ")
                return out          # nothing else can be judged without a vehicle
            if d["driver_status"] != "Available":
                out.append(f"คนขับ{_STATUS_TH.get(d['driver_status'], d['driver_status'])}")
            if d["vehicle_status"] != "Available":
                out.append(f"รถ{_STATUS_TH.get(d['vehicle_status'], d['vehicle_status'])}")
            vt = d["vehicle_type"]
            if vt not in (d["licenses"] or []):
                out.append(f"ไม่มีใบอนุญาตรถ {vt}")
            return out

        board = []
        for t in task_dicts:
            eligible, blocked = [], {}
            for d in drivers:
                iss = issues_for(d, t)
                if iss:
                    blocked[d["employee_id"]] = iss
                else:
                    eligible.append(d["employee_id"])
            board.append({"task": t, "eligible": eligible, "blocked": blocked})

        return {"queue": board, "drivers": drivers, "server_time": now_iso()}


@app.get("/drivers/availability-summary")
def drivers_availability_summary(user=Depends(current_user)):
    """Aggregate counts only (no names/details) — safe to expose to USER and
    DRIVER, not just ADMIN, so a requester can see at a glance whether a
    driver is likely to pick up their task soon."""
    with get_db() as db:
        rows = db.execute(
            "SELECT driver_status, COUNT(*) n FROM users "
            "WHERE (users.role='DRIVER' OR (users.role='DISPATCH' AND EXISTS (SELECT 1 FROM driver_licenses dl WHERE dl.driver_id = users.id))) AND deleted_at IS NULL GROUP BY driver_status"
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
def list_drivers(user=Depends(require_ops())):
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
            "WHERE (users.role='DRIVER' OR (users.role='DISPATCH' AND EXISTS (SELECT 1 FROM driver_licenses dl WHERE dl.driver_id = users.id))) AND users.deleted_at IS NULL GROUP BY users.id"
        ).fetchall()
        # How much work each driver is holding. The dispatcher chose not to cap
        # queue length, so the only thing stopping jobs quietly piling onto one
        # person is being able to see the pile — one query for everyone, never
        # one per driver.
        queued = {r["driver_id"]: r["n"] for r in db.execute(
            "SELECT ta.driver_id, COUNT(*) n FROM task_assignments ta JOIN tasks t ON t.id = ta.task_id "
            "WHERE ta.assignment_status='Active' AND ta.is_active_for_driver=0 "
            "AND t.status NOT IN ('Completed','Cancelled') GROUP BY ta.driver_id"
        ).fetchall()}
        # Home zones for everyone in one query, so the assignment sheet can put
        # the people whose area this job is in at the top without asking per
        # driver (this list is rebuilt on every refresh).
        zones_by_driver = {}
        for r in db.execute(
            "SELECT driver_id, zone_key, priority FROM driver_zones ORDER BY priority, zone_key"
        ).fetchall():
            z = zones_by_driver.setdefault(r["driver_id"], {"primary": None, "secondary": []})
            if r["priority"] == 1:
                z["primary"] = r["zone_key"]
            else:
                z["secondary"].append(r["zone_key"])
        out = []
        for r in rows:
            d = dict(r)
            d["queue_len"] = queued.get(r["id"], 0)
            d["home_zones"] = zones_by_driver.get(r["id"], {"primary": None, "secondary": []})
            out.append(d)
        return out


class VehicleTypeCreate(BaseModel):
    type_key: str
    type_name_th: str
    description: Optional[str] = None
    # Registering with this grants a permanent licence for the type. Empty or
    # omitted means nobody can sign themselves up for it.
    invite_code: Optional[str] = None


def generate_invite_code(type_key: str) -> str:
    """A code that is awkward to guess but still readable over a radio."""
    return f"driver_{type_key.strip().lower()}_{secrets.randbelow(90000) + 10000}"


def check_invite_code_unique(db, code: str, except_type: Optional[str] = None):
    """Two types sharing a code would make the licence it grants a coin flip."""
    row = db.execute(
        "SELECT type_key FROM vehicle_types WHERE invite_code=? AND type_key != ?",
        (code, except_type or ""),
    ).fetchone()
    if row:
        raise HTTPException(400, f"รหัสเชิญนี้ถูกใช้กับประเภท {row['type_key']} อยู่แล้ว")


@app.get("/vehicle-types")
def list_vehicle_types(user=Depends(current_user)):
    with get_db() as db:
        rows = [dict(r) for r in db.execute("SELECT * FROM vehicle_types ORDER BY id").fetchall()]
    # The code is what lets someone make themselves a driver, so only the two
    # roles that run staffing ever see it.
    if user["role"] not in ("ADMIN", "DISPATCH"):
        for r in rows:
            r.pop("invite_code", None)
    return rows


@app.post("/vehicle-types")
def create_vehicle_type(body: VehicleTypeCreate, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        key = body.type_key.strip().upper()
        if not key:
            raise HTTPException(400, "กรุณาระบุรหัสประเภทรถ")
        if db.execute("SELECT 1 FROM vehicle_types WHERE type_key=?", (key,)).fetchone():
            raise HTTPException(400, "ประเภทรถนี้มีอยู่แล้ว")
        code = (body.invite_code or "").strip() or None
        if code:
            check_invite_code_unique(db, code)
        db.execute("INSERT INTO vehicle_types (type_key, type_name_th, description, invite_code) "
                   "VALUES (?,?,?,?)", (key, body.type_name_th, body.description, code))
        audit(db, user["id"], "create_vehicle_type", f"{key} invite={'set' if code else '-'}")
        return {"ok": True, "invite_code": code}


class VehicleTypeUpdate(BaseModel):
    type_name_th: Optional[str] = None
    description: Optional[str] = None
    # "" clears the code (nobody can self-register for this type any more).
    invite_code: Optional[str] = None
    # true ignores invite_code and generates a fresh one.
    regenerate_invite_code: bool = False


@app.put("/vehicle-types/{type_key}")
def update_vehicle_type(type_key: str, body: VehicleTypeUpdate, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        row = db.execute("SELECT * FROM vehicle_types WHERE type_key=?", (type_key,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบประเภทรถนี้")
        code = row["invite_code"]
        if body.regenerate_invite_code:
            code = generate_invite_code(type_key)
            check_invite_code_unique(db, code, type_key)
        elif body.invite_code is not None:
            code = body.invite_code.strip() or None
            if code:
                check_invite_code_unique(db, code, type_key)
        db.execute(
            "UPDATE vehicle_types SET type_name_th=?, description=?, invite_code=? WHERE type_key=?",
            (body.type_name_th if body.type_name_th is not None else row["type_name_th"],
             body.description if body.description is not None else row["description"],
             code, type_key),
        )
        audit(db, user["id"], "update_vehicle_type", f"{type_key} invite={'set' if code else '-'}")
        return {"ok": True, "invite_code": code}


@app.delete("/vehicle-types/{type_key}")
def delete_vehicle_type(type_key: str, user=Depends(require_role("ADMIN"))):
    with get_db() as db:
        live = db.execute(
            "SELECT COUNT(*) n FROM vehicles WHERE vehicle_type=? AND retired_at IS NULL",
            (type_key,),
        ).fetchone()["n"]
        if live:
            raise HTTPException(400, f"ลบไม่ได้ เพราะยังมีรถที่ใช้ประเภทนี้อยู่ {live} คัน "
                                     f"(ปลดระวางรถทั้งหมดก่อน)")
        db.execute("DELETE FROM vehicle_types WHERE type_key=?", (type_key,))
        # Licences for a type that no longer exists would keep counting toward
        # what a driver may operate.
        db.execute("DELETE FROM driver_licenses WHERE vehicle_type=?", (type_key,))
        audit(db, user["id"], "delete_vehicle_type", type_key)
        return {"ok": True}


# Everything operational filters retired vehicles out; only reports and
# historical lookups ever see them.
@app.get("/vehicles")
def list_vehicles(user=Depends(current_user)):
    with get_db() as db:
        rows = db.execute("SELECT * FROM vehicles WHERE retired_at IS NULL ORDER BY code").fetchall()
        return [dict(r) for r in rows]


class VehicleCreate(BaseModel):
    code: str
    vehicle_type: str
    battery_level: int = 100


@app.post("/vehicles")
def create_vehicle(body: VehicleCreate, user=Depends(require_ops())):
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
def delete_vehicle(vehicle_id: int, user=Depends(require_ops())):
    """Take a vehicle out of service.

    It is retired, not deleted. Every finished task and assignment still
    points at the row, so removing it outright failed on a foreign key the
    moment the vehicle had done any work at all — the admin got a 500 with no
    explanation. Retiring is also what the master-fleet sync already does, so
    the two paths now agree.
    """
    with get_db() as db:
        row = db.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบรถคันนี้")
        if row["retired_at"]:
            return {"ok": True, "already_retired": True}
        active_task = db.execute(
            "SELECT 1 FROM tasks WHERE current_vehicle_id=? AND status NOT IN ('Completed','Cancelled')",
            (vehicle_id,),
        ).fetchone()
        if active_task:
            raise HTTPException(400, "ปลดระวางไม่ได้ เพราะรถคันนี้กำลังถูกใช้งานในงานที่ยังไม่เสร็จ")
        if db.execute("SELECT 1 FROM users WHERE checked_in_vehicle_id=?", (vehicle_id,)).fetchone():
            raise HTTPException(400, "ปลดระวางไม่ได้ เพราะมีคนขับเช็คอินรถคันนี้อยู่")
        db.execute("UPDATE vehicles SET retired_at=?, status='Retired' WHERE id=?", (now_iso(), vehicle_id))
        # Nobody should be auto-boarded into a vehicle that has left the fleet.
        db.execute("UPDATE driver_licenses SET default_vehicle_id=NULL WHERE default_vehicle_id=?",
                   (vehicle_id,))
        audit(db, user["id"], "retire_vehicle", row["code"])
        return {"ok": True}


class VehicleUpdateBody(BaseModel):
    vehicle_type: Optional[str] = None


@app.put("/vehicles/{vehicle_id}")
def update_vehicle(vehicle_id: int, body: VehicleUpdateBody, user=Depends(require_ops())):
    with get_db() as db:
        row = db.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบรถคันนี้")
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


def vehicle_holders(db):
    """Who is currently sitting in each vehicle: vehicle_id -> row.

    A driver keeps a vehicle until somebody takes it, so "held" is a normal
    state rather than a problem, and the picker has to say who holds what.
    """
    rows = db.execute(
        "SELECT id, employee_id, full_name, first_name, driver_status, checked_in_vehicle_id "
        "FROM users WHERE (users.role='DRIVER' OR (users.role='DISPATCH' AND EXISTS (SELECT 1 FROM driver_licenses dl WHERE dl.driver_id = users.id))) AND deleted_at IS NULL AND checked_in_vehicle_id IS NOT NULL"
    ).fetchall()
    return {r["checked_in_vehicle_id"]: r for r in rows}


# Someone mid-job cannot have their vehicle taken out from under them;
# anything else is fair game.
HOLD_PROTECTED_STATUSES = ("Busy", "Breakdown")


@app.get("/vehicles/available-for-checkin")
def vehicles_available_for_checkin(user=Depends(require_crew())):
    """Every vehicle this driver is licensed for, including ones another
    driver is currently in.

    Hiding held vehicles used to mean a vehicle left with someone who went
    home stayed invisible to everyone — the fleet quietly shrank. They are
    listed instead, saying who has them, and can be taken over unless that
    person is actually working.
    """
    with get_db() as db:
        rows = db.execute(
            "SELECT * FROM vehicles WHERE retired_at IS NULL AND status != 'Breakdown' ORDER BY code"
        ).fetchall()
        allowed = valid_license_types(db, user["id"])
        holders = vehicle_holders(db)
        defaults = {r["default_vehicle_id"] for r in db.execute(
            "SELECT default_vehicle_id FROM driver_licenses WHERE driver_id=? AND default_vehicle_id IS NOT NULL",
            (user["id"],)).fetchall()}
        out = []
        for r in rows:
            if r["vehicle_type"] not in allowed:
                continue
            h = holders.get(r["id"])
            mine = h is not None and h["id"] == user["id"]
            d = dict(r)
            d["is_default"] = r["id"] in defaults
            d["is_mine"] = mine
            if h is None or mine:
                d["held_by"] = None
                d["can_take"] = True
                d["blocked_reason"] = None
            else:
                d["held_by"] = h["full_name"] or h["employee_id"]
                busy = h["driver_status"] in HOLD_PROTECTED_STATUSES
                d["can_take"] = not busy
                d["blocked_reason"] = (f"{d['held_by']} กำลังใช้งานอยู่" if busy else None)
            out.append(d)
        return out


class DriverStatusBody(BaseModel):
    status: str


@app.post("/drivers/status")
def set_driver_status(body: DriverStatusBody, user=Depends(require_crew())):
    if body.status not in ("Not Checked-in", "Available", "Busy", "Pause", "Breakdown"):
        raise HTTPException(400, "สถานะไม่ถูกต้อง")
    with get_db() as db:
        db.execute("UPDATE users SET driver_status=? WHERE id=?", (body.status, user["id"]))
        if body.status == "Breakdown":
            notify_ops(db, f"คนขับ {person_name(user)} แจ้งสถานะรถเสีย")
        return {"ok": True}


class CheckInBody(BaseModel):
    vehicle_code: str
    battery_level: Optional[int] = None


def valid_license_types(db, driver_id) -> set:
    """Vehicle types this driver may operate TODAY. A licence with no expiry
    is permanent; one past its expiry date no longer counts."""
    today = datetime.utcnow().strftime("%Y-%m-%d")
    rows = db.execute("SELECT vehicle_type, expiry_date FROM driver_licenses WHERE driver_id=?",
                      (driver_id,)).fetchall()
    return {r["vehicle_type"] for r in rows if not r["expiry_date"] or r["expiry_date"] >= today}


@app.post("/vehicles/check-in")
def vehicle_check_in(body: CheckInBody, user=Depends(require_crew())):
    with get_db() as db:
        v = db.execute("SELECT * FROM vehicles WHERE code=?", (body.vehicle_code,)).fetchone()
        if not v:
            raise HTTPException(404, "ไม่พบรถคันนี้")
        if v["retired_at"]:
            raise HTTPException(400, "รถคันนี้ถูกปลดระวางแล้ว ไม่สามารถเช็คอินได้")
        if v["status"] == "Breakdown":
            raise HTTPException(400, "รถคันนี้ถูกแจ้งว่าเสีย จึงเช็คอินไม่ได้")
        # No licence for this type, no boarding. Checked here rather than only
        # at assignment, so a driver can never be sitting in a vehicle they
        # aren't allowed to operate.
        allowed = valid_license_types(db, user["id"])
        if v["vehicle_type"] not in allowed:
            have = ", ".join(sorted(allowed)) or "ไม่มี"
            raise HTTPException(
                403, f"คุณไม่มีใบอนุญาตขับรถประเภท {v['vehicle_type']} (ใบอนุญาตที่มี: {have}) "
                     f"กรุณาติดต่อแอดมิน")
        # Taking a vehicle somebody else is in. Allowed — a driver holds a
        # vehicle until someone needs it, so this is the normal way one
        # changes hands at the end of a shift — unless they are mid-job, in
        # which case nothing may pull it out from under them.
        holder = db.execute(
            "SELECT * FROM users WHERE (users.role='DRIVER' OR (users.role='DISPATCH' AND EXISTS (SELECT 1 FROM driver_licenses dl WHERE dl.driver_id = users.id))) AND deleted_at IS NULL "
            "AND checked_in_vehicle_id=? AND id!=?", (v["id"], user["id"]),
        ).fetchone()
        if holder:
            if holder["driver_status"] in HOLD_PROTECTED_STATUSES:
                raise HTTPException(
                    409, f"รถ {v['code']} กำลังถูกใช้งานโดย "
                         f"{holder['full_name'] or holder['employee_id']} อยู่ กรุณาเลือกคันอื่น")
            db.execute("UPDATE users SET checked_in_vehicle_id=NULL, driver_status='Not Checked-in' "
                       "WHERE id=?", (holder["id"],))
            # They will not be looking at the app, so tell them rather than
            # letting them discover it next time they try to start a job.
            notify(db, holder["id"],
                   f"รถ {v['code']} ถูกรับไปใช้ต่อโดย {person_label(db, user['id'])} "
                   f"— ถ้ายังต้องใช้รถ กรุณาเลือกรถคันใหม่")
            audit(db, user["id"], "vehicle_taken_over",
                  f"{v['code']} from {holder['employee_id']}")

        if body.battery_level is not None:
            db.execute("UPDATE vehicles SET battery_level=? WHERE id=?", (body.battery_level, v["id"]))
        db.execute("UPDATE users SET checked_in_vehicle_id=? WHERE id=?", (v["id"], user["id"]))
        cur = db.execute(
            "SELECT driver_status FROM users WHERE id=?", (user["id"],)
        ).fetchone()
        # Jobs already assigned to this driver that were waiting on a vehicle
        # now get one. This is the "choose a vehicle when you take the job"
        # step: the driver boards, and every job expecting that type binds.
        bound = db.execute(
            "SELECT ta.id, t.task_code FROM task_assignments ta JOIN tasks t ON t.id = ta.task_id "
            "WHERE ta.driver_id=? AND ta.assignment_status='Active' AND ta.vehicle_id IS NULL "
            "AND ta.assigned_vehicle_type=? AND t.status NOT IN ('Completed','Cancelled')",
            (user["id"], v["vehicle_type"]),
        ).fetchall()
        for b in bound:
            db.execute("UPDATE task_assignments SET vehicle_id=? WHERE id=?", (v["id"], b["id"]))
        if bound:
            db.execute("UPDATE tasks SET current_vehicle_id=? WHERE current_driver_id=? "
                       "AND current_vehicle_id IS NULL", (v["id"], user["id"]))
            db.execute("UPDATE vehicles SET status='Busy' WHERE id=?", (v["id"],))
            for b in bound:
                tid = db.execute("SELECT task_id FROM task_assignments WHERE id=?", (b["id"],)).fetchone()["task_id"]
                post_system_message(db, tid, user["id"],
                                    f"{person_label(db, user['id'])} เลือกรถ {v['vehicle_type']} {v['code']} แล้ว")

        # Boarding can change what their status should be — a job that was
        # waiting on a vehicle now has one. recompute rather than guess, so a
        # driver holding queued work does not come out as "Available".
        if cur["driver_status"] in ("Not Checked-in", "Assigned", "Available"):
            recompute_driver_status(db, user["id"])
        audit(db, user["id"], "vehicle_check_in",
              body.vehicle_code + (f" battery={body.battery_level}%" if body.battery_level is not None else ""))
        return {"ok": True}


@app.post("/vehicles/auto-board")
def auto_board_default_vehicle(user=Depends(require_crew())):
    """Put the driver back in their usual vehicle, called once when the app
    opens.

    This is what makes "no picking a vehicle every day" true in practice: the
    driver who always takes WH-01 opens the app and is in WH-01, with nothing
    to tap. It is deliberately a POST the client calls explicitly rather than
    a side effect of reading /me, so nothing boards a vehicle behind the
    driver's back on a background poll.

    Returns {"boarded": code} when it did something, {"boarded": null} plus a
    reason otherwise. Never an error: failing to auto-board is not a problem,
    it just means the driver picks a vehicle the usual way.
    """
    with get_db() as db:
        me = db.execute("SELECT * FROM users WHERE id=?", (user["id"],)).fetchone()
        if me["checked_in_vehicle_id"]:
            return {"boarded": None, "reason": "already_on_vehicle"}
        allowed = valid_license_types(db, user["id"])
        rows = db.execute(
            "SELECT dl.vehicle_type, v.* FROM driver_licenses dl "
            "JOIN vehicles v ON v.id = dl.default_vehicle_id "
            "WHERE dl.driver_id=? AND v.retired_at IS NULL", (user["id"],),
        ).fetchall()
        candidates = [r for r in rows if r["vehicle_type"] in allowed]
        if not candidates:
            return {"boarded": None, "reason": "no_default"}
        holders = vehicle_holders(db)
        for v in candidates:
            if v["status"] == "Breakdown":
                continue
            h = holders.get(v["id"])
            if h and h["driver_status"] in HOLD_PROTECTED_STATUSES:
                continue    # someone is working in it; leave them alone
            try:
                vehicle_check_in(CheckInBody(vehicle_code=v["code"]), user)
            except HTTPException:
                continue    # licence expired between queries, vehicle gone: just skip
            return {"boarded": v["code"], "vehicle_type": v["vehicle_type"],
                    "taken_from": (h["full_name"] or h["employee_id"]) if h else None}
        return {"boarded": None, "reason": "default_in_use"}


@app.post("/vehicles/check-out")
def vehicle_check_out(user=Depends(require_crew())):
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
def suggest_dispatch(task_id: int, user=Depends(require_ops())):
    with get_db() as db:
        task = db.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not task:
            raise HTTPException(404, "ไม่พบงานนี้")
        vehicles = db.execute("SELECT * FROM vehicles WHERE status='Available'").fetchall()
        suggestions = []
        for v in vehicles:
            drivers = db.execute(
                "SELECT DISTINCT users.* FROM users JOIN driver_licenses ON driver_licenses.driver_id = users.id "
                "WHERE (users.role='DRIVER' OR (users.role='DISPATCH' AND EXISTS (SELECT 1 FROM driver_licenses dl WHERE dl.driver_id = users.id))) AND users.driver_status='Available' AND users.deleted_at IS NULL "
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
def my_driver_stats(period: str = "day", date: Optional[str] = None, user=Depends(require_crew())):
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
        raise HTTPException(400, "รูปแบบวันที่ไม่ถูกต้อง")
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


HISTORY_DAY_BUCKETS = 14
HISTORY_MONTH_BUCKETS = 12


def _history_scope(role: str) -> str:
    """Whose work a history summary counts.

    A driver is measured on what they moved, a requester on what they asked
    for, and ops on the floor as a whole. Same screen, three honest answers —
    the alternative, one number for everyone, is true for nobody.
    """
    if role in ("ADMIN", "DISPATCH"):
        return "all"
    if role == "DRIVER":
        return "driver"
    return "requester"


@app.get("/history/summary")
def history_summary(period: str = "day", date: Optional[str] = None,
                    user=Depends(current_user)):
    """Totals for one day or month, plus the buckets around it for a chart.

    The numbers are for the bucket that was picked; the series is the context
    it sits in, so "a good day" can be read against the fortnight rather than
    against nothing. Everything is bucketed in Bangkok time — a job finished
    at 1am here belongs to that day, not to the UTC day before it.
    """
    if period not in ("day", "month"):
        raise HTTPException(400, "period ต้องเป็น day หรือ month")
    now_bkk = datetime.utcnow() + BKK_OFFSET
    try:
        if period == "day":
            end_day = (datetime.strptime(date, "%Y-%m-%d") if date
                       else now_bkk.replace(hour=0, minute=0, second=0, microsecond=0))
            keys = [(end_day - timedelta(days=i)).strftime("%Y-%m-%d")
                    for i in range(HISTORY_DAY_BUCKETS - 1, -1, -1)]
            labels = [(end_day - timedelta(days=i)).strftime("%-d/%-m")
                      for i in range(HISTORY_DAY_BUCKETS - 1, -1, -1)]
            start_bkk = end_day - timedelta(days=HISTORY_DAY_BUCKETS - 1)
            end_bkk = end_day + timedelta(days=1)
            key_of = lambda at: at.strftime("%Y-%m-%d")
        else:
            end_month = (datetime.strptime(date, "%Y-%m") if date
                         else now_bkk.replace(day=1, hour=0, minute=0, second=0, microsecond=0))
            months = []
            y, m = end_month.year, end_month.month
            for _ in range(HISTORY_MONTH_BUCKETS):
                months.append((y, m))
                m -= 1
                if m == 0:
                    y, m = y - 1, 12
            months.reverse()
            keys = [f"{y:04d}-{m:02d}" for y, m in months]
            labels = [f"{m}/{str(y)[-2:]}" for y, m in months]
            start_bkk = datetime(months[0][0], months[0][1], 1)
            end_bkk = (end_month + timedelta(days=32)).replace(day=1)
            key_of = lambda at: at.strftime("%Y-%m")
    except ValueError:
        raise HTTPException(400, "รูปแบบวันที่ไม่ถูกต้อง")

    selected = keys[-1]
    since_utc = (start_bkk - BKK_OFFSET).isoformat(timespec="seconds") + "Z"
    until_utc = (end_bkk - BKK_OFFSET).isoformat(timespec="seconds") + "Z"
    scope = _history_scope(user["role"])

    buckets = {k: {"key": k, "label": lb, "completed": 0, "cancelled": 0, "pallets": 0.0}
               for k, lb in zip(keys, labels)}
    with get_db() as db:
        if scope == "driver":
            # Credit is the job's pallets split between everyone who finished
            # their part of it, so two drivers on one job do not each report
            # having moved all of it.
            rows = db.execute(
                "SELECT tasks.id, tasks.status, tasks.pallet_qty, tasks.completed_at, "
                "(SELECT COUNT(*) FROM task_assignments x WHERE x.task_id = tasks.id "
                " AND x.assignment_status='Completed') AS crew "
                "FROM task_assignments JOIN tasks ON tasks.id = task_assignments.task_id "
                "WHERE task_assignments.driver_id=? AND task_assignments.assignment_status='Completed' "
                "AND tasks.status='Completed' AND tasks.completed_at>=? AND tasks.completed_at<?",
                (user["id"], since_utc, until_utc)).fetchall()
            for r in rows:
                at = _to_bkk(r["completed_at"])
                b = buckets.get(key_of(at)) if at else None
                if not b:
                    continue
                b["completed"] += 1
                b["pallets"] += (r["pallet_qty"] or 0) / max(1, r["crew"] or 1)
        else:
            q = ("SELECT status, pallet_qty, completed_at, created_at FROM tasks "
                 "WHERE status IN ('Completed','Cancelled')")
            params = []
            if scope == "requester":
                q += " AND requester_id=?"
                params.append(user["id"])
            rows = db.execute(q, params).fetchall()
            for r in rows:
                # A cancelled job has no completion time; it is counted on the
                # day it was raised, which is the only date it has.
                at = _to_bkk(r["completed_at"] or r["created_at"])
                b = buckets.get(key_of(at)) if at else None
                if not b:
                    continue
                if r["status"] == "Completed":
                    b["completed"] += 1
                    b["pallets"] += r["pallet_qty"] or 0
                else:
                    b["cancelled"] += 1

    series = []
    for k in keys:
        b = buckets[k]
        b["pallets"] = round(b["pallets"], 1)
        series.append(b)
    chosen = buckets[selected]
    return {
        "period": period, "date": selected, "scope": scope,
        "summary": {"completed": chosen["completed"], "cancelled": chosen["cancelled"],
                    "pallets": chosen["pallets"]},
        "window": {"completed": sum(b["completed"] for b in series),
                   "cancelled": sum(b["cancelled"] for b in series),
                   "pallets": round(sum(b["pallets"] for b in series), 1)},
        "buckets": series,
    }


# ---------------------------------------------------------------------------
# Breakdowns
# ---------------------------------------------------------------------------

class BreakdownBody(BaseModel):
    target_type: str
    vehicle_code: Optional[str] = None
    description: Optional[str] = None


@app.post("/breakdowns")
def report_breakdown(body: BreakdownBody, user=Depends(require_crew())):
    with get_db() as db:
        if body.target_type == "driver":
            target_id = user["id"]
            label = person_name(user)
            db.execute("UPDATE users SET driver_status='Breakdown' WHERE id=?", (user["id"],))
        elif body.target_type == "vehicle":
            if not body.vehicle_code:
                raise HTTPException(400, "กรุณาระบุรหัสรถ")
            v = db.execute("SELECT * FROM vehicles WHERE code=?", (body.vehicle_code,)).fetchone()
            if not v:
                raise HTTPException(404, "ไม่พบรถคันนี้")
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
        notify_ops(db, f"แจ้งเสีย: {'คนขับ' if body.target_type == 'driver' else 'รถ'} {label} โดย {person_name(user)}")
        return {"ok": True}


@app.get("/breakdowns")
def list_breakdowns(user=Depends(require_ops())):
    with get_db() as db:
        rows = db.execute("SELECT * FROM breakdowns ORDER BY created_at DESC").fetchall()
        return [dict(r) for r in rows]


class ResolveBreakdownBody(BaseModel):
    resolution_note: Optional[str] = None


@app.post("/breakdowns/{breakdown_id}/resolve")
def resolve_breakdown(breakdown_id: int, body: ResolveBreakdownBody = ResolveBreakdownBody(), user=Depends(require_ops())):
    with get_db() as db:
        b = db.execute("SELECT * FROM breakdowns WHERE id=?", (breakdown_id,)).fetchone()
        if not b:
            raise HTTPException(404, "ไม่พบรายการแจ้งเสียนี้")
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
def broadcast_notification(body: BroadcastBody, user=Depends(require_ops())):
    message = body.message.strip()
    if not message:
        raise HTTPException(400, "กรุณากรอกข้อความประกาศ")
    with get_db() as db:
        # "ALL" used to mean USER+DRIVER, so an announcement to everyone reached
        # neither the dispatchers nor the other admins — the people most likely
        # to need to act on it. OPS is the new shorthand for the office side.
        if body.role == "ALL":
            targets = db.execute("SELECT id FROM users WHERE deleted_at IS NULL").fetchall()
        elif body.role == "OPS":
            targets = db.execute(
                "SELECT id FROM users WHERE role IN ('ADMIN','DISPATCH') AND deleted_at IS NULL").fetchall()
        elif body.role in ROLES:
            targets = db.execute("SELECT id FROM users WHERE role=? AND deleted_at IS NULL", (body.role,)).fetchall()
        else:
            raise HTTPException(400, f"กลุ่มผู้รับไม่ถูกต้อง: {body.role}")
        # The sender does not need their own announcement back.
        targets = [t for t in targets if t["id"] != user["id"]]
        for t in targets:
            notify(db, t["id"], f"[แอดมิน] {message}")
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
        raise HTTPException(400, "ข้อมูลการสมัครรับแจ้งเตือนไม่ถูกต้อง")
    with get_db() as db:
        db.execute(
            "INSERT INTO push_subscriptions (user_id, endpoint, p256dh, auth, created_at) VALUES (?,?,?,?,?) "
            "ON CONFLICT(endpoint) DO UPDATE SET user_id=excluded.user_id, p256dh=excluded.p256dh, auth=excluded.auth",
            (user["id"], body.endpoint, p256dh, auth, now_iso()),
        )
        return {"ok": True}


class PushResubscribeBody(BaseModel):
    old_endpoint: str
    endpoint: str
    keys: dict


@app.post("/push/resubscribe")
def push_resubscribe(body: PushResubscribeBody):
    """Called by the service worker when the BROWSER replaces a subscription
    on its own (pushsubscriptionchange) — Chrome does this after an update or
    a storage clear. Nothing ran in the page at that moment, so there is no
    logged-in session to authenticate with; the old endpoint is the identity,
    and it only ever moves an existing row to the new address. Without this,
    notifications stopped for that device silently, and stayed stopped until
    someone thought to toggle them off and on.
    """
    p256dh, auth = body.keys.get("p256dh"), body.keys.get("auth")
    if not p256dh or not auth:
        raise HTTPException(400, "ข้อมูลการสมัครรับแจ้งเตือนไม่ถูกต้อง")
    with get_db() as db:
        row = db.execute("SELECT user_id FROM push_subscriptions WHERE endpoint=?",
                         (body.old_endpoint,)).fetchone()
        if not row:
            raise HTTPException(404, "ไม่พบการลงทะเบียนเดิม")
        db.execute("DELETE FROM push_subscriptions WHERE endpoint=?", (body.old_endpoint,))
        db.execute(
            "INSERT INTO push_subscriptions (user_id, endpoint, p256dh, auth, created_at) VALUES (?,?,?,?,?) "
            "ON CONFLICT(endpoint) DO UPDATE SET user_id=excluded.user_id, p256dh=excluded.p256dh, "
            "auth=excluded.auth",
            (row["user_id"], body.endpoint, p256dh, auth, now_iso()))
        log.info("Push subscription renewed by the browser for user %s", row["user_id"])
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
# Combined data endpoint — one request instead of ~14
# ---------------------------------------------------------------------------

@app.get("/app-data")
def app_data(scope: str = "core", task_scope: Optional[str] = None,
             user=Depends(current_user)):
    """Returns everything a screen needs in ONE round-trip instead of the
    frontend firing 14 separate requests every refresh. Each of those was a
    separate serverless invocation + DB round-trip on Vercel/Neon, which is
    what made the app feel constantly busy. This calls the SAME underlying
    endpoint functions (no logic duplicated/forked) and bundles their output.

    scope controls how much to include, so we only fetch what the current tab
    actually shows:
      core      — always: me, tasks, notifications (every role/tab needs these)
      admin_ops — + drivers, vehicles, users, vehicle types, breakdowns, zones,
                  task rules, team chat (the Operation/Settings tabs)
      admin_dash— + dashboard summary (the Dashboard tab)
      driver    — + vehicles, team chat (driver screens)
    """
    # Holding ONE connection open across the whole bundle is what makes the
    # single-round-trip design actually pay off. Each helper below opens its
    # own `with get_db()`, but get_db() is re-entrant, so they all reuse this
    # one instead of each doing its own TLS handshake to the database. This
    # request went from 12 connections to 1 (plus one for authentication).
    with get_db():
        out = {
            "me": me(user),
            "tasks": list_tasks(None, task_scope, user),
            "notifications": get_notifications(user),
            # Small, changes rarely, and needed by the job form and the
            # profile screen — cheaper here than a request of its own.
            "costCenters": list_cost_centers(user),
        }
        if scope == "driver":
            out["vehicles"] = list_vehicles(user)
            out["teamChatMessages"] = get_team_chat_messages(user)
            out["teamChatParticipants"] = team_chat_participants(user)
        elif scope in ("admin_ops", "admin_dash") and user["role"] in ("ADMIN", "DISPATCH"):
            out["drivers"] = list_drivers(user)
            out["vehicles"] = list_vehicles(user)
            out["allUsers"] = list_all_users(user)
            out["vehicleTypes"] = list_vehicle_types(user)
            out["breakdowns"] = list_breakdowns(user)
            out["zones"] = list_zones(user)
            out["teamChatMessages"] = get_team_chat_messages(user)
            out["teamChatParticipants"] = team_chat_participants(user)
            if scope == "admin_dash":
                out["dashboard"] = dashboard(user)
                out["auditLogs"] = get_audit_logs(user)
        return out



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

        avail_drivers = db.execute("SELECT COUNT(*) n FROM users WHERE (users.role='DRIVER' OR (users.role='DISPATCH' AND EXISTS (SELECT 1 FROM driver_licenses dl WHERE dl.driver_id = users.id))) AND driver_status='Available' AND deleted_at IS NULL").fetchone()["n"]
        total_drivers = db.execute("SELECT COUNT(*) n FROM users WHERE (users.role='DRIVER' OR (users.role='DISPATCH' AND EXISTS (SELECT 1 FROM driver_licenses dl WHERE dl.driver_id = users.id))) AND deleted_at IS NULL").fetchone()["n"]
        busy_drivers = db.execute("SELECT COUNT(*) n FROM users WHERE (users.role='DRIVER' OR (users.role='DISPATCH' AND EXISTS (SELECT 1 FROM driver_licenses dl WHERE dl.driver_id = users.id))) AND driver_status='Busy' AND deleted_at IS NULL").fetchone()["n"]
        breakdown_drivers = db.execute("SELECT COUNT(*) n FROM users WHERE (users.role='DRIVER' OR (users.role='DISPATCH' AND EXISTS (SELECT 1 FROM driver_licenses dl WHERE dl.driver_id = users.id))) AND driver_status='Breakdown' AND deleted_at IS NULL").fetchone()["n"]

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
def dashboard_task_volume(days: int = 7, user=Depends(require_ops())):
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


# Bangkok time, because "today" on a dashboard means the warehouse's today.
# Everything is stored in UTC, so a job opened at 06:00 local was landing in
# YESTERDAY's column on every count that used the UTC date.
BKK_OFFSET = timedelta(hours=7)

# Six statuses are too many segments to read on a phone-width bar, so they are
# grouped into the four states anyone looking at a backlog actually asks about.
VOLUME_GROUPS = (
    ("waiting", "รอดำเนินการ", ("Waiting",)),
    ("active", "กำลังดำเนินการ", ("Assigned", "In Progress", "Pause")),
    ("done", "เสร็จสิ้น", ("Completed",)),
    ("cancelled", "ยกเลิก", ("Cancelled",)),
)
_GROUP_OF = {st: key for key, _, sts in VOLUME_GROUPS for st in sts}


def _to_bkk(iso_utc: str) -> Optional[datetime]:
    if not iso_utc:
        return None
    try:
        return datetime.fromisoformat(iso_utc.replace("Z", "").split("+")[0]) + BKK_OFFSET
    except ValueError:
        return None


@app.get("/dashboard/volume-buckets")
def dashboard_volume_buckets(period: str = "today", user=Depends(require_ops())):
    """How much work was opened, when, and where it stands now.

    One bar per hour (today) or per day (this week / this month), stacked by
    what each job's status is now — so the part of the bar that is not yet
    green is the backlog from that slot, in the shape it actually arrived in.
    Both measures come back together: the screen's งาน/พาเลท switch flips
    between them without another round-trip.
    """
    if period not in ("today", "week", "month"):
        raise HTTPException(400, "period ต้องเป็น today, week หรือ month")
    now_bkk = datetime.utcnow() + BKK_OFFSET
    if period == "today":
        start_bkk = now_bkk.replace(hour=0, minute=0, second=0, microsecond=0)
        step, count, fmt = timedelta(hours=1), 24, lambda d: d.strftime("%H")
    else:
        days = 7 if period == "week" else 30
        start_bkk = (now_bkk - timedelta(days=days - 1)).replace(
            hour=0, minute=0, second=0, microsecond=0)
        step, count, fmt = timedelta(days=1), days, lambda d: d.strftime("%-d/%-m")

    buckets = []
    for i in range(count):
        at = start_bkk + step * i
        buckets.append({
            "label": fmt(at),
            "start": at.isoformat(timespec="seconds"),
            "jobs": {g[0]: 0 for g in VOLUME_GROUPS},
            "pallets": {g[0]: 0 for g in VOLUME_GROUPS},
        })

    since_utc = (start_bkk - BKK_OFFSET).isoformat(timespec="seconds") + "Z"
    with get_db() as db:
        rows = db.execute(
            "SELECT created_at, status, pallet_qty FROM tasks WHERE created_at >= ?",
            (since_utc,),
        ).fetchall()
    for r in rows:
        at = _to_bkk(r["created_at"])
        if not at:
            continue
        idx = int((at - start_bkk) / step)
        if not (0 <= idx < count):
            continue
        g = _GROUP_OF.get(r["status"])
        if not g:
            continue
        buckets[idx]["jobs"][g] += 1
        buckets[idx]["pallets"][g] += int(r["pallet_qty"] or 0)

    totals = {m: {g[0]: sum(b[m][g[0]] for b in buckets) for g in VOLUME_GROUPS}
              for m in ("jobs", "pallets")}
    return {
        "period": period,
        "buckets": buckets,
        "totals": totals,
        "groups": [{"key": k, "label": lbl} for k, lbl, _ in VOLUME_GROUPS],
    }


@app.get("/dashboard/overview")
def dashboard_overview(period: str = "today", user=Depends(require_ops())):
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

        total_drivers = db.execute("SELECT COUNT(*) n FROM users WHERE (users.role='DRIVER' OR (users.role='DISPATCH' AND EXISTS (SELECT 1 FROM driver_licenses dl WHERE dl.driver_id = users.id))) AND deleted_at IS NULL").fetchone()["n"]
        busy_drivers = db.execute("SELECT COUNT(*) n FROM users WHERE (users.role='DRIVER' OR (users.role='DISPATCH' AND EXISTS (SELECT 1 FROM driver_licenses dl WHERE dl.driver_id = users.id))) AND driver_status='Busy' AND deleted_at IS NULL").fetchone()["n"]
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
def export_tasks(period: str = "all", user=Depends(require_ops())):
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
                req.get("full_name") or "", req.get("employee_id") or "",
                t["cost_center"] or req.get("cost_center") or "", req.get("contact") or "",
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
    ensure_tables()
    ensure_columns()
    widen_role_constraint()
    ensure_seed_zones()
    ensure_cost_centers()
    try:
        with get_db() as _db:
            sync_master_fleet(_db)
    except Exception as e:
        log.warning("Master fleet sync skipped: %s", e)
    ensure_indexes()
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
    turning an opaque 500 into a readable message.

    `version` comes first: opening /healthz is the one check of which build is
    live that no page cache, service worker or stale tab can get wrong."""
    return {
        "version": APP_VERSION,
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
    def _clean(s):
        # Values pasted into a hosting dashboard very often pick up a trailing
        # space or newline, or get typed with quote marks around them. Those
        # are never meant to be part of the secret, and an exact comparison
        # would reject the right secret with an unhelpful "Invalid secret".
        return (s or "").strip().strip('"').strip("'").strip()

    reset_secret = _clean(os.environ.get("ADMIN_RESET_SECRET"))
    if not reset_secret:
        raise HTTPException(404, "Not found")
    if not hmac.compare_digest(_clean(body.secret).encode(), reset_secret.encode()):
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


_INDEX_HTML = None


@app.get("/")
def index():
    """Serves the app shell with the build version substituted in.

    Read once and held in memory — the file doesn't change between deploys,
    and on serverless each cold start re-reads it anyway.

    Deliberately no-cache: the whole point of a visible version is that the
    person's screen matches the build that's live. A cached shell would keep
    showing (and running) the previous one.
    """
    global _INDEX_HTML
    if _INDEX_HTML is None:
        try:
            with open(os.path.join(FRONTEND_DIR, "index.html"), encoding="utf-8") as f:
                _INDEX_HTML = f.read().replace("__APP_VERSION__", APP_VERSION)
        except OSError:
            return FileResponse(os.path.join(FRONTEND_DIR, "index.html"))
    return Response(content=_INDEX_HTML, media_type="text/html; charset=utf-8",
                    headers={"Cache-Control": "no-cache"})


@app.get("/install")
def install_page():
    """Same app shell, reached by scanning the QR code.

    It is a real path rather than a query string so the QR stays short and the
    link is readable when someone sends it in a chat. The shell notices the
    path and opens the "add to home screen" instructions, which is the step
    people were skipping — and the step iOS requires before notifications work
    at all.
    """
    return index()


@app.get("/service-worker.js")
def service_worker():
    # Served at the root path (not /static/...) so its default scope covers the whole app.
    return FileResponse(os.path.join(FRONTEND_DIR, "service-worker.js"), media_type="application/javascript")


# ---------------------------------------------------------------------------
# App icons and the install manifest
# ---------------------------------------------------------------------------
# iOS only delivers Web Push to a PWA that was really installed to the home
# screen, and that depends entirely on this manifest: no `display: standalone`,
# or a manifest that 404s because the file did not make it into the deployment,
# and notifications are dead on every iPhone with nothing anywhere to say why.
# So the manifest is now built in code — the repo's file is still honoured, it
# just cannot be the single point of failure any more — and the icons it points
# at are drawn here, so they always exist.

BRAND_BG = (0x25, 0x63, 0xEB)       # --amber in the stylesheet
BRAND_FG = (0xFF, 0xFF, 0xFF)

# The forklift mark from the sign-in screen, in its own 24x24 grid:
# (kind, ...) with rects as x,y,w,h and circles as cx,cy,r.
_MARK = [
    ("rect", 2, 15, 7, 2),          # fork tines
    ("rect", 9, 5, 1.6, 12),        # mast
    ("rect", 12.5, 5, 1.6, 12),     # mast
    ("rect", 10.5, 10, 9, 6),       # body
    ("circle", 13, 18.5, 1.8),      # wheel
    ("circle", 18, 18.5, 1.8),      # wheel
]


def _png_bytes(width: int, height: int, pixels: bytearray) -> bytes:
    """Minimal 8-bit RGB PNG. Written by hand because this has to work with
    nothing installed but the standard library."""
    import zlib as _zlib

    def chunk(tag: bytes, data: bytes) -> bytes:
        return (len(data).to_bytes(4, "big") + tag + data
                + (_zlib.crc32(tag + data) & 0xFFFFFFFF).to_bytes(4, "big"))

    raw = bytearray()
    stride = width * 3
    for y in range(height):
        raw.append(0)                               # filter: none
        raw += pixels[y * stride:(y + 1) * stride]
    ihdr = (width.to_bytes(4, "big") + height.to_bytes(4, "big")
            + bytes([8, 2, 0, 0, 0]))               # 8-bit, truecolour
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", _zlib.compress(bytes(raw), 9)) + chunk(b"IEND", b""))


def render_app_icon(size: int) -> bytes:
    """The mark on a full-bleed brand square — full-bleed so it is safe both as
    an Apple touch icon (which turns transparency black) and as an Android
    maskable icon (the mark sits well inside the 80% safe area)."""
    px = bytearray(bytes(BRAND_BG) * (size * size))
    # The glyph takes the middle ~62%, the same proportion as the header mark.
    glyph = size * 0.62
    scale = glyph / 24.0
    ox = oy = (size - glyph) / 2.0

    def put(x: int, y: int, coverage: float = 1.0):
        if not (0 <= x < size and 0 <= y < size) or coverage <= 0:
            return
        i = (y * size + x) * 3
        if coverage >= 1.0:
            px[i:i + 3] = bytes(BRAND_FG)
        else:
            for k in range(3):
                px[i + k] = round(px[i + k] + (BRAND_FG[k] - px[i + k]) * coverage)

    for shape in _MARK:
        if shape[0] == "rect":
            _, x, y, w, h = shape
            x0, y0 = round(ox + x * scale), round(oy + y * scale)
            x1, y1 = round(ox + (x + w) * scale), round(oy + (y + h) * scale)
            for yy in range(y0, max(y1, y0 + 1)):
                for xx in range(x0, max(x1, x0 + 1)):
                    put(xx, yy)
        else:
            _, cx, cy, r = shape
            dcx, dcy, dr = ox + cx * scale, oy + cy * scale, r * scale
            # Anti-aliased, by sampling 4x4 inside each pixel of the circle's
            # own bounding box only — a few hundred pixels, so it stays cheap.
            for yy in range(int(dcy - dr) - 1, int(dcy + dr) + 2):
                for xx in range(int(dcx - dr) - 1, int(dcx + dr) + 2):
                    hits = 0
                    for sy in range(4):
                        for sx in range(4):
                            ddx = xx + (sx + 0.5) / 4 - dcx
                            ddy = yy + (sy + 0.5) / 4 - dcy
                            if ddx * ddx + ddy * ddy <= dr * dr:
                                hits += 1
                    if hits:
                        put(xx, yy, hits / 16.0)
    return _png_bytes(size, size, px)


_ICON_CACHE = {}
ICON_SIZES = (180, 192, 512)


def app_icon_png(size: int) -> bytes:
    if size not in _ICON_CACHE:
        _ICON_CACHE[size] = render_app_icon(size)
    return _ICON_CACHE[size]


@app.get("/icon-{size}.png")
def app_icon(size: int):
    if size not in ICON_SIZES:
        raise HTTPException(404, "ไม่มีไอคอนขนาดนี้")
    return Response(content=app_icon_png(size), media_type="image/png",
                    headers={"Cache-Control": "public, max-age=604800"})


@app.get("/apple-touch-icon.png")
def apple_touch_icon():
    return Response(content=app_icon_png(180), media_type="image/png",
                    headers={"Cache-Control": "public, max-age=604800"})


@app.get("/favicon.ico")
def favicon():
    return Response(content=app_icon_png(192), media_type="image/png",
                    headers={"Cache-Control": "public, max-age=604800"})


BASE_MANIFEST = {
    "name": "CRL ForkliftDispatch",
    "short_name": "ForkliftDispatch",
    "lang": "th",
    "start_url": "/",
    "scope": "/",
    # iOS delivers push only to an installed, standalone PWA.
    "display": "standalone",
    "orientation": "portrait",
    "background_color": "#f5f6f8",
    "theme_color": "#f5f6f8",
    "icons": [
        {"src": "/icon-192.png", "sizes": "192x192", "type": "image/png", "purpose": "any"},
        {"src": "/icon-512.png", "sizes": "512x512", "type": "image/png", "purpose": "any"},
        {"src": "/icon-512.png", "sizes": "512x512", "type": "image/png", "purpose": "maskable"},
    ],
}


def build_manifest():
    """The repo's manifest.json wins wherever it says something, except where
    saying it would stop the app being installable. Returns (manifest, source)
    so /push/status can report which one the phone is actually getting."""
    manifest = dict(BASE_MANIFEST)
    source = "generated"
    try:
        with open(os.path.join(FRONTEND_DIR, "manifest.json"), encoding="utf-8") as fh:
            disk = json.load(fh)
        if isinstance(disk, dict) and disk:
            for k, v in disk.items():
                if v not in (None, "", [], {}):
                    manifest[k] = v
            source = "file"
    except FileNotFoundError:
        log.info("No manifest.json on disk — serving the built-in one")
    except Exception as e:
        log.warning("manifest.json could not be read (%s) — serving the built-in one", e)
    # Two things are not negotiable, because getting them wrong is invisible
    # and costs every iPhone its notifications.
    if manifest.get("display") not in ("standalone", "fullscreen"):
        manifest["display"] = "standalone"
    if not isinstance(manifest.get("icons"), list) or not manifest["icons"]:
        manifest["icons"] = BASE_MANIFEST["icons"]
    return manifest, source


@app.get("/manifest.json")
def manifest():
    data, _ = build_manifest()
    return Response(content=json.dumps(data, ensure_ascii=False),
                    media_type="application/manifest+json",
                    headers={"Cache-Control": "no-cache"})


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
