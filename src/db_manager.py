"""
db_manager.py
─────────────────────────────────────────────────────
Shared database layer for the Door-Sign-Status system.
Both display_app.py and control_app.py import this module.
"""

import sqlite3
import os
import sys
from datetime import datetime
from pathlib import Path

DB_NAME = "door_status.db"


def _db_path() -> str:
    """Return the absolute path of the database file.
    When frozen as .exe, use the folder where the .exe lives.
    When running as .py, use the folder where this script lives.
    """
    if getattr(sys, 'frozen', False):
        # PyInstaller .exe — database goes next to the .exe file
        base = os.path.dirname(os.path.abspath(sys.executable))
    else:
        # Normal Python — database goes next to this script
        base = str(Path(__file__).resolve().parent)
    return os.path.join(base, DB_NAME)


def get_connection() -> sqlite3.Connection:
    """Create and return a new connection with row-factory enabled."""
    conn = sqlite3.connect(_db_path(), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")  # better concurrency
    return conn


# ────────────────────────── Schema Bootstrap ──────────────────────────

def init_db():
    """Ensure all required tables exist.  Safe to call multiple times."""
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS status (
            id                    INTEGER PRIMARY KEY CHECK (id = 1),
            current_status        TEXT    NOT NULL DEFAULT 'Available',
            priority              INTEGER NOT NULL DEFAULT 1,
            return_time           TEXT    DEFAULT '',
            source                TEXT    DEFAULT 'manual',
            last_manual_status    TEXT    DEFAULT 'Available',
            last_manual_return_time TEXT  DEFAULT '',
            last_updated          TEXT    DEFAULT (datetime('now','localtime'))
        )
    """)

    # --- Migration: Add columns if they don't exist ---
    try:
        cur.execute("ALTER TABLE status ADD COLUMN last_manual_status TEXT DEFAULT 'Available'")
    except sqlite3.OperationalError: pass
    try:
        cur.execute("ALTER TABLE status ADD COLUMN last_manual_return_time TEXT DEFAULT ''")
    except sqlite3.OperationalError: pass

    cur.execute("""
        CREATE TABLE IF NOT EXISTS visitors (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            name      TEXT    NOT NULL,
            purpose   TEXT    NOT NULL,
            timestamp TEXT    DEFAULT (datetime('now','localtime'))
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS timetable (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            day_of_week TEXT    NOT NULL,  -- MONDAY, TUESDAY, etc.
            event_name  TEXT    NOT NULL,
            start_time  TEXT    NOT NULL,  -- HH:MM
            end_time    TEXT    NOT NULL   -- HH:MM
        )
    """)

    # --- Migration: Add day_of_week column if it doesn't exist ---
    try:
        cur.execute("ALTER TABLE timetable ADD COLUMN day_of_week TEXT DEFAULT 'ALL'")
    except sqlite3.OperationalError: pass

    # Seed the status row if it doesn't exist yet
    cur.execute("SELECT COUNT(*) FROM status")
    if cur.fetchone()[0] == 0:
        cur.execute(
            "INSERT INTO status (id, current_status, priority, return_time, source, last_manual_status) "
            "VALUES (1, 'Available', 1, '', 'manual', 'Available')"
        )

    conn.commit()
    conn.close()


# ────────────────────────── Status Helpers ────────────────────────────

# Priority map — higher number = stricter / harder to override
STATUS_PRIORITY = {
    "Available":         1,
    "Back soon":         2,
    "Working remotely":  3,
    "Please Knock":      4,
    "In a Meeting":      6,
    "Busy":              7,
    "Do Not Disturb":    8,
    # Outlook-driven statuses use elevated priorities
    "In a Meeting (Outlook)":   9,
    "Do Not Disturb (Outlook)": 10,
}


def get_priority(status_text: str) -> int:
    """Return a numeric priority for a status string."""
    # Strip any " — back by …" suffix before lookup
    base = status_text.split(" — ")[0].strip()
    return STATUS_PRIORITY.get(base, 5)  # default 5 for custom messages


def read_status() -> dict:
    """Return the current status row as a dict."""
    conn = get_connection()
    row = conn.execute("SELECT * FROM status WHERE id = 1").fetchone()
    conn.close()
    if row is None:
        return {
            "current_status": "Available",
            "priority": 1,
            "return_time": "",
            "source": "manual",
            "last_manual_status": "Available",
            "last_manual_return_time": "",
            "last_updated": "",
        }
    return dict(row)


def write_status(status_text: str, return_time: str = "", source: str = "manual"):
    """Persist a new status to the database."""
    priority = get_priority(status_text)
    display = status_text
    if return_time:
        display = f"{status_text} — back by {return_time}"

    conn = get_connection()
    if source == "manual":
        # Save manual status as the 'last known' for auto-restore
        conn.execute(
            """
            UPDATE status
            SET current_status = ?,
                priority       = ?,
                return_time    = ?,
                source         = ?,
                last_manual_status = ?,
                last_manual_return_time = ?,
                last_updated   = datetime('now','localtime')
            WHERE id = 1
            """,
            (display, priority, return_time, source, display, return_time),
        )
    else:
        # Outlook or other auto-source — don't overwrite manual memory
        conn.execute(
            """
            UPDATE status
            SET current_status = ?,
                priority       = ?,
                return_time    = ?,
                source         = ?,
                last_updated   = datetime('now','localtime')
            WHERE id = 1
            """,
            (display, priority, return_time, source),
        )
    conn.commit()
    conn.close()


def update_return_time(new_time: str):
    """Update only the return time on the current status without changing it."""
    row = read_status()
    # Strip any existing " — back by …" suffix to get the base status
    base_status = row["current_status"].split(" — ")[0].strip()

    display = base_status
    if new_time:
        display = f"{base_status} — back by {new_time}"

    conn = get_connection()
    conn.execute(
        """
        UPDATE status
        SET current_status = ?,
            return_time    = ?,
            last_updated   = datetime('now','localtime')
        WHERE id = 1
        """,
        (display, new_time),
    )
    conn.commit()
    conn.close()


# ────────────────────────── Visitor Helpers ────────────────────────────

def add_visitor(name: str, purpose: str):
    """Insert a new visitor message."""
    conn = get_connection()
    conn.execute(
        "INSERT INTO visitors (name, purpose) VALUES (?, ?)",
        (name, purpose),
    )
    conn.commit()
    conn.close()


def get_visitors(limit: int = 100) -> list[dict]:
    """Return recent visitor messages, newest first."""
    conn = get_connection()
    rows = conn.execute(
        "SELECT id, name, purpose, timestamp FROM visitors ORDER BY timestamp DESC LIMIT ?",
        (limit,),
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def delete_visitor(visitor_id: int):
    """Delete a single visitor record by id."""
    conn = get_connection()
    conn.execute("DELETE FROM visitors WHERE id = ?", (visitor_id,))
    conn.commit()
    conn.close()


def visitor_count() -> int:
    """Return the total number of unread visitor messages."""
    conn = get_connection()
    c = conn.execute("SELECT COUNT(*) FROM visitors").fetchone()[0]
    conn.close()
    return c


# ────────────────────────── Timetable Helpers ──────────────────────────

def save_timetable(data_list: list[dict]):
    """Overwrite the existing timetable with new data."""
    conn = get_connection()
    conn.execute("DELETE FROM timetable")
    for item in data_list:
        conn.execute(
            "INSERT INTO timetable (day_of_week, event_name, start_time, end_time) VALUES (?, ?, ?, ?)",
            (item["day"], item["name"], item["start"], item["end"])
        )
    conn.commit()
    conn.close()


def get_timetable(day: str = None) -> list[dict]:
    """Return the daily timetable (optional filter by day)."""
    conn = get_connection()
    if day:
        rows = conn.execute(
            "SELECT day_of_week, event_name, start_time, end_time FROM timetable WHERE day_of_week = ? OR day_of_week = 'ALL'",
            (day,)
        ).fetchall()
    else:
        rows = conn.execute("SELECT day_of_week, event_name, start_time, end_time FROM timetable").fetchall()
    conn.close()
    return [{"day": r["day_of_week"], "name": r["event_name"], "start": r["start_time"], "end": r["end_time"]} for r in rows]


def clear_timetable():
    """Remove all events from the daily timetable."""
    conn = get_connection()
    conn.execute("DELETE FROM timetable")
    conn.commit()
    conn.close()
