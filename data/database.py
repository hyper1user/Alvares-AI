"""
SQLite база даних для управління ролями особового складу.
Файл app.db створюється в директорії проєкту.
"""
import sqlite3
import os
from datetime import datetime
from typing import List, Tuple, Optional, Dict

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app.db")

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS personnel (
    pib TEXT PRIMARY KEY,
    rank TEXT NOT NULL DEFAULT '',
    position TEXT NOT NULL DEFAULT '',
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS roles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL
);

CREATE TABLE IF NOT EXISTS personnel_roles (
    pib TEXT PRIMARY KEY,
    role_id INTEGER,
    FOREIGN KEY (pib) REFERENCES personnel(pib) ON DELETE CASCADE,
    FOREIGN KEY (role_id) REFERENCES roles(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT
);
"""

DEFAULT_ROLES = [
    "Заступник командира роти",
    "Офіцер з МПЗ",
    "Старший технік роти",
    "Головний сержант роти",
    "Сержант із матеріального забезпечення",
    "Старший бойовий медик",
    "Група евакуації",
    "Водій групи евакуації",
    "Екіпажі розрахунків БМП-1ЛБ",
    "Командири штурмових взводів",
    "Водії роти",
    "Чергові зв'язківці",
    "Резервні групи",
]


def get_connection() -> sqlite3.Connection:
    """Повертає нове з'єднання з WAL mode та foreign keys."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    return conn


def init_db() -> None:
    """Створює таблиці та вставляє ролі за замовчуванням."""
    conn = get_connection()
    try:
        conn.executescript(SCHEMA_SQL)
        for role_name in DEFAULT_ROLES:
            conn.execute("INSERT OR IGNORE INTO roles (name) VALUES (?)", (role_name,))
        conn.commit()
    finally:
        conn.close()


def upsert_personnel(pib: str, rank: str, position: str) -> None:
    """Вставляє або оновлює запис про військовослужбовця."""
    conn = get_connection()
    try:
        conn.execute(
            "INSERT OR REPLACE INTO personnel (pib, rank, position, updated_at) VALUES (?, ?, ?, datetime('now'))",
            (pib.strip(), rank.strip(), position.strip())
        )
        conn.commit()
    finally:
        conn.close()


def upsert_personnel_batch(records: List[Tuple[str, str, str]]) -> int:
    """Масовий upsert. Повертає кількість оброблених записів."""
    conn = get_connection()
    try:
        cleaned = [(pib.strip(), rank.strip(), pos.strip()) for pib, rank, pos in records if pib.strip()]
        conn.executemany(
            "INSERT OR REPLACE INTO personnel (pib, rank, position, updated_at) VALUES (?, ?, ?, datetime('now'))",
            cleaned
        )
        conn.commit()
        return len(cleaned)
    finally:
        conn.close()


def get_all_personnel() -> List[Dict]:
    """Повертає список усіх з LEFT JOIN на ролі."""
    conn = get_connection()
    try:
        rows = conn.execute("""
            SELECT p.pib, p.rank, p.position, pr.role_id, r.name AS role_name
            FROM personnel p
            LEFT JOIN personnel_roles pr ON p.pib = pr.pib
            LEFT JOIN roles r ON pr.role_id = r.id
            ORDER BY p.pib
        """).fetchall()
        return [dict(row) for row in rows]
    finally:
        conn.close()


def get_all_roles() -> List[Tuple[int, str]]:
    """Повертає [(id, name), ...]."""
    conn = get_connection()
    try:
        rows = conn.execute("SELECT id, name FROM roles ORDER BY id").fetchall()
        return [(row["id"], row["name"]) for row in rows]
    finally:
        conn.close()


def set_personnel_role(pib: str, role_id: Optional[int]) -> None:
    """Призначає роль. role_id=None видаляє призначення."""
    conn = get_connection()
    try:
        if role_id is None:
            conn.execute("DELETE FROM personnel_roles WHERE pib = ?", (pib,))
        else:
            conn.execute(
                "INSERT OR REPLACE INTO personnel_roles (pib, role_id) VALUES (?, ?)",
                (pib, role_id)
            )
        conn.commit()
    finally:
        conn.close()


def get_personnel_by_role(role_id: int) -> List[Dict]:
    """Повертає людей з конкретною роллю."""
    conn = get_connection()
    try:
        rows = conn.execute("""
            SELECT p.pib, p.rank, p.position
            FROM personnel p
            JOIN personnel_roles pr ON p.pib = pr.pib
            WHERE pr.role_id = ?
            ORDER BY p.pib
        """, (role_id,)).fetchall()
        return [dict(row) for row in rows]
    finally:
        conn.close()


def get_role_composition() -> Dict[str, List[Dict]]:
    """Повертає {role_name: [{pib, rank, position}, ...]} для всіх ролей."""
    conn = get_connection()
    try:
        roles = conn.execute("SELECT id, name FROM roles ORDER BY id").fetchall()
        result = {}
        for role in roles:
            rows = conn.execute("""
                SELECT p.pib, p.rank, p.position
                FROM personnel p
                JOIN personnel_roles pr ON p.pib = pr.pib
                WHERE pr.role_id = ?
                ORDER BY p.pib
            """, (role["id"],)).fetchall()
            result[role["name"]] = [dict(r) for r in rows]
        return result
    finally:
        conn.close()


def get_setting(key: str, default: str = "") -> str:
    """Отримує налаштування за ключем."""
    conn = get_connection()
    try:
        row = conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
        return row["value"] if row else default
    finally:
        conn.close()


def set_setting(key: str, value: str) -> None:
    """Зберігає налаштування."""
    conn = get_connection()
    try:
        conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, value))
        conn.commit()
    finally:
        conn.close()
