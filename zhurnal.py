"""Журнал отправленных листков (SQLite), чтобы никому не отправить листок дважды."""
import hashlib
import sqlite3
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent / "rassylka.db"


def open_journal(path=DB_PATH):
    db = sqlite3.connect(path)
    db.row_factory = sqlite3.Row
    db.execute("""
        CREATE TABLE IF NOT EXISTS sent (
            id        INTEGER PRIMARY KEY,
            period    TEXT NOT NULL,     -- ГГГГ-ММ
            employee  TEXT NOT NULL,     -- табельный номер, а если его нет — ФИО
            email     TEXT NOT NULL,
            amount    INTEGER NOT NULL,  -- к выплате, в тыйынах
            file_hash TEXT NOT NULL,     -- чтобы потом можно было доказать, какой именно файл ушёл
            sent_at   TEXT NOT NULL
        )
    """)
    return db


def when_sent(db, period, employee):
    """Когда и на какой адрес уже отправляли листок за этот месяц (или None)."""
    return db.execute(
        "SELECT sent_at, email FROM sent WHERE period = ? AND employee = ? ORDER BY id DESC",
        (period, employee),
    ).fetchone()


def remember(db, period, p, email, file):
    file_hash = hashlib.sha256(file.read_bytes()).hexdigest()
    now = datetime.now().isoformat(timespec="seconds")
    db.execute(
        "INSERT INTO sent (period, employee, email, amount, file_hash, sent_at) VALUES (?, ?, ?, ?, ?, ?)",
        (period, p.key, email, p.k_vyplate, file_hash, now),
    )
    db.commit()
