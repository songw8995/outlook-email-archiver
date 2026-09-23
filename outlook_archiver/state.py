from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

SCHEMA_VERSION = 1


class ArchiveState:
    def __init__(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.path = path
        self.connection = sqlite3.connect(path, timeout=10)
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA journal_mode=WAL")
        self.connection.execute("PRAGMA foreign_keys=ON")
        self._init_schema()

    def _init_schema(self) -> None:
        self.connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS messages (
                source_key TEXT NOT NULL,
                version INTEGER NOT NULL,
                store_id TEXT NOT NULL,
                folder_id TEXT NOT NULL,
                entry_id TEXT NOT NULL,
                internet_message_id TEXT,
                content_signature TEXT NOT NULL,
                status TEXT NOT NULL,
                archive_path TEXT NOT NULL,
                manifest_sha256 TEXT,
                updated_at TEXT NOT NULL,
                last_error TEXT,
                PRIMARY KEY (source_key, version)
            );
            CREATE INDEX IF NOT EXISTS idx_messages_source ON messages(source_key, version DESC);
            CREATE TABLE IF NOT EXISTS runs (
                run_id TEXT PRIMARY KEY,
                started_at TEXT NOT NULL,
                finished_at TEXT,
                mode TEXT NOT NULL,
                selected_count INTEGER NOT NULL DEFAULT 0,
                completed_count INTEGER NOT NULL DEFAULT 0,
                status TEXT NOT NULL,
                report_path TEXT
            );
            """
        )
        current = self.get_setting("schema_version")
        if current is None:
            self.set_setting("schema_version", str(SCHEMA_VERSION))
        elif int(current) != SCHEMA_VERSION:
            raise RuntimeError(f"不支持的状态数据库版本：{current}")
        self.connection.commit()

    def close(self) -> None:
        self.connection.close()

    def __enter__(self) -> "ArchiveState":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    @contextmanager
    def transaction(self) -> Iterator[sqlite3.Connection]:
        try:
            self.connection.execute("BEGIN IMMEDIATE")
            yield self.connection
            self.connection.commit()
        except Exception:
            self.connection.rollback()
            raise

    def get_setting(self, key: str) -> str | None:
        row = self.connection.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return str(row[0]) if row else None

    def set_setting(self, key: str, value: str) -> None:
        self.connection.execute(
            "INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value),
        )
        self.connection.commit()

    def latest(self, source_key: str):
        return self.connection.execute(
            "SELECT * FROM messages WHERE source_key=? ORDER BY version DESC LIMIT 1", (source_key,)
        ).fetchone()

    def next_version(self, source_key: str) -> int:
        row = self.latest(source_key)
        return (int(row["version"]) + 1) if row else 1

    def upsert_message(
        self,
        *,
        source_key: str,
        version: int,
        store_id: str,
        folder_id: str,
        entry_id: str,
        internet_message_id: str,
        content_signature: str,
        status: str,
        archive_path: str,
        manifest_sha256: str,
        updated_at: str,
        last_error: str = "",
    ) -> None:
        self.connection.execute(
            """
            INSERT INTO messages VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(source_key,version) DO UPDATE SET
              content_signature=excluded.content_signature,
              status=excluded.status,
              archive_path=excluded.archive_path,
              manifest_sha256=excluded.manifest_sha256,
              updated_at=excluded.updated_at,
              last_error=excluded.last_error
            """,
            (
                source_key, version, store_id, folder_id, entry_id, internet_message_id,
                content_signature, status, archive_path, manifest_sha256, updated_at, last_error,
            ),
        )
        self.connection.commit()

