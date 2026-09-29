"""Durable write-intent journal, containing hashes only, with a run lock."""

import hashlib
import os
import sqlite3
from contextlib import contextmanager
from pathlib import Path

from resilience import Stopped


class Journal:
    def __init__(self, path):
        self.path = Path(path).expanduser().resolve()
        self.path.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
        # Establish restrictive permissions before SQLite opens the file.
        fd = os.open(self.path, os.O_CREAT | os.O_WRONLY, 0o600)
        os.close(fd)
        with self.connect() as db:
            db.execute(
                "CREATE TABLE IF NOT EXISTS pending (key TEXT PRIMARY KEY, destination TEXT NOT NULL)"
            )

    @contextmanager
    def connect(self):
        db = sqlite3.connect(self.path, timeout=30)
        try:
            with db:
                yield db
        finally:
            db.close()

    @staticmethod
    def key(email):
        return hashlib.sha256(email.casefold().encode()).hexdigest()

    def pending(self, email, destination):
        with self.connect() as db:
            return (
                db.execute(
                    "SELECT 1 FROM pending WHERE key=?", (self.key(email),)
                ).fetchone()
                is not None
            )

    def conflicts(self, email, destination):
        with self.connect() as db:
            row = db.execute(
                "SELECT destination FROM pending WHERE key=?", (self.key(email),)
            ).fetchone()
            return row is not None and row[0] != self.key(destination)

    def begin(self, email, destination):
        with self.connect() as db:
            try:
                db.execute(
                    "INSERT INTO pending VALUES (?, ?)",
                    (self.key(email), self.key(destination)),
                )
            except sqlite3.IntegrityError:
                raise Stopped(
                    "An unresolved write already exists; readback required"
                ) from None

    def clear(self, email, destination):
        with self.connect() as db:
            db.execute(
                "DELETE FROM pending WHERE key=? AND destination=?",
                (self.key(email), self.key(destination)),
            )

    def __enter__(self):
        lock_path = str(self.path) + ".lock"
        fd = os.open(lock_path, os.O_CREAT | os.O_WRONLY, 0o600)
        os.close(fd)
        self.run_lock = sqlite3.connect(lock_path, timeout=0)
        try:
            self.run_lock.execute("BEGIN EXCLUSIVE")
        except sqlite3.OperationalError:
            self.run_lock.close()
            raise Stopped(
                "Another process uses this state file; wait for it to exit"
            ) from None
        return self

    def __exit__(self, *_):
        self.run_lock.rollback()
        self.run_lock.close()
