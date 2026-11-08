"""Local state in SQLite: the mail delta link and what has been done for each message.

Each step (row logged, files filed, hold booked, draft written) is recorded the
moment it succeeds, so a run that dies halfway picks up where it stopped and
never logs a row twice or books two holds for one email.
"""
from __future__ import annotations

import json
import sqlite3
import time
from pathlib import Path
from typing import Optional

SCHEMA = """
create table if not exists meta (key text primary key, value text);
create table if not exists messages (
  message_key text primary key,      -- internetMessageId (stable even if the message moves folder)
  graph_id text, subject text, sender text, received text,
  status text,                        -- logged | skipped | review | error
  steps text default '{}',            -- json: {"files": [...], "event_id": ..., "row": true, "draft_id": ...}
  extraction text, cost_usd real default 0, errors integer default 0, last_error text, updated real
);
"""


class Store:
    def __init__(self, path: str | Path = ".graph-assistant/state.db"):
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(str(path))
        self.db.executescript(SCHEMA)

    def get_meta(self, key: str) -> Optional[str]:
        row = self.db.execute("select value from meta where key=?", (key,)).fetchone()
        return row[0] if row else None

    def set_meta(self, key: str, value: str) -> None:
        self.db.execute("insert or replace into meta values (?,?)", (key, value))
        self.db.commit()

    def get(self, key: str) -> Optional[dict]:
        cur = self.db.execute("select * from messages where message_key=?", (key,))
        row = cur.fetchone()
        if not row:
            return None
        d = dict(zip([c[0] for c in cur.description], row))
        d["steps"] = json.loads(d["steps"] or "{}")
        return d

    def upsert(self, key: str, **fields) -> None:
        if "steps" in fields:
            fields["steps"] = json.dumps(fields["steps"])
        if "extraction" in fields and not isinstance(fields["extraction"], str):
            fields["extraction"] = json.dumps(fields["extraction"])
        fields["updated"] = time.time()
        if self.get(key) is None:
            self.db.execute("insert into messages (message_key) values (?)", (key,))
        sets = ", ".join(f"{k}=?" for k in fields)
        self.db.execute(f"update messages set {sets} where message_key=?", (*fields.values(), key))
        self.db.commit()

    def step(self, key: str, name: str, value) -> None:
        rec = self.get(key) or {"steps": {}}
        steps = rec["steps"]
        steps[name] = value
        self.upsert(key, steps=steps)

    def record_error(self, key: str, err: str) -> int:
        rec = self.get(key)
        n = (rec["errors"] if rec else 0) + 1
        self.upsert(key, status="error", errors=n, last_error=err[:500])
        return n

    def all(self) -> list[dict]:
        cur = self.db.execute("select * from messages order by received")
        cols = [c[0] for c in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
