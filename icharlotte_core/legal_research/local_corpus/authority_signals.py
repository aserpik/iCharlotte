"""Build soft good-law signals: inbound citation count + latest citing year.

NOT a Shepard's/KeyCite check — only a staleness hint. Inverts each case's
`cites_to` across the corpus, matching by normalized reporter citation.
"""
from __future__ import annotations

import json
import sqlite3


def _norm(cite: str) -> str:
    return (cite or "").replace(" ", "").lower()


def build_signals(con: sqlite3.Connection) -> None:
    # Map normalized citation -> case_uid for resolvable targets.
    cite_to_uid: dict[str, str] = {}
    for row in con.execute("SELECT case_uid, citation FROM cases"):
        n = _norm(row["citation"])
        if n:
            cite_to_uid.setdefault(n, row["case_uid"])

    counts: dict[str, int] = {}
    latest: dict[str, str] = {}
    for row in con.execute("SELECT case_uid, year, cites_to FROM cases"):
        citing_year = row["year"] or ""
        try:
            targets = json.loads(row["cites_to"] or "[]")
        except (TypeError, ValueError):
            targets = []
        for tgt in targets:
            uid = cite_to_uid.get(_norm(tgt))
            if not uid:
                continue
            counts[uid] = counts.get(uid, 0) + 1
            if citing_year > latest.get(uid, ""):
                latest[uid] = citing_year

    for uid, n in counts.items():
        con.execute(
            "UPDATE cases SET citation_count=?, latest_citing_year=? WHERE case_uid=?",
            (n, latest.get(uid, ""), uid),
        )
    con.commit()
