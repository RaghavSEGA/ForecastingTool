"""Suggest which roster row a submission belongs to.

Evidence, strongest first: identical figures to the row's current values, the same
name once legal suffixes are dropped, one name containing the other, the company
name appearing in a roster email domain, shared words, and similar spelling.
A suggestion is only pre-selected when it's clearly ahead of the runner-up.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from difflib import SequenceMatcher

from extractor import IPS, QUARTERS, Submission
from layout import RosterRow

STOP = {"inc", "llc", "ltd", "limited", "co", "corp", "corporation", "company", "the", "group",
        "dba", "international", "intl", "global", "mfg", "manufacturing", "brands", "and"}
AUTO_MIN, AUTO_MARGIN = 50, 15


@dataclass
class Candidate:
    row: int
    name: str
    score: int
    reason: str


def _tokens(s: str) -> list[str]:
    return [t for t in re.findall(r"[a-z0-9]+", s.lower()) if t not in STOP]


def _compact(s: str) -> str:
    return "".join(_tokens(s))


def _domains(email: str) -> set[str]:
    return {d.split(".")[0] for d in re.findall(r"@([\w.-]+)", email.lower())}


def _same_figures(sub: Submission, row: RosterRow) -> bool:
    return sub.total > 0 and all(abs(sub.value(q, ip) - row.values.get((q, ip), 0.0)) < 0.01
                                 for q in QUARTERS for ip in IPS)


def suggest(sub: Submission, roster: list[RosterRow]) -> list[Candidate]:
    name = sub.company or ""
    a, ta = _compact(name), set(_tokens(name))
    out = []
    for row in roster:
        score, why = 0, []
        if _same_figures(sub, row):
            score += 100
            why.append("identical figures")
        b, tb = _compact(row.name), set(_tokens(row.name))
        if a and b:
            if a == b:
                score += 90
                why.append("same name")
            elif min(len(a), len(b)) >= 4 and (a in b or b in a):
                score += 70
                why.append("name contains")
            else:
                common = {t for t in ta & tb if len(t) >= 3}
                if common:
                    score += int(40 * len(common) / min(len(ta), len(tb)))
                    why.append("shared words")
                ratio = SequenceMatcher(None, a, b).ratio()
                if ratio >= 0.8:
                    score += int(60 * ratio)
                    why.append("similar spelling")
            words = [t for t in ta | {a} if len(t) >= 3]
            if any(w in d for w in words for d in _domains(row.email)):
                score += 50
                why.append("email domain")
            # Divisions: "Centric Brands" + category "Sleepwear" -> roster "Centric Sleep"
            extra = [t for t in tb - ta if len(t) >= 4]
            cats = [t for t in _tokens(sub.category or "") if len(t) >= 4]
            if extra and cats and any(c.startswith(t) or t.startswith(c) for t in extra for c in cats):
                score += 30
                why.append("product category")
        if score:
            out.append(Candidate(row.row, row.name, score, ", ".join(why)))
    return sorted(out, key=lambda c: -c.score)


def auto_pick(cands: list[Candidate]) -> Candidate | None:
    if not cands or cands[0].score < AUTO_MIN:
        return None
    if len(cands) > 1 and cands[0].score - cands[1].score < AUTO_MARGIN:
        return None
    return cands[0]