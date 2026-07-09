from __future__ import annotations

from typing import Iterable

TIME_FIELDS = {"duration", "start", "end", "job_start", "job_end", "job_duration"}


def time_to_minutes(value: object) -> int:
    """Convert H:MM / HH:MM text to minutes for stable time sorting."""
    text = str(value or "").strip()
    if not text or ":" not in text:
        return -1
    hours, minutes = text.split(":", 1)
    if not hours.isdigit() or not minutes.isdigit():
        return -1
    return int(hours) * 60 + int(minutes)


def filter_rows(rows: Iterable[dict], query: str = "", record_type: str = "All") -> list[dict]:
    """Filter already-extracted rows in memory; never reparses the PDF."""
    words = [word.lower() for word in query.split()]
    filtered = list(rows)
    if record_type and record_type != "All":
        filtered = [row for row in filtered if row.get("record_type") == record_type]
    if words:
        filtered = [
            row
            for row in filtered
            if all(word in " ".join(str(value).lower() for value in row.values()) for word in words)
        ]
    return filtered


def sort_rows(rows: Iterable[dict], column: str, reverse: bool = False) -> list[dict]:
    """Sort rows without mutating source data; time-ish columns sort by minutes."""
    def key(row: dict):
        value = row.get(column, "")
        if column in TIME_FIELDS:
            return time_to_minutes(value)
        return str(value or "").lower()

    return sorted(rows, key=key, reverse=reverse)
