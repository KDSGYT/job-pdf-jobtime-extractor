from __future__ import annotations

import csv
import io
import re
from dataclasses import dataclass, asdict
from typing import Iterable, List, Optional

from pypdf import PdfReader

TIME_RE = r"\d{1,2}:\d{2}"


@dataclass
class ExtractedRow:
    job_number: str
    valid_days: str
    page: int
    job_start: str
    job_end: str
    job_duration: str
    on_duty_location: str
    record_type: str  # PD time, STBY, Split
    train_number: str
    from_location: str
    to_location: str
    start: str
    end: str
    duration: str
    raw_line: str


def _first(pattern: str, text: str, default: str = "") -> str:
    match = re.search(pattern, text, flags=re.IGNORECASE | re.MULTILINE)
    return match.group(1).strip() if match else default


def _normalize_activity(activity: str) -> str:
    upper = activity.upper().rstrip(":")
    if upper in {"PDTIME", "PDTIM", "PD"}:
        return "PD time"
    if upper == "STBY":
        return "STBY"
    return activity


def _extract_valid_days(text: str) -> str:
    # Examples: " Mon-Fri Valid from", " Friday ONLY Valid from", " Mon-Thu ONLY Valid from"
    return _first(r"^\s*([^\n]*?(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)[^\n]*?)\s+Valid\s+from:", text)


def _parse_activity_line(line: str) -> Optional[dict]:
    """Parse a row like: PDTIME UNION UNION 06:05 06:10 00:05.

    The PDFs are extracted as plain text, so columns are not perfectly preserved.
    This parser works from the right side because Start, End, Duration are always
    the final three time fields.
    """
    clean = " ".join(line.strip().split())
    if not clean:
        return None

    match = re.search(rf"^(.*?)\s+({TIME_RE})\s+({TIME_RE})\s+({TIME_RE})$", clean)
    if not match:
        return None

    left, start, end, duration = match.groups()
    parts = left.split()
    if not parts:
        return None

    activity = parts[0]
    if activity.upper().rstrip(":") not in {"PDTIME", "PDTIM", "STBY"}:
        return None

    # Common shapes:
    # PDTIME UNION UNION 06:05 06:10 00:05
    # STBY WR# WR# 12:17 13:33 01:16
    # If a train number appears, it usually sits before from/to.
    train_number = ""
    from_location = ""
    to_location = ""
    if len(parts) >= 3:
        from_location = parts[-2]
        to_location = parts[-1]
        maybe_train = parts[-3] if len(parts) >= 4 else ""
        if re.search(r"\d", maybe_train):
            train_number = maybe_train
    elif len(parts) == 2:
        from_location = parts[1]
        to_location = parts[1]

    return {
        "record_type": _normalize_activity(activity),
        "train_number": train_number,
        "from_location": from_location,
        "to_location": to_location,
        "start": start,
        "end": end,
        "duration": duration,
        "raw_line": clean,
    }


def parse_pdf(path_or_file) -> List[ExtractedRow]:
    reader = PdfReader(path_or_file)
    rows: List[ExtractedRow] = []

    for page_index, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        job_number = _first(r"Job\s+no\.?:\s*(\d+)", text)
        if not job_number:
            continue

        job_start = _first(r"Job\s+start:\s*(\d{1,2}:\d{2})", text)
        job_end = _first(r"Job\s+end:\s*(\d{1,2}:\d{2})", text)
        job_duration = _first(r"Duration:\s*(\d{1,2}:\d{2})", text)
        on_duty_location = _first(r"On\s+duty\s+location:\s*([^\n]+?)\s+Job\s+end:", text)
        valid_days = _extract_valid_days(text)

        split_duration = _first(r"Split\s+time:\s*(\d{1,2}:\d{2})", text)
        split_match = re.search(rf"Split\s+from\s+({TIME_RE})\s+to\s+({TIME_RE})", text, flags=re.IGNORECASE)
        if split_match or (split_duration and split_duration != "00:00"):
            rows.append(
                ExtractedRow(
                    job_number=job_number,
                    valid_days=valid_days,
                    page=page_index,
                    job_start=job_start,
                    job_end=job_end,
                    job_duration=job_duration,
                    on_duty_location=on_duty_location,
                    record_type="Split",
                    train_number="",
                    from_location="",
                    to_location="",
                    start=split_match.group(1) if split_match else "",
                    end=split_match.group(2) if split_match else "",
                    duration=split_duration,
                    raw_line=split_match.group(0) if split_match else f"Split time: {split_duration}",
                )
            )

        for line in text.splitlines():
            parsed = _parse_activity_line(line)
            if not parsed:
                continue
            rows.append(
                ExtractedRow(
                    job_number=job_number,
                    valid_days=valid_days,
                    page=page_index,
                    job_start=job_start,
                    job_end=job_end,
                    job_duration=job_duration,
                    on_duty_location=on_duty_location,
                    **parsed,
                )
            )

    return rows


def rows_to_dicts(rows: Iterable[ExtractedRow]) -> list[dict]:
    return [asdict(row) for row in rows]


def rows_to_csv(rows: Iterable[ExtractedRow]) -> str:
    output = io.StringIO()
    fieldnames = list(ExtractedRow.__dataclass_fields__.keys())
    writer = csv.DictWriter(output, fieldnames=fieldnames)
    writer.writeheader()
    for row in rows:
        writer.writerow(asdict(row))
    return output.getvalue()
