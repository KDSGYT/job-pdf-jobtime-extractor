from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import parser
from parser import parse_pdf, rows_to_csv


SAMPLE_PAGE_TEXT = """
Job no.: 10001
 Mon-Fri Valid from: 2026-07-01
Job start: 06:00
On duty location: UNION Job end: 14:30
Duration: 08:30
Split time: 00:35
Split from 10:10 to 10:45
PDTIME UNION UNION 06:05 06:10 00:05
PDTIM 1234 UNION OAK 07:00 07:10 00:10
STBY WR# WR# 12:17 13:33 01:16
"""


class FakePage:
    def extract_text(self):
        return SAMPLE_PAGE_TEXT


class FakeReader:
    def __init__(self, _path_or_file):
        self.pages = [FakePage()]


def main():
    original_reader = parser.PdfReader
    parser.PdfReader = FakeReader
    try:
        rows = parse_pdf("synthetic.pdf")
    finally:
        parser.PdfReader = original_reader

    assert rows, "expected rows to be extracted"
    pd_rows = [r for r in rows if r.record_type == "PD time"]
    stby_rows = [r for r in rows if r.record_type == "STBY"]
    split_rows = [r for r in rows if r.record_type == "Split"]

    assert len(pd_rows) == 2, f"expected 2 PD time rows, got {len(pd_rows)}"
    assert len(stby_rows) == 1, f"expected 1 STBY row, got {len(stby_rows)}"
    assert len(split_rows) == 1, f"expected 1 Split row, got {len(split_rows)}"

    first = rows[0]
    assert first.job_number == "10001", f"unexpected first job {first.job_number}"
    assert first.valid_days == "Mon-Fri", f"unexpected valid days {first.valid_days}"
    assert first.job_start == "06:00"
    assert first.job_end == "14:30"
    assert first.job_duration == "08:30"
    assert first.on_duty_location == "UNION"

    assert pd_rows[0].start == "06:05"
    assert pd_rows[0].end == "06:10"
    assert pd_rows[0].duration == "00:05"
    assert pd_rows[1].train_number == "1234"
    assert pd_rows[1].from_location == "UNION"
    assert pd_rows[1].to_location == "OAK"

    assert split_rows[0].start == "10:10"
    assert split_rows[0].end == "10:45"
    assert split_rows[0].duration == "00:35"
    assert rows_to_csv(rows).startswith("job_number,valid_days,page")

    print("OK")
    print(f"total records: {len(rows)}")
    print(f"PD time: {len(pd_rows)}")
    print(f"STBY: {len(stby_rows)}")
    print(f"Split: {len(split_rows)}")
    print(f"jobs with extracted records: {len(set(r.job_number for r in rows))}")


if __name__ == "__main__":
    main()
