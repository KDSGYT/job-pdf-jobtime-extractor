from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from table_state import filter_rows, sort_rows, time_to_minutes


ROWS = [
    {
        "job_number": "10001",
        "record_type": "PD time",
        "start": "09:00",
        "duration": "10:00",
        "on_duty_location": "WR#",
    },
    {
        "job_number": "10002",
        "record_type": "STBY",
        "start": "08:30",
        "duration": "02:00",
        "on_duty_location": "WB#",
    },
    {
        "job_number": "10003",
        "record_type": "Split",
        "start": "10:15",
        "duration": "09:30",
        "on_duty_location": "MSY",
    },
]


def main():
    assert time_to_minutes("10:00") == 600
    assert time_to_minutes("9:30") == 570
    assert time_to_minutes("") == -1

    assert [r["job_number"] for r in sort_rows(ROWS, "duration")] == ["10002", "10003", "10001"]
    assert [r["job_number"] for r in sort_rows(ROWS, "start")] == ["10002", "10001", "10003"]
    assert [r["job_number"] for r in sort_rows(ROWS, "on_duty_location", reverse=True)] == ["10001", "10002", "10003"]

    filtered = filter_rows(ROWS, "stby wb", "STBY")
    assert len(filtered) == 1
    assert filtered[0]["job_number"] == "10002"

    assert ROWS[0]["job_number"] == "10001", "helpers should not mutate source rows"
    print("OK")


if __name__ == "__main__":
    main()
