from pathlib import Path

from openpyxl import Workbook, load_workbook

from dextr_processor import (
    STATUS_CODES,
    apply_selected_changes,
    build_preview_changes,
    build_updates_from_sources,
    extract_person_name,
    normalize_name,
    scan_source_workbook,
)


SHEET = "Crew Not On Duty Report"


def make_workbook(path: Path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.append([
        "Employee", "60 Hours", "Canvassing Day Shift", "Day Shift Notes",
        "Canvassing Night Shift", "Night Shift Notes", "Predicted Hours", "HRs Left"
    ])
    for row in rows:
        ws.append(row)
    source = wb.create_sheet("Source")
    source["A1"] = "do not touch"
    wb.save(path)


def test_extract_person_name_removes_job_classifications_but_keeps_real_parentheses():
    assert extract_person_name("Labelle, Paul QCTO").person_name == "Labelle, Paul"
    assert extract_person_name("Igaal, Warsame (Sam) CSA").person_name == "Igaal, Warsame (Sam)"
    assert extract_person_name("Addison, Aaron CTO (trainee)").person_name == "Addison, Aaron"
    assert extract_person_name("Vargas, (Abreu) Paul CSA").person_name == "Vargas, (Abreu) Paul"


def test_normalize_name_matches_case_spaces_and_comma_spacing():
    assert normalize_name(" Labelle,   Paul ") == normalize_name("labelle , paul")


def test_scan_source_workbook_finds_statuses_only_in_columns_b_to_f(tmp_path):
    source = tmp_path / "source1.xlsx"
    make_workbook(source, [
        ["QCTO", "DNC", None, None, None, None, "DO NOT COPY", "=60-G2"],
        ["Crew Not On Duty Report", None, None, None, None, None, None, None],
        ["Labelle, Paul QCTO", None, "DNC", "call tomorrow", "STB - awaiting update", None, "DO NOT COPY", "=60-G4"],
        ["Applewhite, Jeff CTO", None, None, None, None, None, "DNC in G ignored", "=60-G5"],
        [None, None, None, None, None, None, None, None],
    ])

    rows = scan_source_workbook(source, "Source 1")

    assert len(rows) == 1
    row = rows[0]
    assert row.person_name == "Labelle, Paul"
    assert row.values_by_column == {"C": "DNC", "D": "call tomorrow", "E": "STB - awaiting update"}
    assert row.statuses_by_column == {"C": ["DNC"], "E": ["STB"]}
    assert all(status in STATUS_CODES for status in ["DNC", "STB", "WSIB", "LD"])


def test_newest_nonblank_source_value_wins_and_blanks_do_not_clear(tmp_path):
    oldest = tmp_path / "oldest.xlsx"
    newest = tmp_path / "newest.xlsx"
    make_workbook(oldest, [["Labelle, Paul QCTO", None, "DNC", "old note", None, None, None, None]])
    make_workbook(newest, [["Labelle, Paul QCTO", None, "STB", None, "LD", "new night note", None, None]])

    updates = build_updates_from_sources([oldest, newest], ["Source 1", "Source 2"])
    update = updates[normalize_name("Labelle, Paul")]

    assert update.proposed_values["C"].value == "STB"
    assert update.proposed_values["C"].source_label == "Source 2"
    assert update.proposed_values["D"].value == "old note"
    assert update.proposed_values["D"].source_label == "Source 1"
    assert update.proposed_values["E"].value == "LD"
    assert update.proposed_values["F"].value == "new night note"


def test_preview_matches_target_rows_and_only_columns_b_to_f(tmp_path):
    source = tmp_path / "source.xlsx"
    target = tmp_path / "target.xlsm"
    make_workbook(source, [["Labelle, Paul QCTO", "DNC", None, None, None, None, "source predicted", "source formula"]])
    make_workbook(target, [["Labelle, Paul CTO", "old", None, None, None, None, 44, "=60-G2"]])

    updates = build_updates_from_sources([source], ["Source 1"])
    changes, unmatched = build_preview_changes(target, updates)

    assert unmatched == []
    assert len(changes) == 1
    change = changes[0]
    assert change.employee_name == "Labelle, Paul"
    assert change.target_row == 2
    assert change.column_letter == "B"
    assert change.existing_value == "old"
    assert change.proposed_value == "DNC"


def test_apply_selected_changes_preserves_g_h_formulas_and_source_sheet(tmp_path):
    source = tmp_path / "source.xlsx"
    target = tmp_path / "target.xlsm"
    output = tmp_path / "updated.xlsm"
    make_workbook(source, [["Labelle, Paul QCTO", "DNC", None, None, None, None, "source predicted", "source formula"]])
    make_workbook(target, [["Labelle, Paul CTO", "old", None, None, None, None, 44, "=60-G2"]])

    updates = build_updates_from_sources([source], ["Source 1"])
    changes, _ = build_preview_changes(target, updates)
    apply_selected_changes(target, output, changes)

    wb = load_workbook(output, keep_vba=True, data_only=False)
    ws = wb[SHEET]
    assert ws["A2"].value == "Labelle, Paul CTO"
    assert ws["B2"].value == "DNC"
    assert ws["G2"].value == 44
    assert ws["H2"].value == "=60-G2"
    assert wb["Source"]["A1"].value == "do not touch"
