from pathlib import Path


APP_SOURCE = Path(__file__).resolve().parents[1] / "streamlit_app.py"
README = Path(__file__).resolve().parents[1] / "README.md"


def test_streamlit_app_shows_explicit_numbered_workflow_steps():
    source = APP_SOURCE.read_text()

    expected_step_labels = [
        "Step-by-step workflow",
        "1. Choose the seven source workbooks",
        "2. Confirm source chronology",
        "3. Choose the eighth workbook — target",
        "4. Build preview",
        "5. Review changes before update",
        "6. Create updated workbook copy",
        "7. Download updated workbook",
        "8. Consolidated report",
        "9. Employees not found in target",
    ]

    for label in expected_step_labels:
        assert label in source


def test_streamlit_app_loads_saved_source_files_by_default():
    source = APP_SOURCE.read_text()

    assert "REFERENCE_SOURCE_DIR = Path(__file__).parent / \"reference_source_pdfs\"" in source
    assert "def reference_source_files()" in source
    assert "Use saved source files from the repo" in source
    assert "Loaded 7 saved source files from reference_source_pdfs." in source
    assert "source_item_path(source, tmp, i)" in source
    assert "source_item_name(source)" in source


def test_streamlit_app_loads_saved_target_file_by_default():
    source = APP_SOURCE.read_text()

    assert "REFERENCE_TARGET_DIR = Path(__file__).parent / \"reference_target_workbook\"" in source
    assert "def reference_target_file()" in source
    assert "Use saved eighth target workbook from the repo" in source
    assert "Loaded saved target workbook from reference_target_workbook." in source
    assert "target_item_path(target_workbook, tmp)" in source
    assert "target_item_bytes(target_workbook)" in source
    assert "target_item_name(target_workbook)" in source


def test_streamlit_app_documents_sto_reo_gsr_row_removal_rule():
    source = APP_SOURCE.read_text()

    assert "Removes target rows when column A contains STO, REO, or GSR" in source
    assert "Does not add, delete, or reorder rows outside the QCTO, CTO, CSA, and Trainee target tables" in source


def test_streamlit_app_documents_hrs_left_formula_sort_and_grey_zero_hour_rules():
    source = APP_SOURCE.read_text()

    assert "Adds target column H header `HRs Left`" in source
    assert "formulas `=60-G[row]` for QCTO, CTO, CSA, and Trainee rows" in source
    assert "Sorts rows inside the QCTO, CTO, CSA, and Trainee target tables by `HRs Left`, smallest to largest" in source
    assert "Colors rows grey only when their predicted hours are `0.00` or columns B through F contain `DNC`; `LVM@` time values and Easter, Douglas QCTO are not greyed out" in source


def test_readme_documents_full_workbook_flow():
    readme = README.read_text()

    assert "7. Click `Create updated workbook`." in readme
    assert "8. Download the updated `.xlsm` workbook." in readme
    assert "9. Download the consolidated report." in readme
    assert "10. Check any unmatched employees before using the updated file." in readme
