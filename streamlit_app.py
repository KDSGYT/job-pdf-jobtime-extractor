from __future__ import annotations

import io
import tempfile
from dataclasses import replace
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook

from dextr_processor import (
    REPORT_SHEET_NAME,
    STATUS_CODES,
    TRANSFER_COLUMNS,
    PreviewChange,
    apply_selected_changes,
    build_preview_changes,
    build_updates_from_sources,
    consolidated_report_rows,
    target_name_index,
)

st.set_page_config(page_title="DEXTR DNC Workbook Webapp", page_icon="📘", layout="wide")

APP_CSS = """
<style>
.block-container { padding-top: 1.4rem; }
.small-muted { color: #666; font-size: 0.9rem; }
.warning-box {
  background: #fff7ed; border: 1px solid #fed7aa; border-radius: 10px;
  padding: 0.75rem 1rem; margin: 0.5rem 0 1rem 0;
}
.good-box {
  background: #ecfdf5; border: 1px solid #a7f3d0; border-radius: 10px;
  padding: 0.75rem 1rem; margin: 0.5rem 0 1rem 0;
}
.step-box {
  background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 10px;
  padding: 0.75rem 1rem; margin: 0.5rem 0 1rem 0;
}
</style>
"""
st.markdown(APP_CSS, unsafe_allow_html=True)

WORKFLOW_STEPS = [
    "Use the saved seven source workbooks or upload a replacement source set.",
    "Confirm or adjust the source order from oldest to newest.",
    "Use the saved eighth target workbook or upload a replacement target workbook.",
    "Build the preview so DEXTR can find only safe empty-cell updates.",
    "Review the preview table, edit proposed values if needed, and deselect anything to skip.",
    "Create the updated workbook copy.",
    "Download the updated workbook.",
    "Download the consolidated report.",
    "Check any unmatched employees before using the updated file.",
]

REFERENCE_SOURCE_DIR = Path(__file__).parent / "reference_source_pdfs"
REFERENCE_TARGET_DIR = Path(__file__).parent / "reference_target_workbook"
SOURCE_EXTENSIONS = (".xlsx", ".xlsm")


def reference_source_files() -> list[Path]:
    return sorted(
        path for path in REFERENCE_SOURCE_DIR.iterdir()
        if path.is_file() and path.suffix.lower() in SOURCE_EXTENSIONS
    ) if REFERENCE_SOURCE_DIR.exists() else []


def reference_target_file() -> Path | None:
    targets = sorted(
        path for path in REFERENCE_TARGET_DIR.iterdir()
        if path.is_file() and path.suffix.lower() in SOURCE_EXTENSIONS
    ) if REFERENCE_TARGET_DIR.exists() else []
    return targets[0] if len(targets) == 1 else None


def source_item_name(source) -> str:
    return source.name if isinstance(source, Path) else source.name


def source_item_path(source, folder: Path, index: int) -> Path:
    return source if isinstance(source, Path) else save_upload(source, folder, f"source_{index + 1}")


def target_item_name(target) -> str:
    return target.name if isinstance(target, Path) else target.name


def target_item_bytes(target) -> bytes:
    return target.read_bytes() if isinstance(target, Path) else target.getvalue()


def target_item_path(target, folder: Path) -> Path:
    return target if isinstance(target, Path) else save_upload(target, folder, "target")


def save_upload(uploaded_file, folder: Path, filename_prefix: str) -> Path:
    suffix = Path(uploaded_file.name).suffix or ".xlsx"
    safe_name = f"{filename_prefix}{suffix}"
    path = folder / safe_name
    path.write_bytes(uploaded_file.getvalue())
    return path


def changes_to_dataframe(changes: list[PreviewChange]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Selected": change.selected,
                "Employee name": change.employee_name,
                "Target row": change.target_row,
                "Column": change.column_letter,
                "Column heading": change.column_heading,
                "Existing target value": change.existing_value or "",
                "Proposed value": change.proposed_value or "",
                "Source file": change.source_filename,
                "Source label": change.source_label,
                "Source row": change.source_row,
                "Status found": change.status_found,
                "Change ID": change.change_id,
            }
            for change in changes
        ]
    )


def dataframe_to_changes(df: pd.DataFrame, original_changes: list[PreviewChange]) -> list[PreviewChange]:
    by_id = {change.change_id: change for change in original_changes}
    selected_changes: list[PreviewChange] = []
    for _, row in df.iterrows():
        change_id = str(row["Change ID"])
        if change_id not in by_id:
            continue
        change = by_id[change_id]
        selected = bool(row.get("Selected", False))
        proposed_raw = row.get("Proposed value", "")
        proposed_value = None if pd.isna(proposed_raw) else str(proposed_raw)
        selected_changes.append(replace(change, selected=selected, proposed_value=proposed_value))
    return selected_changes


def consolidated_report_bytes(report_rows) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Consolidated Report"
    headers = [
        "Employee name", "DNC", "STB", "WSIB", "LD", "Column where status was found",
        "Original cell value", "Source filename", "Source worksheet", "Source row",
        "Oldest occurrence", "Most recent occurrence", "Proposed final value",
        "Match result in the eighth workbook",
    ]
    ws.append(headers)
    for row in report_rows:
        ws.append([
            row.employee_name,
            row.dnc,
            row.stb,
            row.wsib,
            row.ld,
            row.column_found,
            row.original_cell_value,
            row.source_filename,
            row.source_worksheet,
            row.source_row,
            row.oldest_occurrence,
            row.most_recent_occurrence,
            row.proposed_final_value,
            row.match_result,
        ])
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def reset_results():
    for key in ["preview_changes", "unmatched", "report_rows", "updated_bytes", "updated_name", "processed"]:
        st.session_state.pop(key, None)


st.title("DEXTR — Crew Not On Duty Workbook Updater")
st.caption("Upload seven source workbooks and one target .xlsm workbook. DEXTR fills empty DNC/STB/WSIB/LD cells in a safe edited copy.")

st.markdown(
    """
<div class="warning-box">
<strong>Safety rules built in:</strong> DEXTR reads only <code>Crew Not On Duty Report</code>, ignores <code>Source</code>,
fills only empty cells in columns <strong>B through F</strong>, and never overwrites existing target data or columns <strong>A, G, or H</strong>.
The original target file is never overwritten.
</div>
""",
    unsafe_allow_html=True,
)

with st.expander("Step-by-step workflow", expanded=True):
    st.markdown(
        "<div class='step-box'>" + "".join(
            f"<p><strong>{index}.</strong> {step}</p>" for index, step in enumerate(WORKFLOW_STEPS, start=1)
        ) + "</div>",
        unsafe_allow_html=True,
    )

with st.sidebar:
    st.header("Settings")
    allow_blank_clear = False
    st.caption("DEXTR only fills empty target cells. Existing target data is never cleared or replaced.")
    st.markdown("Statuses searched case-insensitively:")
    st.code(", ".join(STATUS_CODES))
    st.markdown("Columns copied exactly in-place:")
    st.code("B→B, C→C, D→D, E→E, F→F")

st.subheader("1. Choose the seven source workbooks")
st.write("DEXTR can use the saved source set in the repo, or you can upload all seven source files to replace it for this run.")

saved_sources = reference_source_files()
use_saved_sources = len(saved_sources) == 7 and st.checkbox(
    "Use saved source files from the repo",
    value=True,
    on_change=reset_results,
)

source_uploads = []
source_orders = []
if use_saved_sources:
    source_uploads = saved_sources
    source_orders = list(range(1, 8))
    st.success("Loaded 7 saved source files from reference_source_pdfs.")
else:
    if saved_sources:
        st.info(f"Saved source set has {len(saved_sources)} file(s). Upload all seven files below to use a replacement set.")
    else:
        st.info("No complete saved source set was found. Upload all seven source files below.")
    for index in range(7):
        col_file, col_order = st.columns([4, 1])
        with col_file:
            upload = st.file_uploader(
                f"Source {index + 1}",
                type=["xlsx", "xlsm"],
                key=f"source_upload_{index}",
                on_change=reset_results,
            )
        with col_order:
            order = st.number_input(
                "Order",
                min_value=1,
                max_value=7,
                value=index + 1,
                step=1,
                key=f"source_order_{index}",
                label_visibility="collapsed",
                on_change=reset_results,
            )
        source_uploads.append(upload)
        source_orders.append(order)

st.subheader("3. Choose the eighth workbook — target")
saved_target = reference_target_file()
use_saved_target = saved_target is not None and st.checkbox(
    "Use saved eighth target workbook from the repo",
    value=True,
    on_change=reset_results,
)

if use_saved_target:
    target_workbook = saved_target
    st.success("Loaded saved target workbook from reference_target_workbook.")
    st.write(f"Target: {target_item_name(target_workbook)}")
else:
    if saved_target is None:
        st.info("No single saved target workbook was found. Upload the eighth workbook below.")
    else:
        st.info("Upload a replacement eighth target workbook for this run.")
    target_workbook = st.file_uploader(
        "Target workbook (.xlsm preferred)",
        type=["xlsm", "xlsx"],
        key="target_upload",
        on_change=reset_results,
    )

valid_sources = all(upload is not None for upload in source_uploads)
valid_order = sorted(source_orders) == list(range(1, 8))
valid_target = target_workbook is not None

if not valid_order:
    st.error("Each source must have a unique order from 1 to 7.")

ordered_sources = [upload for _, upload in sorted(zip(source_orders, source_uploads), key=lambda item: item[0])]
ordered_labels = [f"Source {i + 1}" for i in range(7)]
with st.expander("2. Confirm source chronology", expanded=valid_sources and valid_order):
    if valid_sources and valid_order:
        for label, source in zip(ordered_labels, ordered_sources):
            st.write(f"{label}: {source_item_name(source)}")
    else:
        st.info("Choose a complete saved source set or upload all seven source workbooks to confirm the oldest-to-newest order.")

st.subheader("4. Build preview")
process_clicked = st.button(
    "Build preview",
    type="primary",
    disabled=not (valid_sources and valid_order and valid_target),
)

if process_clicked:
    reset_results()
    try:
        with tempfile.TemporaryDirectory(prefix="dextr-webapp-") as tmpdir:
            tmp = Path(tmpdir)
            source_paths = [source_item_path(source, tmp, i) for i, source in enumerate(ordered_sources)]
            target_path = target_item_path(target_workbook, tmp)

            updates = build_updates_from_sources(
                source_paths,
                ordered_labels,
                allow_blank_source_cells_to_clear=allow_blank_clear,
            )
            changes, unmatched = build_preview_changes(target_path, updates)
            matched = set(target_name_index(target_path).keys())
            report_rows = consolidated_report_rows(updates, matched_names=matched)

            st.session_state["preview_changes"] = changes
            st.session_state["unmatched"] = unmatched
            st.session_state["report_rows"] = report_rows
            st.session_state["target_upload_bytes"] = target_item_bytes(target_workbook)
            st.session_state["target_upload_name"] = target_item_name(target_workbook)
            st.session_state["processed"] = True
    except Exception as exc:
        st.error(f"Could not process workbook files: {exc}")

if st.session_state.get("processed"):
    changes: list[PreviewChange] = st.session_state.get("preview_changes", [])
    unmatched = st.session_state.get("unmatched", [])
    report_rows = st.session_state.get("report_rows", [])

    st.subheader("5. Review changes before update")
    if changes:
        st.markdown(
            f"<div class='good-box'><strong>{len(changes)}</strong> empty target cell(s) can be filled. Review, edit proposed values if needed, and deselect anything you do not want applied.</div>",
            unsafe_allow_html=True,
        )
        preview_df = changes_to_dataframe(changes)
        edited_df = st.data_editor(
            preview_df,
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            disabled=[
                "Employee name", "Target row", "Column", "Column heading", "Existing target value",
                "Source file", "Source label", "Source row", "Status found", "Change ID",
            ],
            column_config={
                "Selected": st.column_config.CheckboxColumn("Selected"),
                "Proposed value": st.column_config.TextColumn("Proposed value", width="large"),
                "Change ID": None,
            },
            key="preview_editor",
        )

        st.subheader("6. Create updated workbook copy")
        if st.button("Create updated workbook", type="primary"):
            selected_changes = dataframe_to_changes(edited_df, changes)
            with tempfile.TemporaryDirectory(prefix="dextr-output-") as tmpdir:
                tmp = Path(tmpdir)
                target_name = st.session_state["target_upload_name"]
                target_path = tmp / target_name
                target_path.write_bytes(st.session_state["target_upload_bytes"])
                output_name = f"{Path(target_name).stem}_DEXTR_updated.xlsm"
                output_path = tmp / output_name
                apply_selected_changes(target_path, output_path, selected_changes)
                st.session_state["updated_bytes"] = output_path.read_bytes()
                st.session_state["updated_name"] = output_name
                applied_count = sum(1 for change in selected_changes if change.selected)
                st.success(f"Updated workbook created with {applied_count} selected change(s).")
    else:
        st.info("No matching target rows with proposed updates were found.")

    if st.session_state.get("updated_bytes"):
        st.subheader("7. Download updated workbook")
        st.download_button(
            "Download updated .xlsm workbook",
            data=st.session_state["updated_bytes"],
            file_name=st.session_state["updated_name"],
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
        )

    st.subheader("8. Consolidated report")
    if report_rows:
        report_df = pd.DataFrame([row.__dict__ for row in report_rows])
        simple_cols = ["employee_name", "dnc", "stb", "wsib", "ld", "column_found", "proposed_final_value", "match_result"]
        st.dataframe(report_df[simple_cols], use_container_width=True, hide_index=True)
        with st.expander("Technical details"):
            st.dataframe(report_df, use_container_width=True, hide_index=True)
        st.download_button(
            "Download consolidated report .xlsx",
            data=consolidated_report_bytes(report_rows),
            file_name="DEXTR_consolidated_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("No DNC/STB/WSIB/LD statuses were found in the seven source workbooks.")

    if unmatched:
        st.subheader("9. Employees not found in target")
        st.warning(f"{len(unmatched)} employee(s) appeared in source files but were not found in the eighth workbook. DEXTR does not add new rows.")
        st.dataframe(
            pd.DataFrame([
                {
                    "Employee name": item.employee_name,
                    "Source files": ", ".join(item.source_files),
                    "Statuses": ", ".join(item.statuses),
                }
                for item in unmatched
            ]),
            use_container_width=True,
            hide_index=True,
        )

with st.expander("Exact workbook rules implemented"):
    st.markdown(
        f"""
- Reads worksheet: `{REPORT_SHEET_NAME}`
- Ignores worksheet: `Source`
- Scans source columns B through F only for DNC, STB, WSIB, LD
- Copies source B→target B, C→C, D→D, E→E, F→F
- Never copies source A, G, H, or overwrites populated target cells in B through F
- Adds target column H header `HRs Left` and formulas `=60-G[row]` for QCTO, CTO, CSA, and Trainee rows
- Sorts rows inside the QCTO, CTO, CSA, and Trainee target tables by `HRs Left`, smallest to largest
- Colors rows grey only when their predicted hours are `0.00` or columns B through F contain `DNC`; `LVM@` time values and Easter, Douglas QCTO are not greyed out
- Preserves target workbook formatting by editing a copied workbook with OpenPyXL and `keep_vba=True`
- Removes target rows when column A contains STO, REO, or GSR before saving the downloadable workbook
- Does not add, delete, or reorder rows outside the QCTO, CTO, CSA, and Trainee target tables
- Blank source cells never clear target cells
"""
    )
