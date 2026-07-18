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
</style>
"""
st.markdown(APP_CSS, unsafe_allow_html=True)


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
st.caption("Upload seven source workbooks and one target .xlsm workbook. DEXTR previews DNC/STB/WSIB/LD updates before creating a safe edited copy.")

st.markdown(
    """
<div class="warning-box">
<strong>Safety rules built in:</strong> DEXTR reads only <code>Crew Not On Duty Report</code>, ignores <code>Source</code>,
updates only columns <strong>B through F</strong>, and never overwrites columns <strong>A, G, or H</strong>.
The original target file is never overwritten.
</div>
""",
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Settings")
    allow_blank_clear = st.checkbox(
        "Allow blank source cells to clear target cells",
        value=False,
        help="Disabled by default. Leave off unless you intentionally want newer blanks to clear older/target information.",
    )
    st.markdown("Statuses searched case-insensitively:")
    st.code(", ".join(STATUS_CODES))
    st.markdown("Columns copied exactly in-place:")
    st.code("B→B, C→C, D→D, E→E, F→F")

st.subheader("1. Upload the seven source workbooks")
st.write("Source 1 is treated as oldest. Source 7 is treated as newest. You can change the order before processing.")

source_uploads = []
source_orders = []
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

st.subheader("2. Upload the eighth workbook — target")
target_upload = st.file_uploader(
    "Target workbook (.xlsm preferred)",
    type=["xlsm", "xlsx"],
    key="target_upload",
    on_change=reset_results,
)

valid_sources = all(upload is not None for upload in source_uploads)
valid_order = sorted(source_orders) == list(range(1, 8))
valid_target = target_upload is not None

if not valid_order:
    st.error("Each source must have a unique order from 1 to 7.")

ordered_sources = [upload for _, upload in sorted(zip(source_orders, source_uploads), key=lambda item: item[0])]
ordered_labels = [f"Source {i + 1}" for i in range(7)]
if valid_sources and valid_order:
    with st.expander("Current source chronology", expanded=True):
        for label, upload in zip(ordered_labels, ordered_sources):
            st.write(f"{label}: {upload.name}")

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
            source_paths = [save_upload(upload, tmp, f"source_{i + 1}") for i, upload in enumerate(ordered_sources)]
            target_path = save_upload(target_upload, tmp, "target")

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
            st.session_state["target_upload_bytes"] = target_upload.getvalue()
            st.session_state["target_upload_name"] = target_upload.name
            st.session_state["processed"] = True
    except Exception as exc:
        st.error(f"Could not process workbook files: {exc}")

if st.session_state.get("processed"):
    changes: list[PreviewChange] = st.session_state.get("preview_changes", [])
    unmatched = st.session_state.get("unmatched", [])
    report_rows = st.session_state.get("report_rows", [])

    st.subheader("3. Preview changes before update")
    if changes:
        st.markdown(
            f"<div class='good-box'><strong>{len(changes)}</strong> proposed cell update(s) found. Review, edit proposed values if needed, and deselect anything you do not want applied.</div>",
            unsafe_allow_html=True,
        )
        preview_df = changes_to_dataframe(changes)
        edited_df = st.data_editor(
            preview_df,
            hide_index=True,
            use_container_width=True,
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
        st.download_button(
            "Download updated .xlsm workbook",
            data=st.session_state["updated_bytes"],
            file_name=st.session_state["updated_name"],
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
        )

    st.subheader("4. Consolidated report")
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
        st.subheader("Employees not found in target")
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
- Never copies or overwrites A, G, or H
- Preserves target workbook formatting by editing a copied workbook with OpenPyXL and `keep_vba=True`
- Does not add, delete, or reorder employee rows
- Blank source cells do not clear target cells unless the sidebar setting is enabled
"""
    )
