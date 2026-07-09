from __future__ import annotations

import hashlib
import io
from pathlib import Path

import streamlit as st

from parser import parse_pdf, rows_to_csv, rows_to_dicts
from table_state import filter_rows, sort_rows


@st.cache_data(show_spinner=False)
def extract_pdf(pdf_bytes: bytes) -> tuple[list, list[dict], str]:
    """Parse each unique PDF once; searches/sorts reuse these rows."""
    rows = parse_pdf(io.BytesIO(pdf_bytes))
    return rows, rows_to_dicts(rows), rows_to_csv(rows)


st.set_page_config(
    page_title="Job PDF PD/STBY/Split Extractor",
    page_icon="📄",
    layout="wide",
)

st.markdown(
    """
    <style>
    .main-title {font-size: 2.6rem; font-weight: 800; margin-bottom: .25rem;}
    .muted {color: #64748b; font-size: 1.05rem;}
    .metric-card {border: 1px solid rgba(148, 163, 184, .25); border-radius: 16px; padding: 1rem;}
    div[data-testid="stMetricValue"] {font-size: 1.8rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="main-title">📄 Job PDF PD/STBY/Split Extractor</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="muted">Upload a crew/job description PDF and extract job numbers, PD time, STBY, and split start/end/duration.</div>',
    unsafe_allow_html=True,
)
st.divider()

uploaded_file = st.file_uploader("Upload job description PDF", type=["pdf"])

if not uploaded_file:
    st.info("Choose a PDF file to begin. The output can be searched, filtered, sorted, and downloaded as CSV.")
    st.stop()

pdf_bytes = uploaded_file.getvalue()
pdf_hash = hashlib.sha256(pdf_bytes).hexdigest()
is_new_pdf = pdf_hash != st.session_state.get("pdf_hash")

status = st.spinner("Reading PDF and extracting PD time, STBY, and split records...") if is_new_pdf else st.container()
with status:
    try:
        rows, data, csv_text = extract_pdf(pdf_bytes)
        st.session_state["pdf_hash"] = pdf_hash
        st.session_state["rows"] = rows
        st.session_state["data"] = data
        st.session_state["csv_text"] = csv_text
    except Exception as exc:
        st.error(f"Could not read this PDF: {exc}")
        st.stop()

data = st.session_state["data"]
csv_text = st.session_state["csv_text"]

if not data:
    st.warning("No PD time, STBY, or non-zero split records were found in this PDF.")
    st.stop()

job_count = len({row["job_number"] for row in data})
pd_count = sum(1 for row in data if row["record_type"] == "PD time")
stby_count = sum(1 for row in data if row["record_type"] == "STBY")
split_count = sum(1 for row in data if row["record_type"] == "Split")

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Jobs", job_count)
c2.metric("Total records", len(data))
c3.metric("PD time", pd_count)
c4.metric("STBY", stby_count)
c5.metric("Split", split_count)

base_name = Path(uploaded_file.name).stem.replace(" ", "_") or "job_times"
st.download_button(
    "⬇️ Download CSV",
    data=csv_text,
    file_name=f"{base_name}_pd_stby_split.csv",
    mime="text/csv",
    use_container_width=True,
)

st.subheader("Extracted records")
filter_col_1, filter_col_2, sort_col_1, sort_col_2 = st.columns([2, 1, 1, 1])
search = filter_col_1.text_input("Search", placeholder="Job number, location, valid days...")
type_options = ["All", "PD time", "STBY", "Split"]
record_type = filter_col_2.selectbox("Record type", type_options)
sort_by = sort_col_1.selectbox(
    "Sort by",
    [
        "Click table header",
        "duration",
        "start",
        "end",
        "job_start",
        "job_end",
        "job_duration",
        "on_duty_location",
        "job_number",
        "record_type",
    ],
)
sort_desc = sort_col_2.toggle("Descending", value=False)

filtered = filter_rows(data, search, record_type)
if sort_by != "Click table header":
    filtered = sort_rows(filtered, sort_by, sort_desc)

st.caption(f"Showing {len(filtered)} of {len(data)} records")
st.caption("Tip: click any table column header to sort in-place. Search and sorting reuse the uploaded PDF's extracted data; the PDF is only reparsed when the file changes.")
st.dataframe(
    filtered,
    use_container_width=True,
    hide_index=True,
    column_order=[
        "job_number",
        "valid_days",
        "record_type",
        "start",
        "end",
        "duration",
        "from_location",
        "to_location",
        "job_start",
        "job_end",
        "job_duration",
        "on_duty_location",
        "page",
        "raw_line",
    ],
    column_config={
        "job_number": "Job #",
        "valid_days": "Valid days",
        "record_type": "Type",
        "start": "Start",
        "end": "End",
        "duration": "Duration",
        "from_location": "From",
        "to_location": "To",
        "job_start": "Job start",
        "job_end": "Job end",
        "job_duration": "Job duration",
        "on_duty_location": "On duty",
        "page": "PDF page",
        "raw_line": "Raw line",
    },
)

with st.expander("What the app extracts"):
    st.markdown(
        """
        - **PD time**: rows labelled `PDTIME`, `PDTIM`, or `PD`.
        - **STBY**: rows labelled `STBY`.
        - **Split**: non-zero split time plus `Split from ... to ...` when present.
        - Each row includes the parent **job number**, valid days, job start/end, page number, start/end, and duration.
        """
    )
