# DEXTR

DEXTR is a Streamlit web app for updating a Crew Not On Duty target workbook from seven source workbooks.

It follows these safety rules:

- Reads the `Crew Not On Duty Report` worksheet.
- Ignores the `Source` worksheet.
- Searches source columns B through F for `DNC`, `STB`, `WSIB`, and `LD`.
- Preserves same-column placement: B to B, C to C, D to D, E to E, F to F.
- Never copies or overwrites columns A, G, or H.
- Uses OpenPyXL with `keep_vba=True` for `.xlsm` target workbooks.
- Edits a copy of the target workbook, never the original upload.
- Preserves the target workbook's formatting, formulas, sheet names, and workbook structure.
- Removes rows from the downloadable target copy when column A contains `STO`, `REO`, or `GSR`, case-insensitively.
- Adds `HRs Left` in column H with `=60-G[row]` formulas for QCTO, CTO, CSA, and Trainee rows in the downloadable target copy.
- Sorts rows inside the QCTO, CTO, CSA, and Trainee target tables by `HRs Left`, smallest to largest.
- Colors rows grey only when their predicted hours are `0.00` or columns B through F contain `DNC`; `LVM@` time values and Easter, Douglas QCTO are not greyed out.
- Lets the user preview, edit, and deselect individual changes before creating the updated workbook.
- Creates a downloadable consolidated report.

## Run locally

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Test

```bash
pytest -q
```

## Workbook flow

1. Use the saved seven source workbooks or upload a replacement source set.
2. Confirm or adjust the source order from oldest to newest.
3. Use the saved eighth target workbook or upload a replacement target workbook.
4. Click `Build preview`.
5. Review proposed updates.
6. Optionally edit proposed values or deselect individual changes.
7. Click `Create updated workbook`.
8. Download the updated `.xlsm` workbook.
9. Download the consolidated report.
10. Check any unmatched employees before using the updated file.
