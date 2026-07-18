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

1. Upload seven source workbooks.
2. Reorder them if needed. Source 1 is oldest and Source 7 is newest.
3. Upload the eighth workbook as the target workbook.
4. Click `Build preview`.
5. Review proposed updates.
6. Optionally edit proposed values or deselect individual changes.
7. Download the updated `.xlsm` workbook and consolidated report.
