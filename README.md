# Job PDF PD/STBY/Split Extractor

Streamlit web app for uploading crew/job description PDFs and extracting:

- Job number
- Valid days
- Job start/end/duration
- On-duty location
- PD time / PDTIME / PDTIM rows
- STBY rows
- Split start/end/duration
- CSV export

## Run locally

```bash
cd job-pdf-jobtime-extractor
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/streamlit run streamlit_app.py --server.port 8020
```

Open on this Mac:

```text
http://127.0.0.1:8020/
```

## Deploy to Streamlit Community Cloud

Repository:

```text
https://github.com/KDSGYT/job-pdf-jobtime-extractor
```

Deployment settings:

- Repository: `KDSGYT/job-pdf-jobtime-extractor`
- Branch: `main`
- Main file path: `streamlit_app.py`
- Python dependencies: `requirements.txt`

Prefilled deploy URL:

```text
https://share.streamlit.io/deploy?repository=KDSGYT%2Fjob-pdf-jobtime-extractor&branch=main&mainModule=streamlit_app.py
```

## Test parser with the sample PDF

```bash
.venv/bin/python tests/smoke_test.py
```

## Optional API/static web app

The repository also includes a Starlette API app with a static HTML/JS frontend:

```bash
.venv/bin/uvicorn app:app --reload --port 8020
```
