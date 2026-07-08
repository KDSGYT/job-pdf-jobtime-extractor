# Job PDF PD/STBY/Split Extractor

Local web app for uploading crew/job description PDFs and extracting:

- Job number
- Valid days
- Job start/end/duration
- On-duty location
- PD time / PDTIME / PDTIM rows
- STBY rows
- Split start/end/duration
- CSV export

## Run

```bash
cd /Users/alfred/job-pdf-jobtime-extractor
python3 -m uvicorn app:app --host 0.0.0.0 --port 8020
```

Open on this Mac:

```text
http://127.0.0.1:8020/
```

Open from another device on the same Wi-Fi/LAN, replace the IP if needed:

```text
http://10.0.0.167:8020/
```

## Test with the sample PDF

```bash
python3 tests/smoke_test.py
```
