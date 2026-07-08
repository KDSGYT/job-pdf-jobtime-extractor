from __future__ import annotations

import html
import io
import json
from pathlib import Path

from starlette.applications import Starlette
from starlette.datastructures import UploadFile
from starlette.requests import Request
from starlette.responses import HTMLResponse, JSONResponse, PlainTextResponse, Response
from starlette.routing import Route
from starlette.staticfiles import StaticFiles

from parser import parse_pdf, rows_to_csv, rows_to_dicts

BASE_DIR = Path(__file__).resolve().parent


def render_index() -> str:
    return (BASE_DIR / "static" / "index.html").read_text()


async def homepage(request: Request) -> HTMLResponse:
    return HTMLResponse(render_index())


async def extract(request: Request) -> JSONResponse:
    form = await request.form()
    upload = form.get("pdf")
    if not isinstance(upload, UploadFile):
        return JSONResponse({"error": "Please upload a PDF file using the field named 'pdf'."}, status_code=400)
    if not upload.filename.lower().endswith(".pdf"):
        return JSONResponse({"error": "Only PDF files are supported."}, status_code=400)

    content = await upload.read()
    if not content:
        return JSONResponse({"error": "The uploaded PDF was empty."}, status_code=400)

    try:
        rows = parse_pdf(io.BytesIO(content))
    except Exception as exc:
        return JSONResponse({"error": f"Could not read this PDF: {exc}"}, status_code=400)

    data = rows_to_dicts(rows)
    counts = {
        "records": len(data),
        "jobs": len({row["job_number"] for row in data}),
        "pd_time_records": sum(1 for row in data if row["record_type"] == "PD time"),
        "stby_records": sum(1 for row in data if row["record_type"] == "STBY"),
        "split_records": sum(1 for row in data if row["record_type"] == "Split"),
    }
    return JSONResponse({"filename": upload.filename, "counts": counts, "rows": data})


async def extract_csv(request: Request) -> Response:
    form = await request.form()
    upload = form.get("pdf")
    if not isinstance(upload, UploadFile):
        return PlainTextResponse("Please upload a PDF file using the field named 'pdf'.", status_code=400)
    content = await upload.read()
    try:
        rows = parse_pdf(io.BytesIO(content))
    except Exception as exc:
        return PlainTextResponse(f"Could not read this PDF: {exc}", status_code=400)
    csv_text = rows_to_csv(rows)
    safe_name = Path(upload.filename).stem.replace(" ", "_") or "job_times"
    return Response(
        csv_text,
        media_type="text/csv",
        headers={"Content-Disposition": f'attachment; filename="{safe_name}_pd_stby_split.csv"'},
    )


routes = [
    Route("/", homepage),
    Route("/api/extract", extract, methods=["POST"]),
    Route("/api/extract.csv", extract_csv, methods=["POST"]),
]

app = Starlette(debug=True, routes=routes)
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
