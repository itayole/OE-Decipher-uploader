import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import Response
from fastapi.staticfiles import StaticFiles

from processing import process_workbook, zip_dat_files

app = FastAPI(title="OE Decipher Uploader", version="0.1.0")

# job_id -> {"blocks": [...], "dat_files": {filename: bytes}}
JOBS: dict[str, dict] = {}


@app.post("/api/process")
async def process(otc_file: UploadFile = File(...)):
    if not otc_file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Please upload an .xlsx file")

    file_bytes = await otc_file.read()
    try:
        result = process_workbook(file_bytes)
    except Exception as exc:
        raise HTTPException(400, f"Failed to process file: {exc}") from exc

    if not result["blocks"]:
        raise HTTPException(400, "No question blocks with coded answers were detected")

    job_id = uuid.uuid4().hex
    JOBS[job_id] = result

    return {
        "job_id": job_id,
        "blocks": [
            {k: v for k, v in b.items() if k != "dat_files"} for b in result["blocks"]
        ],
    }


@app.get("/api/download/{job_id}")
async def download(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found or expired")
    zip_bytes = zip_dat_files(job["dat_files"])
    return Response(
        content=zip_bytes,
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=coded_dat_files.zip"},
    )


STATIC_DIR = Path(__file__).parent / "static"
if STATIC_DIR.exists():
    app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
