import uuid
from pathlib import Path

from fastapi import Body, FastAPI, File, HTTPException, UploadFile
from fastapi.responses import Response
from fastapi.staticfiles import StaticFiles

from processing import detect_and_describe, generate_outputs, zip_dat_files

app = FastAPI(title="OE Decipher Uploader", version="0.2.0")

# job_id -> {"file_bytes": bytes, "dat_files": {filename: bytes} | None}
JOBS: dict[str, dict] = {}


@app.post("/api/upload")
async def upload(otc_file: UploadFile = File(...)):
    if not otc_file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Please upload an .xlsx file")

    file_bytes = await otc_file.read()
    try:
        description = detect_and_describe(file_bytes)
    except Exception as exc:
        raise HTTPException(400, f"Failed to process file: {exc}") from exc

    if not description["blocks"]:
        raise HTTPException(400, "No question blocks with coded answers were detected")

    job_id = uuid.uuid4().hex
    JOBS[job_id] = {"file_bytes": file_bytes, "dat_files": None}

    return {"job_id": job_id, "blocks": description["blocks"]}


@app.post("/api/generate")
async def generate(payload: dict = Body(...)):
    job_id = payload.get("job_id")
    mapping = payload.get("mapping", [])
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found or expired")

    try:
        result = generate_outputs(job["file_bytes"], mapping)
    except Exception as exc:
        raise HTTPException(400, f"Failed to generate output: {exc}") from exc

    job["dat_files"] = result["dat_files"]

    return {
        "job_id": job_id,
        "warnings": result["warnings"],
        "blocks": [{k: v for k, v in b.items()} for b in result["blocks"]],
    }


@app.get("/api/download/{job_id}")
async def download(job_id: str):
    job = JOBS.get(job_id)
    if not job or not job.get("dat_files"):
        raise HTTPException(404, "Job not found or not yet generated")
    zip_bytes = zip_dat_files(job["dat_files"])
    return Response(
        content=zip_bytes,
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=coded_dat_files.zip"},
    )


STATIC_DIR = Path(__file__).parent / "static"
if STATIC_DIR.exists():
    app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
