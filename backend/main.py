import json
import uuid
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import PlainTextResponse, Response
from fastapi.staticfiles import StaticFiles

from processing import detect_and_describe, generate_outputs, zip_dat_files
from xml_template import extract_template_block, read_default_template, write_default_template

app = FastAPI(title="OE Decipher Uploader", version="0.4.0")

# job_id -> {"file_bytes": bytes, "dat_files": {filename: bytes} | None}
JOBS: dict[str, dict] = {}


@app.post("/api/upload")
async def upload(
    otc_file: UploadFile = File(...),
    raw_data_file: UploadFile | None = File(None),
):
    if not otc_file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Please upload an .xlsx file")

    file_bytes = await otc_file.read()

    raw_file_bytes = None
    if raw_data_file is not None and raw_data_file.filename:
        if not raw_data_file.filename.lower().endswith((".xlsx", ".xlsm")):
            raise HTTPException(400, "קובץ הנתונים הגולמי חייב להיות מסוג xlsx")
        raw_file_bytes = await raw_data_file.read()

    try:
        description = detect_and_describe(file_bytes, raw_file_bytes)
    except Exception as exc:
        raise HTTPException(400, f"Failed to process file: {exc}") from exc

    if not description["blocks"]:
        raise HTTPException(400, "No question blocks with coded answers were detected")

    job_id = uuid.uuid4().hex
    JOBS[job_id] = {"file_bytes": file_bytes, "raw_file_bytes": raw_file_bytes, "dat_files": None}

    return {"job_id": job_id, "blocks": description["blocks"]}


@app.post("/api/generate")
async def generate(
    job_id: str = Form(...),
    mapping: str = Form(...),
    xml_template_file: UploadFile | None = File(None),
):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found or expired")

    try:
        mapping_list = json.loads(mapping)
    except json.JSONDecodeError as exc:
        raise HTTPException(400, f"Invalid mapping payload: {exc}") from exc

    template_text = None
    if xml_template_file is not None:
        template_bytes = await xml_template_file.read()
        template_text = template_bytes.decode("utf-8")

    try:
        result = generate_outputs(
            job["file_bytes"],
            mapping_list,
            template_text=template_text,
            raw_file_bytes=job.get("raw_file_bytes"),
        )
    except Exception as exc:
        raise HTTPException(400, f"Failed to generate output: {exc}") from exc

    job["dat_files"] = result["dat_files"]

    return {
        "job_id": job_id,
        "warnings": result["warnings"],
        "blocks": result["blocks"],
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


@app.get("/api/xml-template")
async def get_xml_template():
    content = read_default_template()
    return PlainTextResponse(
        content,
        headers={"Content-Disposition": "attachment; filename=xml_template.txt"},
    )


@app.post("/api/xml-template")
async def update_xml_template(template_file: UploadFile = File(...)):
    content = (await template_file.read()).decode("utf-8")
    block = extract_template_block(content)
    if not block.strip():
        raise HTTPException(400, "Template file appears to be empty")
    write_default_template(content)
    return {"status": "ok"}


STATIC_DIR = Path(__file__).parent / "static"
if STATIC_DIR.exists():
    app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
