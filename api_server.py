"""FastAPI server for CONOPS uploads and parsing."""
from __future__ import annotations

import logging
import shutil
import subprocess
import uuid
from pathlib import Path
from typing import Any, Dict, Optional

from fastapi import FastAPI, File, HTTPException, UploadFile, Request
from fastapi.concurrency import run_in_threadpool
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import re


from JSON_TO_PDF.JSON_TO_DRAW_PDF import generate_draw_pdf, render_preview_pdf

try:
    from generate_draw import generate_draw_for_conop
    GENERATE_DRAW_IMPORT_ERROR: Optional[Exception] = None
except Exception as exc:  # pragma: no cover - optional dependency
    generate_draw_for_conop = None  # type: ignore[assignment]
    GENERATE_DRAW_IMPORT_ERROR = exc

from parse_conop import extract_text_from_pptx, parse_conop_sections

logger = logging.getLogger(__name__)

app = FastAPI(title="Mission Ready In 20 API", version="0.1.0")

if GENERATE_DRAW_IMPORT_ERROR:
    logger.warning("DRAW generation disabled: %s", GENERATE_DRAW_IMPORT_ERROR)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_ROOT = Path("/app/uploaded_conops")
UPLOAD_ROOT.mkdir(parents=True, exist_ok=True)
UPLOAD_ROOT_RESOLVED = UPLOAD_ROOT.resolve()

DRAW_OUTPUT_ROOT = Path("/app/generated_draws")
DRAW_OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

app.mount("/uploads", StaticFiles(directory=str(UPLOAD_ROOT)), name="uploads")
app.mount("/generated_draws", StaticFiles(directory=str(DRAW_OUTPUT_ROOT)), name="generated_draws")


class PreviewConversionError(RuntimeError):
    """Raised when a PPTX cannot be converted into a PDF preview."""


def _find_libreoffice() -> str | None:
    """Return the path to the LibreOffice/soffice executable if available."""
    for candidate in ("soffice", "libreoffice"):
        located = shutil.which(candidate)
        if located:
            return located
    return None


def convert_pptx_to_pdf(ppt_path: Path) -> Path:
    """Convert a PPTX file to PDF using LibreOffice and return the PDF path."""
    soffice_cmd = _find_libreoffice()
    if soffice_cmd is None:
        raise PreviewConversionError(
            "LibreOffice (soffice) is not installed or not on PATH. Install it via 'brew install --cask libreoffice'."
        )

    output_dir = ppt_path.parent
    pdf_path = output_dir / f"{ppt_path.stem}.pdf"

    try:
        completed = subprocess.run(
            [
                soffice_cmd,
                "--headless",
                "--convert-to",
                "pdf:impress_pdf_Export",
                "--outdir",
                str(output_dir),
                str(ppt_path),
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        if completed.stderr:
            logger.debug("LibreOffice stderr: %s", completed.stderr.decode("utf-8", errors="ignore").strip())
    except (subprocess.CalledProcessError, FileNotFoundError) as exc:
        stderr = ""
        if isinstance(exc, subprocess.CalledProcessError) and exc.stderr:
            stderr = exc.stderr.decode("utf-8", errors="ignore").strip()
        logger.error("LibreOffice conversion failed: %s", stderr or exc)
        raise PreviewConversionError("LibreOffice failed to convert the PPTX to PDF.") from exc

    if not pdf_path.exists():
        raise PreviewConversionError("LibreOffice did not produce a PDF output.")

    return pdf_path


@app.post("/api/conops/upload")
async def upload_conop(file: UploadFile = File(...)) -> dict[str, object]:
    """Receive a CONOPS PPTX, persist it, and return parsed content."""
    filename = file.filename or "uploaded-conop.pptx"
    suffix = Path(filename).suffix.lower()
    if suffix != ".pptx":
        raise HTTPException(status_code=400, detail="Only .pptx files are supported.")

    unique_name = f"{Path(filename).stem}-{uuid.uuid4().hex}{suffix}"
    saved_path = UPLOAD_ROOT / unique_name

    try:
        with saved_path.open("wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
    finally:
        file.file.close()

    text_content = extract_text_from_pptx(saved_path)
    if text_content is None:
        saved_path.unlink(missing_ok=True)
        raise HTTPException(status_code=422, detail="Unable to extract text from the uploaded PPTX.")

    sections = parse_conop_sections(text_content)

    preview_url: str | None = None
    try:
        pdf_path = convert_pptx_to_pdf(saved_path)
    except PreviewConversionError as exc:
        logger.warning("Skipping PDF preview for %s: %s", saved_path.name, exc)
    else:
        preview_url = f"/uploads/{pdf_path.name}"

    return {
        "filename": filename,
        "stored_path": str(saved_path.resolve()),
        "raw_text": text_content,
        "sections": sections,
        "preview_url": preview_url,
    }


class ConvertPreviewRequest(BaseModel):
    stored_path: str


@app.post("/api/conops/convert-preview")
def convert_preview(request: ConvertPreviewRequest) -> dict[str, str]:
    """Convert an already-uploaded CONOPS PPTX into PDF for previewing."""

    ppt_path = Path(request.stored_path).expanduser().resolve()
    if not ppt_path.exists() or not ppt_path.is_file():
        raise HTTPException(status_code=404, detail="Stored CONOPS file was not found.")

    try:
        ppt_path.relative_to(UPLOAD_ROOT_RESOLVED)
    except ValueError:
        raise HTTPException(status_code=400, detail="Invalid CONOPS path for conversion.")

    existing_pdf = ppt_path.with_suffix(".pdf")
    if existing_pdf.exists():
        return {"preview_url": f"/uploads/{existing_pdf.name}"}

    try:
        pdf_path = convert_pptx_to_pdf(ppt_path)
    except PreviewConversionError as exc:
        raise HTTPException(status_code=503, detail=str(exc))

    return {"preview_url": f"/uploads/{pdf_path.name}"}


class GenerateDrawRequest(BaseModel):
    filename: str
    raw_text: str
    sections: Dict[str, str]


@app.post("/api/conops/generate-draw")
async def generate_draw_endpoint(payload: GenerateDrawRequest, request: Request) -> dict[str, Any]:
    """Generate a DRAW from previously parsed CONOPS content."""

    if generate_draw_for_conop is None:
        detail = "DRAW generator unavailable on this server."
        if GENERATE_DRAW_IMPORT_ERROR is not None:
            detail += f" Reason: {GENERATE_DRAW_IMPORT_ERROR}"
        raise HTTPException(status_code=503, detail=detail)

    draw_payload: Optional[Dict[str, Any]] = None
    draw_error: Optional[str] = None
    draw_pdf_url: Optional[str] = None
    draw_pdf_preview_url: Optional[str] = None

    request_data = {
        "conops": {
            "filename": payload.filename,
            "raw_text": payload.raw_text,
            "sections": payload.sections,
        }
    }

    try:
        draw_payload = await run_in_threadpool(
            generate_draw_for_conop,
            request_data,
            None,
        )
    except Exception as exc:  # pragma: no cover - relies on external services
        logger.error("DRAW generation failed for %s", payload.filename, exc_info=exc)
        draw_error = str(exc)

    if draw_payload:
        original_stem = Path(payload.filename).stem
        safe_stem = re.sub(r"[^A-Za-z0-9_-]+", "_", original_stem)
        pdf_filename = f"{safe_stem}-draw-{uuid.uuid4().hex}.pdf"
        preview_filename = f"{safe_stem}-draw-{uuid.uuid4().hex}-preview.pdf"

        pdf_path = DRAW_OUTPUT_ROOT / pdf_filename
        preview_filename = pdf_filename.replace(".pdf", "-preview.pdf")
        preview_path = DRAW_OUTPUT_ROOT / preview_filename

        try:
            await run_in_threadpool(generate_draw_pdf, draw_payload, pdf_path)
            try:
                draw_pdf_url = str(request.url_for("generated_draws", path=pdf_filename))
            except Exception:
                backend_base = str(request.base_url).rstrip("/")
                draw_pdf_url = f"{backend_base}/generated_draws/{pdf_filename}"


            try:
                await run_in_threadpool(render_preview_pdf, pdf_path, preview_path, draw_payload)
            except Exception as exc:
                logger.warning("DRAW preview rendering failed for %s: %s", payload.filename, exc)
            else:
                try:
                    draw_pdf_preview_url = str(request.url_for("generated_draws", path=preview_filename))
                except Exception:
                    draw_pdf_preview_url = f"{backend_base}/generated_draws/{preview_filename}"

        except Exception as exc:  # pragma: no cover - depends on local PDF tooling
            logger.error("DRAW PDF rendering failed for %s", payload.filename, exc_info=exc)
            if draw_error is None:
                draw_error = f"DRAW PDF could not be rendered: {exc}"

    return {
        "draw": draw_payload,
        "draw_error": draw_error,
        "draw_pdf_url": draw_pdf_url,
        "draw_pdf_preview_url": draw_pdf_preview_url,
    }
