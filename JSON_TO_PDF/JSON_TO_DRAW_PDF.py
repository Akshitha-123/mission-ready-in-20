#!/usr/bin/env python3
"""
JSON_TO_DRAW_PDF.py

Generates a fully editable DD Form 2977 (Deliberate Risk Assessment Worksheet)
from input_draw.json, using DD-Form-2977.pdf as the template.

This version:
- Keeps ALL form fields editable (no flattening)
- Fixes unicode "block" characters in text
- Copies AcroForm + widgets correctly
- Ensures Block 10 & 12 radio / checkbox groups stay usable
"""

import json
from pathlib import Path

from PyPDF2 import PdfReader, PdfWriter  # type: ignore[import]
from PyPDF2.generic import NameObject, BooleanObject  # type: ignore[import]
from PyPDF2.errors import DependencyError  # type: ignore[attr-defined]

try:  # PyMuPDF is needed only for preview rendering
    import fitz  # type: ignore[import]
except ImportError:  # pragma: no cover - optional dependency
    fitz = None  # type: ignore[assignment]


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "DD-Form-2977.pdf"
INPUT_JSON = BASE_DIR / "input_draw.json"
OUTPUT_PDF = BASE_DIR / "generated_DRAW.pdf"

MAX_ROWS = 19  # DD 2977 supports sub_1 through sub_19


# ============================================================
# HELPERS
# ============================================================

def normalize_text(s):
    """Fix bad unicode characters & None -> empty string."""
    if s is None:
        return ""
    s = str(s)

    # Replace weird unicode spaces that show up as â–¡ blocks
    bad_spaces = {0x00A0, 0x202F, 0x2009, 0x2011}
    return "".join(" " if ord(c) in bad_spaces else c for c in s)


def build_field_values(draw_data: dict) -> dict:
    """
    Convert JSON into a dict of PDF field_name -> value.

    Field names match those actually present in DD-Form-2977.pdf:
    mission, prep_name, sub_1, haz_1, control_1, how_1, who_1, init_risk1, ...
    plus:
      - overall_res  (Block 10 radio group, export: EH/H/M/L)
      - xapp         (Block 12 approve/disapprove radio group, export: app/dis)
    """

    fields = {}

    # --------------------------- Header ---------------------------
    fields["mission"] = normalize_text(draw_data.get("mission_task_and_description", ""))
    fields["date"] = normalize_text(draw_data.get("date", ""))

    prepared = draw_data.get("prepared_by", {}) or {}
    fields["prep_name"]  = normalize_text(prepared.get("name_last_first_middle_initial", ""))
    fields["prep_rank"]  = normalize_text(prepared.get("rank_grade", ""))
    fields["prep_title"] = normalize_text(prepared.get("duty_title_position", ""))
    fields["prep_unit"]  = normalize_text(prepared.get("unit", ""))
    fields["prep_email"] = normalize_text(prepared.get("work_email", ""))
    fields["prep_phone"] = normalize_text(prepared.get("telephone", ""))
    fields["prep_uic"]   = normalize_text(prepared.get("uic_cin", ""))
    fields["prep_plan"]  = normalize_text(prepared.get("training_support_or_lesson_plan_or_opord", ""))

    # ---------------------- Overall Supervision ----------------------
    fields["overall_plan"] = normalize_text(draw_data.get("overall_supervision_plan", ""))

    # -------------------------- Subtasks ----------------------------
    subtasks = draw_data.get("subtasks", []) or []

    # Map risk levels to the PDF combo export values
    # (per form: ['0',' '], ['1','EH'], ['2','H'], ['3','M'], ['4','L'])
    risk_map = {"EH": "1", "E": "1", "H": "2", "M": "3", "L": "4"}

    for idx, item in enumerate(subtasks, 1):
        if idx > MAX_ROWS:
            break

        # 4. SUBTASK
        sub_name = ""
        subobj = item.get("subtask")
        if isinstance(subobj, dict):
            sub_name = subobj.get("name", "")
        fields[f"sub_{idx}"] = normalize_text(sub_name)

        # 5. HAZARD
        fields[f"haz_{idx}"] = normalize_text(item.get("hazard", ""))

        # 7. CONTROL
        ctrl = item.get("control")
        if isinstance(ctrl, dict):
            vals = ctrl.get("values", []) or []
            text = "".join(f"- {normalize_text(v)}\n" for v in vals).rstrip()
            fields[f"control_{idx}"] = text

        # 8. HOW / WHO
        how_to = item.get("how_to_implement", {}) or {}

        how_vals = []
        hv = how_to.get("how")
        if isinstance(hv, dict):
            how_vals = hv.get("values", []) or []

        who_vals = []
        wv = how_to.get("who")
        if isinstance(wv, dict):
            who_vals = wv.get("values", []) or []

        if how_vals:
            fields[f"how_{idx}"] = "\n".join(normalize_text(v) for v in how_vals)
        if who_vals:
            fields[f"who_{idx}"] = "\n".join(normalize_text(v) for v in who_vals)

        # 6. INITIAL RISK
        init_level = normalize_text(item.get("initial_risk_level", "")).upper()
        if init_level:
            fields[f"init_risk{idx}"] = risk_map.get(init_level, "0")

        # 9. RESIDUAL RISK
        res_level = normalize_text(item.get("residual_risk_level", "")).upper()
        if res_level:
            fields[f"res_risk{idx}"] = risk_map.get(res_level, "0")

    # ------------------ Block 10 & 12 radio groups ------------------

    # Block 10: overall_res (export names /EH, /H, /M, /L)
    overall_json = normalize_text(
        draw_data.get("overall_residual_risk_level", "")
    ).upper()
    overall_map = {"EH": "EH", "E": "EH", "H": "H", "M": "M", "L": "L"}

    if overall_json:
        fields["overall_res"] = overall_map.get(overall_json, "L")
    else:
        # Default matches blank form: LOW selected
        fields["overall_res"] = "L"

    # Block 12: xapp (export names /app, /dis)
    approve_json = normalize_text(draw_data.get("approval_decision", "")).lower()
    if approve_json in {"approve", "app", "approved"}:
        fields["xapp"] = "app"
    elif approve_json in {"disapprove", "dis", "disapproved"}:
        fields["xapp"] = "dis"
    else:
        # Default matches template: DISAPPROVE (you can change in Adobe)
        fields["xapp"] = "dis"

    return fields


# ============================================================
# ACROFORM + WIDGET COPY
# ============================================================

def fill_pdf(template_path: Path, output_path: Path, field_values: dict):
    reader = PdfReader(str(template_path))
    writer = PdfWriter()

    # Copy pages
    for p in reader.pages:
        writer.add_page(p)

    # Copy AcroForm
    root = writer._root_object
    if "/AcroForm" in reader.trailer["/Root"]:
        acro = reader.trailer["/Root"]["/AcroForm"]
        acro_new = acro.clone(writer)

        # Remove XFA
        if "/XFA" in acro_new:
            del acro_new["/XFA"]

        acro_new[NameObject("/NeedAppearances")] = BooleanObject(True)
        root[NameObject("/AcroForm")] = acro_new

    # Helper to deref
    def deref(obj):
        if hasattr(obj, "idnum"):
            return reader.get_object(obj)
        return obj

    # Helper: find all widget annotations for a field
    def find_widgets(field_name):
        widgets = []
        for page in writer.pages:
            annots = page.get("/Annots")
            if not annots:
                continue

            # Normalize: annots can be array OR a single indirect object
            if hasattr(annots, "idnum"):
                annots = [annots]  # make it iterable

            for a in annots:
                obj = deref(a)
                if obj.get("/T") == field_name:
                    widgets.append(obj)

        return widgets

    fields = reader.get_fields()
    text_field_values: dict[str, str] = {}

    for name, field in fields.items():
        if name not in field_values:
            continue

        val = field_values[name]
        ft = field.get("/FT")

        # --- TEXT / DROPDOWN ---
        if ft in ("/Tx", "/Ch"):
            text_field_values[name] = val
            continue

        # --- CHECKBOX / RADIO ---
        if ft == "/Btn":
            export = val if isinstance(val, str) and val.startswith("/") else f"/{val}"
            export = NameObject(export)

            widgets = find_widgets(name)
            if not widgets:
                continue  # should never happen

            try:
                parent_obj = deref(field.get("obj"))
                if parent_obj:
                    parent_obj.update({NameObject("/V"): export})
            except Exception:
                pass

            for widget in widgets:
                ap = widget.get("/AP")
                if not ap or "/N" not in ap:
                    continue

                appearances = ap["/N"]

                if export in appearances:
                    widget.update({NameObject("/AS"): export})
                else:
                    widget.update({NameObject("/AS"): NameObject("/Off")})

    if text_field_values:
        for page in writer.pages:
            writer.update_page_form_field_values(page, text_field_values)

    # Save
    with open(output_path, "wb") as f:
        writer.write(f)


# ============================================================
# FORM APPEARANCE REFRESH
# ============================================================

def refresh_form_appearances(pdf_path: Path | str) -> None:
    """Regenerate widget appearances so non-Acrobat viewers show filled text."""

    if fitz is None:
        return

    pdf_path = Path(pdf_path)
    with fitz.open(pdf_path) as doc:  # type: ignore[arg-type]
        changed = False
        for page in doc:
            widgets = list(page.widgets() or [])
            if not widgets:
                continue
            for widget in widgets:
                value = widget.field_value
                if value in (None, ""):
                    continue
                widget.field_value = value
                widget.update()
                changed = True

        if changed:
            encrypt_opt = getattr(fitz, "PDF_ENCRYPT_KEEP", 0)
            doc.save(pdf_path, deflate=True, incremental=True, encryption=encrypt_opt)

# ============================================================
# PREVIEW RENDERING
# ============================================================

def render_preview_pdf(editable_pdf: Path | str, preview_path: Path | str, zoom: float = 1.5) -> Path:
    """Render a flattened preview PDF from the editable DRAW.

    The result keeps the exact page dimensions but bakes the filled form fields
    into the page graphics so browser viewers (PDF.js) display the text.
    """

    if fitz is None:
        raise RuntimeError("PyMuPDF is required to render DRAW previews. Install via 'pip install pymupdf'.")

    editable_pdf = Path(editable_pdf)
    preview_path = Path(preview_path)

    matrix = fitz.Matrix(zoom, zoom)

    with fitz.open(editable_pdf) as source_doc, fitz.open() as preview_doc:  # type: ignore[arg-type]
        for page in source_doc:
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            new_page = preview_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(page.rect, pixmap=pix)

        preview_doc.save(preview_path, deflate=True)

    return preview_path


# ============================================================
# PUBLIC API + MAIN
# ============================================================

def generate_draw_pdf(draw_data: dict, output_path: Path | str, template_path: Path | str | None = None) -> Path:
    """Render the provided DRAW JSON into a DD-2977 PDF and return the path."""

    template = Path(template_path) if template_path else TEMPLATE_PATH
    output_path = Path(output_path)

    if not template.exists():
        raise FileNotFoundError(f"Template not found: {template}")

    field_values = build_field_values(draw_data)
    try:
        fill_pdf(template, output_path, field_values)
        refresh_form_appearances(output_path)
    except DependencyError as exc:
        raise RuntimeError(
            "PDF rendering requires PyCryptodome. Install it via 'pip install pycryptodome'."
        ) from exc

    return output_path


def main():
    if not INPUT_JSON.exists():
        raise SystemExit(f"Input JSON not found: {INPUT_JSON}")

    with INPUT_JSON.open("r", encoding="utf-8") as f:
        draw_data = json.load(f)

    output = generate_draw_pdf(draw_data, OUTPUT_PDF)
    print(f"Created: {output}")


if __name__ == "__main__":
    main()