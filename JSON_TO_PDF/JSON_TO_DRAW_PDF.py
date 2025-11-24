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
from copy import deepcopy
from pathlib import Path

from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject, BooleanObject, ArrayObject


# ============================================================
# CONFIG
# ============================================================

TEMPLATE_PATH = Path("DD-Form-2977.pdf")
INPUT_JSON = Path("input_draw.json")
OUTPUT_PDF = Path("generated_DRAW.pdf")

MAX_ROWS = 19  # DD 2977 supports sub_1 through sub_19


# ============================================================
# HELPERS
# ============================================================

def normalize_text(s):
    """Fix bad unicode characters & None -> empty string."""
    if s is None:
        return ""
    s = str(s)

    # Replace weird unicode spaces that show up as □ blocks
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

def fill_pdf(template_path: Path, output_path: Path, field_values: dict) -> None:
    """
    Copy template, keep all AcroForm fields editable, and set values.
    """

    reader = PdfReader(str(template_path))
    writer = PdfWriter()

    # Copy pages
    for page in reader.pages:
        writer.add_page(page)

    # Copy AcroForm (with proper PdfObject types)
    if "/AcroForm" in reader.trailer["/Root"]:
        src_acro = reader.trailer["/Root"]["/AcroForm"]
        new_acro = deepcopy(src_acro)

        # Fields array must be an ArrayObject for PyPDF2
        if "/Fields" in new_acro:
            copied = []
            for f in new_acro["/Fields"]:
                copied.append(deepcopy(f))
            new_acro[NameObject("/Fields")] = ArrayObject(copied)

        # Ask viewer to regenerate appearances (important)
        new_acro[NameObject("/NeedAppearances")] = BooleanObject(True)

        writer._root_object.update({NameObject("/AcroForm"): new_acro})

    # Fill fields (text, combos, buttons)
    for page in writer.pages:
        writer.update_page_form_field_values(page, field_values)

    # Explicitly sync /V and /AS for button fields (checkboxes & radios)
    try:
        acro = writer._root_object["/AcroForm"]
        for field in acro["/Fields"]:
            name = field.get("/T")
            if not name:
                continue
            key = name.strip("()")
            if key not in field_values:
                continue
            if field.get("/FT") == "/Btn":
                val = str(field_values[key])
                field.update({
                    NameObject("/V"): NameObject(f"/{val}") if not val.startswith("/") else NameObject(val),
                    NameObject("/AS"): NameObject(f"/{val}") if not val.startswith("/") else NameObject(val),
                })
    except Exception:
        # If anything weird happens here, we still keep text fields working
        pass

    # Write output PDF
    with output_path.open("wb") as f:
        writer.write(f)


# ============================================================
# MAIN
# ============================================================

def main():
    if not TEMPLATE_PATH.exists():
        raise SystemExit(f"Template not found: {TEMPLATE_PATH}")

    if not INPUT_JSON.exists():
        raise SystemExit(f"Input JSON not found: {INPUT_JSON}")

    with INPUT_JSON.open("r", encoding="utf-8") as f:
        draw_data = json.load(f)

    field_values = build_field_values(draw_data)
    fill_pdf(TEMPLATE_PATH, OUTPUT_PDF, field_values)

    print(f"Created: {OUTPUT_PDF}")


if __name__ == "__main__":
    main()
