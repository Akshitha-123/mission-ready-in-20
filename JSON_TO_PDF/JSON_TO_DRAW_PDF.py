import json
import re
from pathlib import Path
from copy import deepcopy

import pikepdf
from lxml import etree as ET

# ============================================================
# Paths: Script auto-reads files from the same folder
# ============================================================
BASE_DIR = Path(__file__).resolve().parent

PDF_IN = BASE_DIR / "dd2977.pdf"
PDF_OUT = BASE_DIR / "dd2977_filled.pdf"
JSON_IN = BASE_DIR / "input_draw.json"

# ============================================================
# XFA namespace
# ============================================================
XFA_DATA_NS = "http://www.xfa.org/schema/xfa-data/1.0/"
NSMAP = {"xfa": XFA_DATA_NS}

# ============================================================
# Utility: Clean weird unicode characters (fixes □ issues)
# ============================================================
def clean_ascii(s: str):
    """
    Remove zero-width spaces, Adobe special NBSP, en-space, em-space, 
    and any unicode outside the standard printable ASCII range.
    """
    if s is None:
        return ""
    return re.sub(r"[^\x20-\x7E]", "", s)

# ============================================================
# Read JSON input
# ============================================================
def load_json():
    with JSON_IN.open("r", encoding="utf-8") as f:
        return json.load(f)

# ============================================================
# Locate datasets XML inside /AcroForm/XFA array
# ============================================================
def find_xfa_datasets(pdf: pikepdf.Pdf):
    acroform = pdf.Root.get("/AcroForm", None)
    if acroform is None:
        raise RuntimeError("No /AcroForm in PDF")

    xfa = acroform.get("/XFA", None)
    if xfa is None:
        raise RuntimeError("No /XFA array found")

    for i in range(0, len(xfa), 2):
        if str(xfa[i]).lower() == "datasets":
            return i + 1, xfa[i + 1].read_bytes()

    raise RuntimeError("datasets XFA packet not found")

# ============================================================
# CORE: Update only the text values, never restructure XML
# ============================================================
def rebuild_datasets_in_place(xml_root, data):

    # 🔹 Navigate to main nodes (all already exist)
    data_node = xml_root.find("xfa:data", NSMAP)
    form1 = data_node.find("form1")
    page1 = form1.find("Page1")

    # Simple fields
    one = page1.find("One")
    two = page1.find("Two")
    A = page1.find("A"); B = page1.find("B"); C = page1.find("C")
    D = page1.find("D"); E = page1.find("E"); F = page1.find("F")
    G = page1.find("G"); H = page1.find("H")

    eleven = page1.find("Eleven")
    ten = page1.find("Ten")
    twelve = page1.find("Twelve")
    part4 = page1.find("Part4thru9")

    # ============================================================
    # Fill Page1 fields
    # ============================================================
    if one is not None:
        one.text = clean_ascii(data.get("mission_task_and_description", ""))

    raw_date = data.get("date", "")
    clean_date = clean_ascii(raw_date).replace("-", "")
    if two is not None:
        two.text = clean_ascii(clean_date)

    prep = data.get("prepared_by", {}) or {}
    if A is not None: A.text = clean_ascii(prep.get("name_last_first_middle_initial", ""))
    if B is not None: B.text = clean_ascii(prep.get("rank_grade", ""))
    if C is not None: C.text = clean_ascii(prep.get("duty_title_position", ""))
    if D is not None: D.text = clean_ascii(prep.get("unit", ""))
    if E is not None: E.text = clean_ascii(prep.get("work_email", ""))
    if F is not None: F.text = clean_ascii(prep.get("telephone", ""))
    if G is not None: G.text = clean_ascii(prep.get("uic_cin", ""))
    if H is not None: H.text = clean_ascii(prep.get("training_support_or_lesson_plan_or_opord", ""))

    if eleven is not None:
        eleven.text = clean_ascii(data.get("overall_supervision_plan", ""))

    # ============================================================
    # Block 10 — Overall RRL
    # ============================================================
    if ten is not None:
        overall = clean_ascii((data.get("overall_residual_risk_level") or "").upper())

        for tag in ["EHigh", "High", "Med", "Low"]:
            node = ten.find(tag)
            if node is not None:
                node.text = "0"

        map_rrl = {"EH": "EHigh", "H": "High", "M": "Med", "L": "Low"}
        if overall in map_rrl:
            tgt = ten.find(map_rrl[overall])
            if tgt is not None:
                tgt.text = "1"

    # ============================================================
    # Block 12 — Approval / Disapproval
    # ============================================================
    if twelve is not None:
        appr = data.get("approval_or_disapproval_of_mission_or_task", {}) or {}

        approve = twelve.find("Approve")
        dis = twelve.find("Disapprove")

        if approve is not None:
            approve.text = "1" if appr.get("approve") else "0"

        if dis is not None:
            dis.text = "1" if appr.get("disapprove") else "0"

    # ============================================================
    # Hazard Table (Part4thru9)
    # ============================================================
    if part4 is not None:
        template_row = part4.find("Row1")

        # Remove existing rows
        for child in list(part4):
            if child.tag == "Row1":
                part4.remove(child)

        # Rebuild rows
        for st in data.get("subtasks", []):
            row = deepcopy(template_row)

            # Subtask
            sub = row.find("Subtask-Substep")
            if sub is not None:
                sub.text = clean_ascii((st.get("subtask") or {}).get("name", ""))

            # Hazard
            haz = row.find("Hazard")
            if haz is not None:
                haz.text = clean_ascii(st.get("hazard", ""))

            # Initial Risk Level
            irl = row.find("InitialRiskLevel")
            if irl is not None:
                irl.text = clean_ascii((st.get("initial_risk_level") or "").upper())

            # Control
            ctrl = row.find("Control")
            if ctrl is not None:
                ctrl.text = clean_ascii("\n".join((st.get("control") or {}).get("values", [])))

            # HOW / WHO
            table2 = row.find("Table2")
            if table2 is not None:
                r1 = table2.find("Row1"); tf1 = r1.find("TextField1")
                r2 = table2.find("Row2"); tf2 = r2.find("TextField2")

                how_vals = (st.get("how_to_implement") or {}).get("how", {}).get("values", [])
                who_vals = (st.get("how_to_implement") or {}).get("who", {}).get("values", [])

                if tf1 is not None:
                    tf1.text = clean_ascii("\n".join(how_vals))
                if tf2 is not None:
                    tf2.text = clean_ascii("\n".join(who_vals))

            # Residual Risk Level
            rrl = row.find("RRL")
            if rrl is not None:
                rrl.text = clean_ascii((st.get("residual_risk_level") or "").upper())

            part4.append(row)

# ============================================================
# MAIN
# ============================================================
def main():
    data = load_json()

    pdf = pikepdf.Pdf.open(PDF_IN)
    xfa_index, datasets_bytes = find_xfa_datasets(pdf)

    xml_root = ET.fromstring(datasets_bytes)
    rebuild_datasets_in_place(xml_root, data)

    new_xml = ET.tostring(xml_root, encoding="utf-8", xml_declaration=False)

    datasets_stream = pdf.Root["/AcroForm"]["/XFA"][xfa_index]
    datasets_stream.write(new_xml)

    pdf.save(PDF_OUT)
    print("✅ SUCCESS — Filled (editable) DD2977 created at:")
    print(PDF_OUT)


if __name__ == "__main__":
    main()
