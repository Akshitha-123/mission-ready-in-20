import json
from pathlib import Path
from copy import deepcopy

import pikepdf
from lxml import etree as ET

# Get the directory where THIS script lives
BASE_DIR = Path(__file__).resolve().parent

# ====== CONFIG (dynamic paths) ======
PDF_IN = BASE_DIR / "dd2977.pdf"
PDF_OUT = BASE_DIR / "dd2977_filled.pdf"
JSON_IN = BASE_DIR / "input_draw.json"
# ====================================

XFA_DATA_NS = "http://www.xfa.org/schema/xfa-data/1.0/"
NSMAP = {"xfa": XFA_DATA_NS}


# ============================================================
# Read input JSON
# ============================================================
def load_json():
    with JSON_IN.open("r", encoding="utf-8") as f:
        return json.load(f)


# ============================================================
# Extract datasets node from /AcroForm/XFA
# ============================================================
def find_xfa_datasets(pdf: pikepdf.Pdf):
    acroform = pdf.Root.get("/AcroForm", None)
    if acroform is None:
        raise RuntimeError("No /AcroForm found")

    xfa = acroform.get("/XFA", None)
    if xfa is None:
        raise RuntimeError("No /XFA array found")

    # ["template", stream, "datasets", stream, ...]
    for i in range(0, len(xfa), 2):
        if str(xfa[i]).lower() == "datasets":
            return i + 1, xfa[i + 1].read_bytes()

    raise RuntimeError("datasets stream not found")


# ============================================================
# Modify datasets XML *without changing structure*
# ============================================================
def rebuild_datasets_in_place(xml_root, data):
    # ---------- Extract major nodes ----------
    data_node = xml_root.find("xfa:data", NSMAP)
    form1 = data_node.find("form1")
    page1 = form1.find("Page1")

    # These nodes ALREADY exist in correct order. We only change text.
    one   = page1.find("One")
    two   = page1.find("Two")
    A = page1.find("A"); B = page1.find("B"); C = page1.find("C")
    D = page1.find("D"); E = page1.find("E"); F = page1.find("F")
    G = page1.find("G"); H = page1.find("H")

    fivesteps = page1.find("FiveSteps")
    part4 = page1.find("Part4thru9")
    ten = page1.find("Ten")
    eleven = page1.find("Eleven")
    twelve = page1.find("Twelve")

    # ---------- Fill Page1 simple fields ----------
    if one is not None:
        one.text = data.get("mission_task_and_description", "")

    if two is not None:
        two.text = data.get("date", "").replace("-", "")

    prep = data.get("prepared_by", {}) or {}

    if A is not None: A.text = prep.get("name_last_first_middle_initial", "")
    if B is not None: B.text = prep.get("rank_grade", "")
    if C is not None: C.text = prep.get("duty_title_position", "")
    if D is not None: D.text = prep.get("unit", "")
    if E is not None: E.text = prep.get("work_email", "")
    if F is not None: F.text = prep.get("telephone", "")
    if G is not None: G.text = prep.get("uic_cin", "")
    if H is not None: H.text = prep.get("training_support_or_lesson_plan_or_opord", "")

    # ---------- Eleven - Overall Supervision Plan ----------
    if eleven is not None:
        eleven.text = data.get("overall_supervision_plan", "")

    # ---------- Ten - Overall RRL ----------
    if ten is not None:
        overall = (data.get("overall_residual_risk_level") or "").upper()

        for name in ["EHigh", "High", "Med", "Low"]:
            node = ten.find(name)
            if node is not None:
                node.text = "0"

        mapping = {
            "EH": "EHigh",
            "H": "High",
            "M": "Med",
            "L": "Low",
        }
        if overall in mapping:
            tgt = ten.find(mapping[overall])
            if tgt is not None:
                tgt.text = "1"

    # ---------- Twelve - Approval ----------
    if twelve is not None:
        appr = data.get("approval_or_disapproval_of_mission_or_task", {}) or {}
        approve_node = twelve.find("Approve")
        dis_node = twelve.find("Disapprove")

        if approve_node is not None:
            approve_node.text = "1" if appr.get("approve") else "0"
        if dis_node is not None:
            dis_node.text = "1" if appr.get("disapprove") else "0"

    # ---------- Part4thru9 - Hazards Table ----------
    if part4 is not None:
        # Locate template row (existing Row1)
        template_row = part4.find("Row1")

        # Remove ALL existing rows
        for child in list(part4):
            if child.tag == "Row1":
                part4.remove(child)

        # Create new rows based on template
        for st in data.get("subtasks", []):
            row = deepcopy(template_row)

            # Subtask-Substep
            sub = row.find("Subtask-Substep")
            if sub is not None:
                sub.text = (st.get("subtask") or {}).get("name", "")

            # Hazard
            haz = row.find("Hazard")
            if haz is not None:
                haz.text = st.get("hazard", "")

            # Initial Risk Level
            irl = row.find("InitialRiskLevel")
            if irl is not None:
                irl.text = (st.get("initial_risk_level") or "").upper()

            # Control
            ctrl = row.find("Control")
            if ctrl is not None:
                ctrl.text = "\n".join((st.get("control") or {}).get("values", []))

            # HOW / WHO
            table2 = row.find("Table2")
            if table2 is not None:
                r1 = table2.find("Row1"); tf1 = r1.find("TextField1")
                r2 = table2.find("Row2"); tf2 = r2.find("TextField2")

                how_vals = (st.get("how_to_implement") or {}).get("how", {}).get("values", [])
                who_vals = (st.get("how_to_implement") or {}).get("who", {}).get("values", [])

                if tf1 is not None:
                    tf1.text = "\n".join(how_vals)
                if tf2 is not None:
                    tf2.text = "\n".join(who_vals)

            # RRL
            rrl = row.find("RRL")
            if rrl is not None:
                rrl.text = (st.get("residual_risk_level") or "").upper()

            part4.append(row)


# ============================================================
# MAIN
# ============================================================
def main():
    data = load_json()

    pdf = pikepdf.Pdf.open(PDF_IN)
    xfa_index, datasets_bytes = find_xfa_datasets(pdf)

    # Parse XFA datasets XML
    xml_root = ET.fromstring(datasets_bytes)

    # Update only allowed fields while keeping order
    rebuild_datasets_in_place(xml_root, data)

    # Convert back to bytes
    new_xml = ET.tostring(xml_root, encoding="utf-8", xml_declaration=False)

    # ---- SAFE UPDATE: write new XML without replacing stream object ----
    datasets_stream = pdf.Root["/AcroForm"]["/XFA"][xfa_index]
    datasets_stream.write(new_xml)

    pdf.save(PDF_OUT)
    print("✅ SUCCESS – Editable DD2977 saved:", PDF_OUT)


if __name__ == "__main__":
    main()
