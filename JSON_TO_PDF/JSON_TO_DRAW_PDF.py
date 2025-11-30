import json
import re
import shutil
from pathlib import Path
from copy import deepcopy
from typing import Any, Dict, Union

import pikepdf
from lxml import etree as ET

# ============================================================
# Paths
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
# Utility: Clean text
# ============================================================
def clean_ascii(s: str):
    """
    Remove zero-width spaces, Adobe special NBSP, en-space, em-space, 
    and any unicode outside the standard printable ASCII range.
    """
    if s is None:
        return ""
    return re.sub(r"[^\x20-\x7E\n\r\t]", "", str(s))

# ============================================================
# XFA Logic
# ============================================================
def find_xfa_datasets(pdf: pikepdf.Pdf):
    acroform = pdf.Root.get("/AcroForm", None)
    if acroform is None:
        raise RuntimeError("No /AcroForm in PDF")

    xfa = acroform.get("/XFA", None)
    if xfa is None:
        raise RuntimeError("No /XFA array found")

    # XFA is an array of [key, stream, key, stream, ...]
    for i in range(0, len(xfa), 2):
        if str(xfa[i]) == "datasets":
            return i + 1, xfa[i + 1].read_bytes()

    raise RuntimeError("datasets XFA packet not found")

def rebuild_datasets_in_place(xml_root, data):
    # 🔹 Navigate to main nodes (all already exist in the template)
    data_node = xml_root.find("xfa:data", NSMAP)
    if data_node is None:
        # Fallback if namespace prefix is missing or different
        data_node = xml_root.find("{http://www.xfa.org/schema/xfa-data/1.0/}data")
    
    if data_node is None:
        # Try finding without namespace if strictly necessary, but XFA usually has it.
        # Let's print root children if we fail?
        pass

    form1 = data_node.find("form1")
    page1 = form1.find("Page1")

    # Simple fields mapping based on the original file content
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

        # Reset all
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
        
        if template_row is not None:
            # Remove existing rows (keep template in memory)
            # Note: In XFA, repeated elements are usually siblings.
            # We need to be careful not to remove the only Row1 if we need it for cloning.
            # But here we clone it first.
            
            # Find all Row1 elements and remove them
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
                    r1 = table2.find("Row1"); tf1 = r1.find("TextField1") if r1 is not None else None
                    r2 = table2.find("Row2"); tf2 = r2.find("TextField2") if r2 is not None else None

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
# Main Exported Function
# ============================================================
def generate_draw_pdf(data: Dict[str, Any], output_path: Union[str, Path]):
    """
    Generates a filled DD2977 PDF based on the provided data using XFA injection.
    """
    output_path = Path(output_path)
    
    if not PDF_IN.exists():
        raise FileNotFoundError(f"Template PDF not found at {PDF_IN}")

    pdf = pikepdf.Pdf.open(PDF_IN)
    try:
        xfa_index, datasets_bytes = find_xfa_datasets(pdf)

        xml_root = ET.fromstring(datasets_bytes)
        rebuild_datasets_in_place(xml_root, data)

        new_xml = ET.tostring(xml_root, encoding="utf-8", xml_declaration=False)

        # Update the XFA stream
        datasets_stream = pdf.Root["/AcroForm"]["/XFA"][xfa_index]
        
        # pikepdf stream update
        datasets_stream.write(new_xml)

        pdf.save(output_path)
        print(f"✅ SUCCESS — Filled (XFA) DD2977 created at: {output_path}")
    finally:
        pdf.close()

def render_preview_pdf(input_path: Union[str, Path], output_path: Union[str, Path]):
    """
    Creates a preview version of the PDF. 
    For now, this simply copies the file.
    """
    input_path = Path(input_path)
    output_path = Path(output_path)
    shutil.copy2(input_path, output_path)

# ============================================================
# CLI Entrypoint
# ============================================================
def main():
    def load_json():
        with JSON_IN.open("r", encoding="utf-8") as f:
            return json.load(f)
            
    data = load_json()
    generate_draw_pdf(data, PDF_OUT)

if __name__ == "__main__":
    main()
