#!/usr/bin/env python3
"""
Enhanced DD2977 DRAW Parser

Key improvements:
1. Template-based JSON structure for efficiency - all forms use the same base structure
2. Checkbox fields return 1 if checked/marked, 0 if unchecked
3. Modular helper functions for better code organization
4. Consistent field initialization to prevent undefined values
5. Enhanced risk calculation logic
"""
import argparse, json, re, sys
from pathlib import Path
from datetime import datetime

try:  # Optional dependencies for XFA extraction
    import pikepdf  # type: ignore
    from lxml import etree  # type: ignore
except Exception:  # pragma: no cover - optional feature
    pikepdf = None
    etree = None

COMMON_HAZARD_PREFIXES = {
    "ACCIDENT", "AIRCRAFT", "AIR", "AMMUNITION", "APPROACH", "BAD", "BIOLOGICAL",
    "BURN", "CHEMICAL", "COLD", "CONTAMINATION", "CARGO", "DAMAGE", "DEHYDRATION",
    "DETONATION", "DISCHARGE", "DROWNING", "ELECTRICAL", "EMERGENCY", "EXPLOSION",
    "FAIL", "FAILURE", "FALL", "FALLS", "FATIGUE", "FIRE", "FRATRICIDE", "HAZARD",
    "HEAT", "HELICOPTER", "HOT", "ILLNESS", "IMPROPER", "INJURY", "INJURIES",
    "INSUFFICIENT", "INTERNAL", "LIGHTNING", "LOSS", "LOST", "LOW", "MECHANICAL",
    "MISFIRE", "MISSING", "MIS", "NEGLECT", "NEGATIVE", "NEGLIGENT", "POOR",
    "RADIATION", "RESIDUAL", "RISK", "ROAD", "ROLL", "ROLLOVER", "SLING", "SPILL",
    "SURFACE", "TRIP", "TRIPS", "UXO", "VEHICLE", "VEHICLES", "WEATHER", "WILDLIFE",
    "WRONG", "LACK", "INSUFFICIENT", "INAPPROPRIATE", "SLICK", "ICE", "SNOW",
    "RAIN", "WIND", "MECHANICAL", "STRUCTURAL", "FUEL", "SMOKE", "FOD", "FROST",
    "EXPOSURE", "HYPOTHERMIA", "HYPERTHERMIA", "CASUALTY", "CASUALTIES", "OVERHEAT",
    "OVEREXERTION", "CONGESTION", "COLLISION", "IMPACT", "CONTACT", "HAZMAT",
    "INCLEMENT", "DAMAGED", "ENVIRONMENTAL", "SURVIVABILITY", "VEGETATION", "HAZE",
    "POISON", "TOXIC", "FUMES", "OBSTACLE", "RUNWAY", "LANDING", "PZ", "LZ",
    "MARKING", "WEAK", "UNSECURED", "UNSTABLE", "CONTAGION"
}

HAZARD_STOPWORDS = {"THE", "AND", "OF", "IN", "ON", "AT", "TO", "A", "AN", "WITH", "BY", "FOR"}

# ----------------------------
# Text extraction: PyMuPDF -> pdfminer -> OCR (optional)
# ----------------------------
def extract_text_pymupdf(pdf_path: Path) -> str:
    try:
        import fitz
    except Exception:
        return ""
    try:
        out = []
        with fitz.open(pdf_path) as doc:
            for p in doc:
                out.append(p.get_text("text") or "")
        return "\n".join(out)
    except Exception:
        return ""

def extract_text_pdfminer(pdf_path: Path) -> str:
    try:
        from pdfminer.high_level import extract_text
    except Exception:
        return ""
    try:
        return extract_text(str(pdf_path)) or ""
    except Exception:
        return ""

def extract_text_ocr(pdf_path: Path) -> str:
    try:
        import pytesseract
        from pdf2image import convert_from_path
    except Exception:
        return ""
    try:
        pages = convert_from_path(str(pdf_path))
    except Exception:
        return ""
    return "\n".join(pytesseract.image_to_string(img) for img in pages)

def normalize_text(t: str) -> str:
    t = t.replace("\r", "\n")
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t

def extract_text_multibackend(pdf_path: Path, force_ocr: bool=False) -> str:
    if force_ocr:
        return normalize_text(extract_text_ocr(pdf_path))
    for fn in (extract_text_pymupdf, extract_text_pdfminer, extract_text_ocr):
        txt = fn(pdf_path)
        if txt and txt.strip():
            return normalize_text(txt)
    return ""

# ----------------------------
# Parsing helpers tuned to your DD2977 text layer
# ----------------------------
def pick(pattern, text, flags=re.I):
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else None

def pick_all(pattern, text, flags=re.I|re.DOTALL):
    vals = re.findall(pattern, text, flags)
    return [v.strip() for v in vals]

def clean_items(s: str) -> list:
    """Split multi-line controls into bullet-ish items."""
    if not s:
        return []
    # break on line breaks or leading dashes
    parts = re.split(r"(?:\n|(?:^|\s)[–\-]\s*)", s)
    parts = [re.sub(r"\s+", " ", p).strip(" -") for p in parts if p and p.strip(" -")]
    return parts

def value_after(tag_letter: str, text: str) -> str | None:
    """Extract field values that follow lettered bullets within section 3."""
    label_pattern = re.compile(rf"(?:^|\n)\s*{re.escape(tag_letter)}\.\s*", re.I)
    label_match = label_pattern.search(text)
    if not label_match:
        return None

    # Grab everything after the label start
    remainder = text[label_match.end():]
    if not remainder:
        return None

    # Determine the earliest boundary (next label or numbered section)
    boundaries = []
    newline_label = re.search(r"(?:^|\n)\s*[a-i]\.\s", remainder, flags=re.I)
    if newline_label:
        boundaries.append(newline_label.start())
    inline_label = re.search(r"\s{2,}[a-i]\.\s", remainder, flags=re.I)
    if inline_label:
        boundaries.append(inline_label.start())
    numbered_section = re.search(r"(?:^|\n)\s*\d+\.\s", remainder)
    if numbered_section:
        boundaries.append(numbered_section.start())

    cutoff = min(boundaries) if boundaries else len(remainder)
    window = remainder[:cutoff].strip()
    if not window:
        return None

    # Remove instructions that occasionally bleed into the capture
    instruction_patterns = [
        r"\(1\)\s*Identify the hazards.*?equal to numbered items on form\)",
        r"Five steps of Risk Management:.*?equal to numbered items on form\)",
    ]
    for inst_pattern in instruction_patterns:
        window = re.sub(inst_pattern, "", window, flags=re.I | re.DOTALL).strip()

    if not window:
        return None

    # Break into lines to separate descriptors from actual content
    lines = [line.strip() for line in window.splitlines() if line.strip()]
    if not lines:
        return None

    descriptor_patterns = {
        "a": r"^NAME(?:\s*\([^)]*\))?\s*[:\-]?\s*",
        "b": r"^RANK\s*/?\s*GRADE(?:\s*\([^)]*\))?\s*[:\-]?\s*",
        "c": r"^DUTY\s*TITLE\s*/?\s*POSITION(?:\s*\([^)]*\))?\s*[:\-]?\s*",
        "d": r"^UNIT(?:\s*\([^)]*\))?\s*[:\-]?\s*",
        "e": r"^WORK\s*EMAIL(?:\s*\([^)]*\))?\s*[:\-]?\s*",
        "f": r"^TELEPHONE(?:\s*\([^)]*\))?\s*[:\-]?\s*",
        "g": r"^UIC\s*/?\s*CIN(?:\s*\([^)]*\))?\s*[:\-]?\s*",
        "h": r"^TRAINING\s*SUPPORT/LESSON\s*PLAN\s*OR\s*OPORD(?:\s*\([^)]*\))?\s*[:\-]?\s*",
        "i": r"^SIGNATURE\s*OF\s*PREPARER(?:\s*\([^)]*\))?\s*[:\-]?\s*",
    }

    descriptor_pattern = descriptor_patterns.get(tag_letter.lower())
    if descriptor_pattern and lines:
        cleaned_first = re.sub(descriptor_pattern, "", lines[0], flags=re.I).strip()
        if cleaned_first:
            lines[0] = cleaned_first
        else:
            lines = lines[1:]

    if lines and lines[0]:
        # Strip dangling punctuation left behind from descriptor removal (e.g., leading parenthesis)
        lines[0] = lines[0].lstrip("):- ")

    lines = [line for line in lines if line]
    if not lines:
        return None

    value_text = " ".join(lines)
    value_text = re.sub(r"\s+", " ", value_text).strip()
    return value_text or None

# ----------------------------
# Base JSON structure template for efficiency
# ----------------------------
def get_dd2977_template() -> dict:
    """Returns the base structure for DD2977 forms with all fields initialized."""
    return {
        "mission_task_and_description": None,
        "date": None,
        "prepared_by": {
            "name_last_first_middle_initial": None,
            "rank_grade": None,
            "duty_title_position": None,
            "unit": None,
            "work_email": None,
            "telephone": None,
            "uic_cin": None,
            "training_support_or_lesson_plan_or_opord": None,
            "signature_of_preparer": None,
        },
        "subtasks": [],
        "overall_residual_risk_level": None,
        "overall_supervision_plan": None,
        "approval_or_disapproval_of_mission_or_task": {
            "approve": 0,
            "disapprove": 0
        }
    }

def get_subtask_template() -> dict:
    """Returns the base structure for a subtask row."""
    return {
        "subtask": {
            "name": None
        },
        "hazard": None,
        "initial_risk_level": None,
        "control": {"values": []},
        "how_to_implement": {
            "how": {"values": []},
            "who": {"values": []}
        },
        "residual_risk_level": None
    }


def _coerce_to_string(value) -> str | None:
    """Best-effort conversion of nested XML-derived values into a trimmed string."""
    if value is None:
        return None
    if isinstance(value, str):
        cleaned = value.strip()
        return cleaned or None
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, dict):
        for v in value.values():
            result = _coerce_to_string(v)
            if result:
                return result
        return None
    if isinstance(value, (list, tuple)):
        for item in value:
            result = _coerce_to_string(item)
            if result:
                return result
    return None


def _split_multiline(text: str | None) -> list[str]:
    """Split bullet-style content into clean individual lines."""
    if not text:
        return []
    parts = re.split(r"[\r\n]+", text)
    cleaned = []
    for part in parts:
        fragment = part.strip()
        if not fragment:
            continue
        fragment = re.sub(r"^[•*\-]+\s*", "", fragment)
        fragment = re.sub(r"\s+", " ", fragment)
        if fragment:
            cleaned.append(fragment)
    return cleaned


def _normalize_risk_level(value) -> str | None:
    """Normalize risk wording or numeric codes into canonical short forms."""
    raw = _coerce_to_string(value)
    if not raw:
        return None

    raw_upper = raw.upper()
    mapping = {
        "EXTREMELY HIGH": "EH",
        "VERY HIGH": "H",
        "HIGH": "H",
        "MEDIUM": "M",
        "MED": "M",
        "LOW": "L",
        "MODERATE": "M",
        "NEGLIGIBLE": "L",
    }
    if raw_upper in mapping:
        return mapping[raw_upper]

    numeric_map = {"0": "L", "1": "M", "2": "H", "3": "EH"}
    if raw in numeric_map:
        return numeric_map[raw]

    if raw_upper in {"EH", "H", "M", "L"}:
        return raw_upper

    return raw  # Return original string if unrecognized


def _is_marked(value) -> bool:
    """Determine whether an XFA checkbox-like field is marked."""
    raw = _coerce_to_string(value)
    if not raw:
        return False

    normalized = raw.strip().lower()
    return normalized in {"1", "x", "true", "checked", "yes", "on"}


def _xml_node_to_obj(node):
    """Recursively convert an lxml Element into Python primitives."""
    children = list(node)
    if not children:
        return (node.text or "").strip()

    bucket = {}
    for child in children:
        tag = child.tag.split("}")[-1]
        payload = _xml_node_to_obj(child)
        if tag in bucket:
            if not isinstance(bucket[tag], list):
                bucket[tag] = [bucket[tag]]
            bucket[tag].append(payload)
        else:
            bucket[tag] = payload
    return bucket


def extract_xfa_dataset_from_pdf(pdf_path: Path):
    """Return the XFA datasets payload as nested Python objects if present."""
    if not (pikepdf and etree):
        return None
    try:
        with pikepdf.open(pdf_path) as pdf:
            acro_form = pdf.Root.get("/AcroForm")
            if not acro_form:
                return None

            xfa = acro_form.get("/XFA")
            if not xfa:
                return None

            datasets_bytes = None
            if isinstance(xfa, pikepdf.Stream):
                datasets_bytes = xfa.read_bytes()
            elif isinstance(xfa, pikepdf.Array):
                for idx in range(0, len(xfa), 2):
                    key = xfa[idx]
                    stream = xfa[idx + 1]
                    if str(key) == "datasets" and isinstance(stream, pikepdf.Stream):
                        datasets_bytes = stream.read_bytes()
                        break

            if not datasets_bytes:
                return None

        root = etree.fromstring(datasets_bytes)
        data_node = root.find("xfa:data", {"xfa": "http://www.xfa.org/schema/xfa-data/1.0/"})
        if data_node is None:
            return None
        return _xml_node_to_obj(data_node)
    except Exception:
        return None


def parse_dd2977_xfa(pdf_path: Path) -> dict | None:
    """Parse DD2977 content from XFA-based PDFs into the standard JSON structure."""
    payload = extract_xfa_dataset_from_pdf(pdf_path)
    if not payload or not isinstance(payload, dict):
        return None

    form = payload.get("form1")
    if not isinstance(form, dict):
        return None

    page1 = form.get("Page1") or {}
    if not isinstance(page1, dict):
        return None

    data = get_dd2977_template()
    data["mission_task_and_description"] = _coerce_to_string(page1.get("One"))
    data["date"] = _coerce_to_string(page1.get("Two"))

    prepared_map = {
        "name_last_first_middle_initial": "A",
        "rank_grade": "B",
        "duty_title_position": "C",
        "unit": "D",
        "work_email": "E",
        "telephone": "F",
        "uic_cin": "G",
        "training_support_or_lesson_plan_or_opord": "H",
        "signature_of_preparer": "I",
    }
    for dest, key in prepared_map.items():
        data["prepared_by"][dest] = _coerce_to_string(page1.get(key))

    part_section = page1.get("Part4thru9") or {}
    rows_payload = []
    if isinstance(part_section, dict):
        for key, value in part_section.items():
            if not isinstance(key, str) or not key.lower().startswith("row"):
                continue
            if isinstance(value, list):
                rows_payload.extend(value)
            else:
                rows_payload.append(value)

    subtasks = []
    last_subtask_name = None
    for entry in rows_payload:
        if not isinstance(entry, dict):
            continue

        row = get_subtask_template()

        subtask_name = _coerce_to_string(
            entry.get("Subtask-Substep") or entry.get("Subtask_Substep") or entry.get("Subtask")
        )
        hazard_text = _coerce_to_string(entry.get("Hazard"))
        initial_risk = _normalize_risk_level(entry.get("InitialRiskLevel"))
        residual_risk_value = None
        for key in ("RRL", "ResidualRiskLevel", "ResidualRiskLvl", "ResidualRiskLevel1", "ResidualRiskLevel_1"):
            candidate = entry.get(key)
            if candidate not in (None, ""):
                residual_risk_value = candidate
                break

        residual_risk = _normalize_risk_level(residual_risk_value)

        control_text = _coerce_to_string(entry.get("Control"))
        control_values = _split_multiline(control_text)

        table2 = entry.get("Table2") or {}
        if isinstance(table2, list):
            table2 = table2[0] if table2 else {}
        if not isinstance(table2, dict):
            table2 = {}

        how_text = _coerce_to_string(table2.get("Row1"))
        who_text = _coerce_to_string(table2.get("Row2"))

        if subtask_name:
            last_subtask_name = subtask_name
        elif last_subtask_name:
            subtask_name = last_subtask_name

        if hazard_text:
            hazard_text = re.sub(r"\s+", " ", hazard_text).strip()

        row["subtask"]["name"] = subtask_name
        row["hazard"] = hazard_text
        row["initial_risk_level"] = initial_risk
        row["control"]["values"] = control_values
        row["how_to_implement"]["how"]["values"] = [re.sub(r"\s+", " ", how_text)] if how_text else []
        row["how_to_implement"]["who"]["values"] = [re.sub(r"\s+", " ", who_text)] if who_text else []
        row["residual_risk_level"] = residual_risk

        subtasks.append(row)

    data["subtasks"] = subtasks

    ten_block = page1.get("Ten")
    if isinstance(ten_block, dict):
        for key, level in [("EHigh", "EH"), ("High", "H"), ("Med", "M"), ("Low", "L")]:
            if _is_marked(ten_block.get(key)):
                data["overall_residual_risk_level"] = level
                break

    if not data["overall_residual_risk_level"]:
        residual_risks = [row["residual_risk_level"] for row in subtasks]
        data["overall_residual_risk_level"] = calculate_overall_risk(residual_risks)

    data["overall_supervision_plan"] = _coerce_to_string(page1.get("Eleven"))

    approval_block = page1.get("Twelve") or {}
    if isinstance(approval_block, dict):
        data["approval_or_disapproval_of_mission_or_task"]["approve"] = 1 if _is_marked(approval_block.get("Approve")) else 0
        data["approval_or_disapproval_of_mission_or_task"]["disapprove"] = 1 if _is_marked(approval_block.get("Disapprove")) else 0

    return data

def parse_checkbox_value(text: str, pattern: str) -> int:
    """Parse checkbox fields - returns 1 if checked/marked, 0 if not."""
    match = re.search(pattern, text, re.I)
    if not match:
        return 0
    
    value = match.group(1).strip()
    # Check for various indicators of a checked box
    if value in ["1", "X", "x", "✓", "☑", "checked", "yes", "true"]:
        return 1
    # Check for numeric values that might indicate selection
    if value.isdigit() and int(value) > 0:
        return 1
    
    return 0

def extract_prepared_by_fields(text: str) -> dict:
    """Extract prepared by fields using the template structure."""
    fields = {
        "name_last_first_middle_initial": value_after("a", text),
        "rank_grade": value_after("b", text),
        "duty_title_position": value_after("c", text),
        "unit": value_after("d", text),
        "work_email": value_after("e", text),
        "telephone": value_after("f", text),
        "uic_cin": value_after("g", text),
        "training_support_or_lesson_plan_or_opord": value_after("h", text),
        "signature_of_preparer": value_after("i", text),
    }
    
    # Clean up signature field - if it contains instructional text, set to null
    if fields["signature_of_preparer"]:
        sig_text = fields["signature_of_preparer"].lower()
        if any(phrase in sig_text for phrase in ["five steps", "risk management", "identify the hazards", "assess the hazards"]):
            fields["signature_of_preparer"] = None
    
    return fields

def extract_subtask_rows(text: str) -> list:
    """Extract and structure all subtask rows efficiently."""
    rows = []
    
    # Find the data section after the headers
    # Look for the pattern after "9. RESIDUAL RISK LEVEL" and before signatures/next section
    data_match = re.search(r'9\.\s*RESIDUAL.*?RISK LEVEL.*?\n(.*?)(?=\n10\.\s|\Z)', text, re.DOTALL | re.I)
    
    if not data_match:
        return rows
    
    data_section = data_match.group(1)
    
    # Split by the row separators (+ and - symbols at the start of rows)
    # Each row starts with the + / - pair and contains subtask, hazard, risk, controls, how/who, residual risk
    row_pattern = r'(?:^|\n)\+\s*\n-\s*\n(.*?)(?=(?:\n\+\s*\n-\s*\n)|\Z)'
    row_matches = re.findall(row_pattern, data_section, re.DOTALL)
    
    last_subtask_name = None

    for row_content in row_matches:
        prev_subtask_name = last_subtask_name
        if not row_content.strip():
            continue
            
        row = get_subtask_template()
        
        # Parse the row content - it's laid out in a tabular format
        lines = [line.strip() for line in row_content.split('\n') if line.strip()]
        
        if not lines:
            continue
        
        # Parse the row content more intelligently
        # Structure: SUBTASK | HAZARD | RISK | CONTROLS | HOW/WHO | RESIDUAL RISK
        
        subtask_lines = []
        hazard_lines = []
        controls_lines = []
        how_lines = []
        who_lines = []
        
        current_section = "subtask"
        initial_risk = None
        residual_risk = None
        found_first_risk = False
        
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            
            # Skip empty lines
            if not line:
                i += 1
                continue
            
            # Check for risk level indicators (single letters: M, H, L, EH, or single digits)
            if re.match(r'^([MHLE]+|[0-5])$', line) and len(line) <= 3:
                if not found_first_risk:
                    initial_risk = line
                    found_first_risk = True
                    current_section = "controls"
                else:
                    residual_risk = line
                    break
            elif line.startswith('How:'):
                current_section = "how"
                how_content = line[4:].strip()
                if how_content:
                    how_lines.append(how_content)
            elif line.startswith('Who:'):
                current_section = "who" 
                who_content = line[4:].strip()
                if who_content:
                    who_lines.append(who_content)
            elif line.startswith('-') and current_section in ["controls", "subtask"]:
                # This is likely a control measure, switch to controls section
                current_section = "controls"
                controls_lines.append(line)
            elif current_section == "subtask":
                # Still building subtask description
                subtask_lines.append(line)
            elif current_section == "controls":
                controls_lines.append(line)
            elif current_section == "how":
                how_lines.append(line)
            elif current_section == "who":
                who_lines.append(line)
            
            i += 1
        
        # If we only have subtask_lines and no clear separation, try to intelligently split
        if subtask_lines and not hazard_lines:
            # Look for common patterns to separate subtask from hazard
            combined_text = " ".join(subtask_lines)
            
            # Try to find a pattern where subtask and hazard are mentioned together
            # Example: "RANGE EXECUTION Negligent Discharge" -> subtask: "RANGE EXECUTION", hazard: "Negligent Discharge"
            if len(subtask_lines) >= 2:
                # Check if first line looks like a main category/subtask
                first_line = subtask_lines[0].upper()
                main_categories = ['RANGE EXECUTION', 'MOVEMENT', 'ENVIRONMENTAL', 'TRAINING']
                
                if any(cat in first_line for cat in main_categories):
                    # First line is likely the subtask category
                    hazard_lines = subtask_lines[1:]
                    subtask_lines = subtask_lines[:1]
                else:
                    # Try to split based on hazard keywords
                    hazard_keywords = COMMON_HAZARD_PREFIXES

                    skip_first_keywords = {'EVACUATION', 'MEDICAL'}
                    for idx, line in enumerate(subtask_lines):
                        upper_line = line.upper()
                        matched = next((keyword for keyword in hazard_keywords if keyword in upper_line.split()), None)
                        if not matched:
                            continue
                        if idx == 0 and matched in skip_first_keywords:
                            continue
                        hazard_lines = subtask_lines[idx:]
                        subtask_lines = subtask_lines[:idx] if idx > 0 else []
                        break

            # If hazard lines still weren't identified, assume the first one or two lines describe the subtask
            if subtask_lines and not hazard_lines and len(subtask_lines) > 1:
                keep = 1
                second_line_words = len(subtask_lines[1].split()) if len(subtask_lines) > 1 else 0
                if second_line_words and second_line_words <= 3 and len(subtask_lines) > 2:
                    hazard_candidate = " ".join(subtask_lines[1:]).upper()
                    hazard_cues = [
                        "HAZARD", "INJURY", "CASUALTY", "FAILURE", "FIRE", "ROLL", "ACCIDENT",
                        "DAMAGE", "LOSS", "PZ", "DZ", "LZ", "MARK", "ACCOUNT", "WEATHER",
                        "HELICOPTER", "AIRCRAFT", "SLING", "CARGO"
                    ]
                    if not any(cue in hazard_candidate for cue in hazard_cues):
                        keep = 2

                if len(subtask_lines) > keep:
                    hazard_lines = subtask_lines[keep:]
                    subtask_lines = subtask_lines[:keep]
        
        # Final cleanup: if we still don't have hazard_lines but have subtask_lines, 
        # it means the parsing above should have handled it already
        
        # Populate the row with properly concatenated text
        subtask_text = " ".join(subtask_lines) if subtask_lines else None
        hazard_text = " ".join(hazard_lines) if hazard_lines else None
        
        # If we have a combined subtask/hazard text, try to split it better
        if subtask_text and not hazard_text:
            # Look for pattern like "CATEGORY SpecificHazard" (case-insensitive)
            main_categories = [
                'RANGE EXECUTION', 'RANGE OPERATIONS', 'MOVEMENT TO', 'MOVEMENT',
                'VEHICLE MOVEMENT', 'AMMUNITION', 'ENVIRONMENTAL', 'MEDICAL',
                'TRAINING', 'FIRE CONTROL', 'WEAPONS HANDLING', 'COMMUNICATIONS',
                'NIGHT OPERATIONS', 'WEATHER'
            ]

            split_applied = False
            upper_subtask = subtask_text.upper()

            for category in main_categories:
                if upper_subtask.startswith(category):
                    # Preserve original casing for the subtask name by slicing
                    name_fragment = subtask_text[:len(category)].strip()
                    remainder = subtask_text[len(category):].strip()
                    remainder = remainder.lstrip("-–—:;/,").strip()

                    row["subtask"]["name"] = name_fragment or subtask_text
                    row["hazard"] = remainder or None
                    split_applied = True
                    break

            if not split_applied:
                # Fall back to splitting on obvious separators (colon, dash, slash, double spaces)
                parts = re.split(r"\s{2,}|\s[-–—/:]\s|:\s|/", subtask_text, maxsplit=1)
                if len(parts) == 2:
                    row["subtask"]["name"] = parts[0].strip()
                    row["hazard"] = parts[1].strip() or None
                else:
                    row["subtask"]["name"] = subtask_text
                    row["hazard"] = None
        else:
            row["subtask"]["name"] = subtask_text
            row["hazard"] = hazard_text

        if row["subtask"]["name"] and row["hazard"] and prev_subtask_name:
            tokens = [tok for tok in re.split(r"[\s,/]+", row["subtask"]["name"].upper()) if tok]
            leading_token = tokens[0] if tokens else ""
            if leading_token in COMMON_HAZARD_PREFIXES and leading_token not in HAZARD_STOPWORDS:
                candidate_subtask = row["subtask"]["name"].strip()
                hazard_body = (row["hazard"] or "").strip()

                hazard_body_clean = re.sub(r"[,/]+", " ", hazard_body)
                hazard_words = [w for w in hazard_body_clean.split() if w]
                subtask_words = [w for w in candidate_subtask.split() if w]

                should_prepend = False
                if not hazard_body:
                    should_prepend = True
                elif len(hazard_words) <= 2:
                    should_prepend = True
                elif subtask_words and subtask_words[-1].upper() in {"OF", "TO", "AND", "THE"}:
                    should_prepend = True

                if should_prepend:
                    combined_hazard = f"{candidate_subtask} {hazard_body}".strip()
                else:
                    combined_hazard = hazard_body

                combined_hazard = re.sub(r"\s+", " ", combined_hazard).strip()

                if combined_hazard:
                    row["hazard"] = combined_hazard
                    row["subtask"]["name"] = prev_subtask_name
        row["initial_risk_level"] = initial_risk
        
        # Group control lines into coherent paragraphs/sentences
        if controls_lines:
            # Clean and group related lines into complete sentences/paragraphs
            cleaned_controls = []
            current_paragraph = []
            
            for line in controls_lines:
                cleaned_line = re.sub(r'\s+', ' ', line.strip())
                if not cleaned_line:
                    continue
                
                # Check if this line starts a new sentence/paragraph
                # New sentences typically start with capital letters and complete words
                starts_new_sentence = False
                
                if current_paragraph:
                    last_line = current_paragraph[-1].lower()
                    current_line_lower = cleaned_line.lower()
                    
                    # More aggressive grouping - only start new sentence if we see clear separators
                    # Look for bullet points, dashes, or clearly independent sentences
                    is_bullet_point = cleaned_line.startswith(('-', '•', '*', '○', '▪'))
                    prev_is_bullet = current_paragraph[0].startswith(('-', '•', '*', '○', '▪')) if current_paragraph else False
                    
                    # Start new sentence only if:
                    # 1. This line is a bullet point and previous wasn't, OR
                    # 2. Previous was a bullet point and this is a new bullet point, OR  
                    # 3. Line starts with clear sentence starters after punctuation
                    if (is_bullet_point and not prev_is_bullet) or \
                       (is_bullet_point and prev_is_bullet) or \
                       (current_paragraph and 
                        (last_line.endswith('.') or last_line.endswith('!') or last_line.endswith('?')) and
                        cleaned_line[0].isupper() and
                        not current_line_lower.startswith(('and ', 'or ', 'but ', 'so ', 'yet ', 'for ', 'nor '))):
                        starts_new_sentence = True
                
                if starts_new_sentence:
                    # Finish current paragraph and start new one
                    if current_paragraph:
                        paragraph_text = ' '.join(current_paragraph).strip()
                        if paragraph_text:
                            cleaned_controls.append(paragraph_text)
                    current_paragraph = [cleaned_line]
                else:
                    # Continue current paragraph
                    current_paragraph.append(cleaned_line)
            
            # Add the last paragraph
            if current_paragraph:
                paragraph_text = ' '.join(current_paragraph).strip()
                if paragraph_text:
                    cleaned_controls.append(paragraph_text)
            
            row["control"]["values"] = cleaned_controls
        else:
            row["control"]["values"] = []
            
        # Concatenate how and who implementation details
        if how_lines:
            how_text = " ".join(how_lines)
            how_text = re.sub(r'\s+', ' ', how_text).strip()
            row["how_to_implement"]["how"]["values"] = [how_text] if how_text else []
        else:
            row["how_to_implement"]["how"]["values"] = []
            
        if who_lines:
            who_text = " ".join(who_lines)
            who_text = re.sub(r'\s+', ' ', who_text).strip()
            row["how_to_implement"]["who"]["values"] = [who_text] if who_text else []
        else:
            row["how_to_implement"]["who"]["values"] = []
            
        row["residual_risk_level"] = residual_risk
        
        # If hazard text was captured under the subtask column, treat it as hazard and reuse the previous subtask
        if prev_subtask_name and row["subtask"]["name"] and not row["hazard"]:
            hazard_candidate = row["subtask"]["name"].strip()
            if hazard_candidate:
                row["hazard"] = hazard_candidate
                row["subtask"]["name"] = prev_subtask_name

        # Carry forward subtask names when the current row omits the field text
        current_subtask = row["subtask"]["name"]
        if current_subtask and current_subtask.strip():
            last_subtask_name = current_subtask.strip()
        elif last_subtask_name:
            row["subtask"]["name"] = last_subtask_name

        rows.append(row)
    
    return rows

def calculate_overall_risk(residual_risks: list) -> str | None:
    """Calculate overall residual risk from individual residual risks."""
    if not residual_risks:
        return None
    
    # Risk level ordering (highest to lowest)
    risk_order = {"EH": 4, "H": 3, "M": 2, "L": 1}
    
    # Filter out None values and clean up the risk levels
    clean_risks = [r.strip().upper() for r in residual_risks if r and r.strip()]
    
    # Try categorical risk levels first
    categorical_risks = [r for r in clean_risks if r in risk_order]
    if categorical_risks:
        return max(categorical_risks, key=lambda r: risk_order[r])
    
    # Try numeric risk levels
    numeric_risks = []
    for r in clean_risks:
        try:
            numeric_risks.append(float(r))
        except (ValueError, TypeError):
            continue
    
    if numeric_risks:
        return str(int(max(numeric_risks)))
    
    return None

# ----------------------------
# Main parser shaped to your requested JSON structure
# ----------------------------
def parse_dd2977(text: str) -> dict:
    """Parse DD2977 form using efficient template-based structure."""
    # Start with the base template
    data = get_dd2977_template()

    # 1. Mission/Task Description - extract the actual content
    mission_match = re.search(r"1\.\s*MISSION/TASK.*?\n(.+?)(?=\n2\.\s*DATE|\Z)", text, re.I | re.DOTALL)
    if mission_match:
        mission_text = mission_match.group(1).strip()
        # Clean up common formatting issues
        mission_text = re.sub(r"\s+", " ", mission_text)
        data["mission_task_and_description"] = mission_text if mission_text else None
    
    # 2. Date Prepared - extract the actual date value  
    date_match = re.search(r"2\.\s*DATE PREPARED.*?\n(.+?)(?=\n3\.\s*PREPARED|\Z)", text, re.I | re.DOTALL)
    if date_match:
        date_text = date_match.group(1).strip()
        # Clean up common formatting
        date_text = re.sub(r"\s+", " ", date_text)
        data["date"] = date_text if date_text else None

    # 3. PREPARED BY — extract all fields efficiently
    data["prepared_by"] = extract_prepared_by_fields(text)

    # 4–9. Extract all subtask rows
    subtask_rows = extract_subtask_rows(text)

    data["subtasks"] = subtask_rows

    # 10. Overall residual risk (explicit, else calculate from individual risks)
    overall_residual = None
    overall_match = re.search(
        r"10\.\s*OVERALL\s+RESIDUAL\s+RISK LEVEL.*?:\s*(.+?)(?=\n11\.\s|\n12\.\s|\Z)",
        text,
        flags=re.I | re.DOTALL,
    )
    if overall_match:
        block = overall_match.group(1).strip()
        option_lines = [line.strip() for line in block.splitlines() if line.strip()]
        risk_options = {"EXTREMELY HIGH", "HIGH", "MEDIUM", "LOW"}

        if option_lines:
            selected_option = None

            if len(option_lines) == 1:
                single = option_lines[0].upper()
                if single in risk_options:
                    selected_option = single
            else:
                for raw_line in option_lines:
                    match = re.search(r"(EXTREMELY\s+HIGH|HIGH|MEDIUM|LOW)", raw_line, flags=re.I)
                    if not match:
                        continue
                    option = match.group(1).upper()
                    remainder = (raw_line[:match.start()] + raw_line[match.end():]).strip()
                    if re.search(r"[X✓☑1]", remainder):
                        selected_option = option
                        break
                    if re.search(r"\b(SELECTED|CHECKED|YES)\b", remainder, flags=re.I):
                        selected_option = option
                        break

                if not selected_option:
                    marked_lines = [
                        line for line in option_lines
                        if re.sub(r"^[•*\-\u25A0\u25A1\[\]()]+", "", line).strip().upper() in risk_options
                    ]
                    if len(marked_lines) == 1:
                        selected_option = re.sub(
                            r"^[•*\-\u25A0\u25A1\[\]()]+", "", marked_lines[0]
                        ).strip().upper()

            if selected_option:
                overall_residual = selected_option

    if not overall_residual:
        residual_risks = [row["residual_risk_level"] for row in subtask_rows]
        overall_residual = calculate_overall_risk(residual_risks)
    data["overall_residual_risk_level"] = overall_residual

    # 11. Overall supervision plan
    data["overall_supervision_plan"] = pick(
        r"11\.\s*OVERALL SUPERVISION PLAN.*?:\s*(.+?)(?=\n(?:APPROVE|DISAPPROVE|12\.|14\.|15\.)|\Z)",
        text, flags=re.I|re.DOTALL
    )

    # Try to refine overall residual risk using narrative text if available
    narrative = data["overall_supervision_plan"] or ""
    if narrative:
        plan_match = re.search(
            r"overall\s+residual\s+risk[^.]*?(?:assessed\s+as|assessed\s+to\s+be|is)\s+(extremely\s+high|very\s+high|high|medium|moderate|low|very\s+low|negligible)",
            narrative,
            flags=re.I
        )
        if plan_match:
            plan_level = plan_match.group(1).strip().upper()
            risk_map = {
                "EXTREMELY HIGH": "EXTREMELY HIGH",
                "VERY HIGH": "HIGH",
                "HIGH": "HIGH",
                "MEDIUM": "MEDIUM",
                "MODERATE": "MEDIUM",
                "LOW": "LOW",
                "VERY LOW": "LOW",
                "NEGLIGIBLE": "LOW",
            }
            mapped_level = risk_map.get(plan_level)
            if mapped_level:
                data["overall_residual_risk_level"] = mapped_level

    # 12. Approval OR Disapproval (checkbox fields - 1 if checked, 0 if not)
    data["approval_or_disapproval_of_mission_or_task"]["approve"] = parse_checkbox_value(text, r"\bAPPROVE:\s*([0-9X✓☑x]+)")
    data["approval_or_disapproval_of_mission_or_task"]["disapprove"] = parse_checkbox_value(text, r"\bDISAPPROVE:\s*([0-9X✓☑x]+)")

    # Fallback: infer approval/disapproval from signature block if checkboxes are blank
    approval_section_match = re.search(
        r"12\.\s*APPROVAL\s+OR\s+DISAPPROVAL.*?(?=\n13\.\s|\nRISK ASSESSMENT MATRIX|\Z)",
        text,
        flags=re.I | re.DOTALL,
    )
    if approval_section_match:
        approval_section = approval_section_match.group(0)
        approve_mark = data["approval_or_disapproval_of_mission_or_task"]["approve"]
        disapprove_mark = data["approval_or_disapproval_of_mission_or_task"]["disapprove"]

        if not approve_mark and not disapprove_mark:
            has_signature = bool(re.search(r"Digitally\s+signed\s+by", approval_section, re.I))

            # Remove static labels to avoid false positives when scanning for phrases
            filtered_section_lines = []
            for line in approval_section.splitlines():
                if re.match(r"\s*(12\.\s*APPROVAL\s+OR\s+DISAPPROVAL.*|APPROVE|DISAPPROVE)\s*\Z", line, re.I):
                    continue
                filtered_section_lines.append(line)
            filtered_section = "\n".join(filtered_section_lines)

            # Look for explicit language indicating the mission was rejected
            disapproval_phrases = [
                r"\bDISAPPROVED\b",
                r"\bDISAPPROVAL\b",
                r"\bNOT\s+APPROVED\b",
                r"\bDO\s+NOT\s+APPROVE\b",
                r"\bWITHHOLD\s+APPROVAL\b",
            ]
            mentions_disapprove = any(
                re.search(phrase, filtered_section, re.I) for phrase in disapproval_phrases
            )

            if has_signature and not mentions_disapprove:
                data["approval_or_disapproval_of_mission_or_task"]["approve"] = 1
            elif mentions_disapprove and not has_signature:
                data["approval_or_disapproval_of_mission_or_task"]["disapprove"] = 1

    return data

# ----------------------------
# Naming & output directory
# ----------------------------
def slugify(s: str) -> str:
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+","-",s)
    return re.sub(r"-{2,}","-",s).strip("-") or "draw"

def find_date_in_name(stem: str) -> str | None:
    m = re.search(r"\b(20\d{6})\b", stem) or re.search(r"\b(\d{8})\b", stem)
    return m.group(1) if m else None

def normalize_date_to_yyyymmdd(s: str) -> str | None:
    if not s: return None
    fmts = ["%Y-%m-%d","%m/%d/%Y","%d/%m/%Y","%d%b%Y","%d%b%y","%d %b %Y","%d %B %Y"]
    for f in fmts:
        try:
            return datetime.strptime(s.strip(), f).strftime("%Y%m%d")
        except Exception:
            pass
    m = re.search(r"(20\d{2})[-/.]?(0[1-9]|1[0-2])[-/.]?(0[1-9]|[12]\d|3[01])", s)
    return "".join(m.groups()) if m else None

def build_outpath(pdf_path: Path, parsed: dict, outdir: Path) -> Path:
    outdir.mkdir(parents=True, exist_ok=True)
    stem = pdf_path.stem
    yyyymmdd = find_date_in_name(stem) or normalize_date_to_yyyymmdd(parsed.get("date") or "") \
               or datetime.fromtimestamp(pdf_path.stat().st_mtime).strftime("%Y%m%d")
    return outdir / f"{yyyymmdd}-{slugify(stem)}.json"

# ----------------------------
# CLI
# ----------------------------
def process_pdf(pdf_path: Path, outdir: Path, force_ocr: bool = False) -> bool:
    """Process a single PDF file and return success status."""
    try:
        # Check for XFA-based PDFs first
        parsed = parse_dd2977_xfa(pdf_path)
        if parsed:
            outpath = build_outpath(pdf_path, parsed, outdir)
            outdir.mkdir(parents=True, exist_ok=True)
            outpath.write_text(json.dumps(parsed, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"Wrote: {outpath}")
            return True

        text = extract_text_multibackend(pdf_path, force_ocr=force_ocr)
        if "Please wait" in text and "Adobe Reader" in text:
            print(
                "Detected an XFA-based PDF. Install optional dependencies (pikepdf, lxml) "
                "and re-run using the project virtual environment to enable direct parsing.",
                file=sys.stderr,
            )
            return False
        if not text.strip():
            print(f"Failed to extract text from: {pdf_path}", file=sys.stderr)
            return False

        parsed = parse_dd2977(text)
        outpath = build_outpath(pdf_path, parsed, outdir)
        outpath.write_text(json.dumps(parsed, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Wrote: {outpath}")
        return True
    except Exception as e:
        print(f"Error processing {pdf_path}: {e}", file=sys.stderr)
        return False

def batch_process(input_dir: Path, outdir: Path, force_ocr: bool = False) -> None:
    """Process all PDF files in a directory and its subdirectories."""
    pdf_files = list(input_dir.rglob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in: {input_dir}", file=sys.stderr)
        return

    print(f"Found {len(pdf_files)} PDF files to process...")
    
    success_count = 0
    for pdf_path in pdf_files:
        if process_pdf(pdf_path, outdir, force_ocr):
            success_count += 1
    
    print(f"Successfully processed {success_count}/{len(pdf_files)} files")

def main():
    ap = argparse.ArgumentParser(description="Parse DD2977 DRAW forms to structured JSON")
    ap.add_argument("input", help="Path to a DRAW (DD2977) PDF file or directory containing PDFs")
    ap.add_argument("--outdir", default="PARSED_DRAWS", help="Directory for JSON output")
    ap.add_argument("--force-ocr", action="store_true", help="Force OCR backend")
    ap.add_argument("--batch", action="store_true", help="Process all PDFs in the input directory")
    args = ap.parse_args()

    input_path = Path(args.input)
    outdir = Path(args.outdir)

    if not input_path.exists():
        print(f"Input not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    if args.batch or input_path.is_dir():
        batch_process(input_path, outdir, args.force_ocr)
    else:
        if not process_pdf(input_path, outdir, args.force_ocr):
            sys.exit(2)

if __name__ == "__main__":
    main()
