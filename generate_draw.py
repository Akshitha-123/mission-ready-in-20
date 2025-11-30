#!/usr/bin/env python3
"""
generate_draw.py

User Guide:
1. Export `DB_HOST`, `DB_PORT`, `DB_NAME`, `DB_USER`, and `DB_PASSWORD` with your PostgreSQL credentials.
2. Call `export OLLAMA_API_KEY="API_KEY_HERE"` to set your Ollama Cloud API key.
3. Run this script with `python generate_draw.py` to ingest training data and generate DRAW outputs.
4. I've currently set it to generate DRAW for "0269-drivers-training-conop-conop.json" as an example.
5. The draw json will be saved to "draw_output.json".

Pipeline:
1. Reads all merged CONOP→DRAW JSON files from merged_conop_draws/.
2. Computes embeddings for CONOP sections.
3. Stores JSON + embedding in PostgreSQL (pgvector).
4. generate_draw_for_conop(new_conop_json):
      - Embed new CONOP
      - Retrieve similar CONOPs via pgvector
      - Build few-shot prompt using retrieved pairs
      - Generate DRAW using Ollama Cloud
      - Save generated DRAW to a JSON file
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Dict, Sequence

import psycopg2
from psycopg2.extras import Json
from sentence_transformers import SentenceTransformer
from ollama import Client

# ===========================
# CONFIG
# ===========================

TRAINING_DIR = "MERGED_CONOPS_DRAWS"

DB_CONN = {
    "host": os.getenv("DB_HOST", "localhost"),
    "port": int(os.getenv("DB_PORT", 5432)),
    "dbname": os.getenv("DB_NAME", "mrit_db"),
    "user": os.getenv("DB_USER", "arpithaprakash"),
    "password": os.getenv("DB_PASSWORD", "MRI-20"),
}

# if not DB_CONN["password"]:
#     raise RuntimeError(
#         "DB_PASSWORD environment variable is not set. Please export your database credentials before running generate_draw.py."
#     )

OLLAMA_CLOUD_URL = "https://api.ollama.com/v1/chat/completions"
OLLAMA_MODEL = "llama3.1:70b"
EMBED_MODEL = SentenceTransformer("sentence-transformers/all-mpnet-base-v2")


# ===========================
# DB SETUP
# ===========================

def init_db():
    conn = psycopg2.connect(**DB_CONN)
    cur = conn.cursor()

    cur.execute("""
        DROP TABLE IF EXISTS conop_draw_pairs;
    """)

    cur.execute("""
        CREATE TABLE conop_draw_pairs (
            id SERIAL PRIMARY KEY,
            conop_json JSONB NOT NULL,
            draw_json JSONB NOT NULL,
            embedding VECTOR(768)
        );
    """)

    conn.commit()
    cur.close()
    conn.close()


# ===========================
# HELPERS
# ===========================

def extract_conop_text(pair_json):
    """
    Extract all section text from either:
      1) merged CONOP→DRAW pair: { "conops": { "sections": {...}, ... }, "draw": {...} }
      2) standalone CONOP: { "sections": {...}, ... }
    """
    sections = {}
    if "conops" in pair_json and "sections" in pair_json["conops"]:
        sections = pair_json["conops"]["sections"]
    elif "sections" in pair_json:
        sections = pair_json["sections"]

    if not sections:
        return None

    return "\n".join([str(v) for v in sections.values()])



def embed_text(text: str) -> Sequence[float]:
    emb = EMBED_MODEL.encode([text], normalize_embeddings=True)[0]
    return emb.tolist()


# ===========================
# INGEST DIRECTORY OF PAIRS
# ===========================

def ingest_directory(directory=TRAINING_DIR):
    """
    Loads ALL *.json files in merged_conop_draws/ and inserts into PostgreSQL.
    Skips insertion if:
      - conop_text is missing
      - draw.subtasks == []
      - draw.prepared_by.training_support_or_lesson_plan_or_opord is null
      - database insertion fails
    """

    conn = psycopg2.connect(**DB_CONN)
    cur = conn.cursor()

    files = [f for f in os.listdir(directory) if f.endswith(".json")]
    print(f"[INFO] Loading {len(files)} training files from {directory}/")

    for fname in files:
        path = os.path.join(directory, fname)
        try:
            with open(path, "r") as f:
                pair = json.load(f)

            # ---- Extract required fields ----
            conop_text = extract_conop_text(pair)
            if conop_text is None:
                print(f"[SKIP] File {fname} missing 'conops' or 'sections'. Skipping.")
                continue

            # DRAW
            draw = pair.get("draw", {})

            # ---- Skip if draw.subtasks is [] ----
            if isinstance(draw, dict) and draw.get("subtasks") == []:
                print(f"[SKIP] {fname} skipped because draw.subtasks is empty ([]).")
                continue

            # ---- Compute embeddings ----
            emb = embed_text(conop_text)

            # ---- Insert into database ----
            try:
                cur.execute("""
                    INSERT INTO conop_draw_pairs (conop_json, draw_json, embedding)
                    VALUES (%s, %s, %s)
                """, (Json(pair["conops"]), Json(draw), emb))

                print(f"[OK] Inserted {fname}")

            except Exception as e:
                print(f"[ERROR] Failed to insert {fname}: {e}. Skipping.")
                continue

        except Exception as e:
            print(f"[ERROR] Failed to process {fname}: {e}. Skipping.")
            continue

    conn.commit()
    cur.close()
    conn.close()

# ===========================
# RAG RETRIEVAL
# ===========================

def retrieve_similar_conops(query_emb, k=5):
    conn = psycopg2.connect(**DB_CONN)
    cur = conn.cursor()

    # Convert Python list to PostgreSQL ARRAY literal string
    query_emb_str = "ARRAY[" + ",".join([str(x) for x in query_emb]) + "]::vector"

    sql = f"""
        SELECT conop_json, draw_json
        FROM conop_draw_pairs
        ORDER BY embedding <-> {query_emb_str}
        LIMIT %s
    """
    cur.execute(sql, (k,))

    rows = cur.fetchall()
    cur.close()
    conn.close()

    return rows


# ===========================
# FEW-SHOT PROMPT
# ===========================

def build_prompt(context_pairs, new_conop):
    prompt = (
        "You are a military risk assessment assistant. "
        "Given training examples of CONOP → DRAW pairs, generate a DRAW for the new CONOP.\n\n"
        "=== TRAINING EXAMPLES ===\n"
    )

    for conop_json, draw_json in context_pairs:
        print(draw_json)
        conop_sections = conop_json.get("sections") or conop_json.get("conops", {}).get("sections", {})
        prompt += (
            "\n---\n"
            "CONOP SECTIONS:\n"
            f"{json.dumps(conop_sections, indent=2)}\n\n"
            "DRAW OUTPUT:\n"
            f"{json.dumps(draw_json, indent=2)}\n"
        )

    # New CONOP sections
    new_sections = new_conop.get("sections") or new_conop.get("conops", {}).get("sections", {})
    prompt += (
        "\n\n=== NEW CONOP ===\n"
        f"{json.dumps(new_sections, indent=2)}\n\n"
        "Now output the complete DRAW JSON.\n"
    )

    return prompt


# ===========================
# OLLAMA CLOUD GENERATION
# ===========================

def call_ollama_cloud(prompt):
    api_key = os.environ.get("OLLAMA_API_KEY", "9f4e1f135c35424f82fde6596ae12569.krawhX9x4C3ua3Qn2snMmucQ")
    if not api_key:
        raise RuntimeError("OLLAMA_API_KEY environment variable is not set. Please export your Ollama Cloud key before generating a DRAW.")

    client = Client(
        host="https://ollama.com",
        headers={"Authorization": f"Bearer {api_key}"}
    )

    messages = [
            {"role": "system", "content": "Respond only with valid JSON."},
            {"role": "user", "content": prompt}
    ]
    output = []

    for part in client.chat('gpt-oss:120b', messages=messages, stream=True):
        output.append(part['message']['content'])

    output = ''.join(output)
    return output   

# ===========================
# MAIN GENERATION
# ===========================

def generate_draw_for_conop(new_conop: Dict[str, Any], output_path: str | None = "generated_draw.json"):
    """
    Input: merged CONOP JSON:
        { "conops": {...}, "draw": {...?} }
    Output: separate DRAW JSON file
    """

    # 1. Embed
    text = extract_conop_text(new_conop)
    if text is None:
        print(f"[ERROR] Provided CONOP JSON is missing 'conops' or 'sections'. Aborting generation.")
        return None

    query_emb = embed_text(text)

    # 2. Retrieve top similar conops
    retrieved = retrieve_similar_conops(query_emb)

    # 3. Build RAG prompt
    prompt = build_prompt(retrieved, new_conop)
    print(f"[DEBUG] Prompt:\n{prompt}\n")

    # 4. Call Ollama Cloud
    output = call_ollama_cloud(prompt)

    # 5. Parse JSON safely
    try:
        draw_json = json.loads(output)
    except json.JSONDecodeError:
        raise RuntimeError(f"Model output was not valid JSON:\n{output}")

    # 6. Save to file
    if output_path:
        output_file = Path(output_path)
        output_file.write_text(json.dumps(draw_json, indent=2), encoding="utf-8")
        print(f"[DONE] DRAW written to {output_file}")

    return draw_json

# ===========================
# SCRIPT ENTRY POINT
# ===========================

if __name__ == "__main__":
    init_db()

    # Step 1: Ingest all training files:
    ingest_directory(TRAINING_DIR)

    # Step 2: Example generation (uncomment for use):
    new_conop = json.load(open("0269-drivers-training-conop-conop.json"))
    generate_draw_for_conop(new_conop, "draw_output.json")

