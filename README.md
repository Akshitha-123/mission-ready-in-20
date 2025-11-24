# Capstone Parsing Utilities

Utilities for converting CONOP PowerPoint files and DD2977 DRAW PDF packages into structured JSON.

## Environment Setup

1. Install Python 3.11 or later.
2. (Recommended) Create and activate a virtual environment:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Scripts

### `parse_conop`

Parses a single CONOP PowerPoint file into structured sections.

```bash
python parse_conop.py path/to/file.pptx --output-dir PARSED_CONOPS
```

- Writes a JSON file named after the source slide deck with slugified metadata.
- Requires `python-pptx` and other packages from `requirements.txt`.

### `parse_draw.py`

Parses a single DD2977 DRAW PDF. Supports both XFA-based and text-based documents.

```bash
python parse_draw.py path/to/file.pdf --output-dir PARSED_DRAWS
```

- Attempts XFA extraction first, then falls back to text extraction.
- Produces normalized risk and approval information.

### `batch_parse_conops_draws.py`

Batch processes paired CONOP/DRAW documents grouped in subdirectories.

```bash
python batch_parse_conops_draws.py \
    "1-2CR CONOPS&DRAWs - Copy" \
    "3-2CR CONOPs&DRAWs" \
    --conops-outdir PARSED_CONOPS \
    --draws-outdir PARSED_DRAWS \
    --skip-report skipped_documents_report.json
```

- Each subdirectory receives a sequential `source_directory_id` shared by the CONOP and DRAW outputs.
- JSON outputs include `source_directory_name`, `source_directory_id`, and `source_base_directory` metadata.
- Any files that fail to parse are recorded in both the console output and the skip report.
- The skip report is a JSON document containing:
  - generation timestamp (UTC)
  - total processed directories
  - count of skipped files
  - detailed list of skipped documents with file paths, reasons, and source metadata

## Merging and Database Upload

### Merging CONOPS and DRAWs

Use the merging script to combine CONOPS and DRAW JSON files from the same directory into a single file in `MERGED_CONOPS_DRAWS/`. Each merged file is named `<directory_id>-merged.json`.

### Uploading to PostgreSQL

1. Ensure PostgreSQL is installed and running (`brew install postgresql` and `brew services start postgresql@14`).
2. Create a database and table:
   - In `psql`:
     ```sql
     CREATE DATABASE mrit_db OWNER "arpithaprakash";
     \c mrit_db
     CREATE TABLE merged_conops_draws (
         id SERIAL PRIMARY KEY,
         source_directory_id TEXT NOT NULL,
         merged_data JSONB NOT NULL
     );
     ```
3. Export the database environment variables used by the scripts (or copy `env.sh.example` to `env.sh` and fill in your credentials):
   ```bash
   export DB_HOST=localhost
   export DB_PORT=5432
   export DB_NAME=mrit_db
   export DB_USER=arpithaprakash
   export DB_PASSWORD=your_password
   ```
4. Run the script:
   ```sh
   python3 upload_merged_json_to_postgres.py
   ```
5. Verify upload in `psql`:
   ```sql
   SELECT COUNT(*) FROM merged_conops_draws;
   SELECT * FROM merged_conops_draws LIMIT 1;
   ```

## Web Application

The `conops-to-draw-main/` directory contains the React/Vite frontend that interacts with a FastAPI backend exposed via `api_server.py`.

### Backend API

#### One-Time Backend Setup

1. Install Python dependencies as described above.
2. (Required for inline CONOPS previews) Install LibreOffice so the backend can convert PPTX files to PDF:
   ```bash
   brew install --cask libreoffice
   ```
3. (Required for DRAW previews) Ensure PyMuPDF is installed. It is already listed in `requirements.txt`, but if you skip the bundled requirements install for any reason run:
   ```bash
   pip install pymupdf
   ```
   The backend uses PyMuPDF to rasterize a separate `*-preview.pdf` copy of each generated DRAW so the React iframe shows filled text even in read-only viewers.
4. Ensure PostgreSQL with the pgvector extension is available (the project uses a Docker-hosted instance listening on `localhost:5432`). Stop any Homebrew Postgres service so it does not compete for the port:
   ```bash
   brew services stop postgresql@14 2>/dev/null
   export PGHOST=localhost PGPORT=5432 PGUSER=arpithaprakash PGPASSWORD=MRI-20
   createdb mrit_db 2>/dev/null
   psql -h "$PGHOST" -p "$PGPORT" -U "$PGUSER" -d mrit_db \
     -c "CREATE EXTENSION IF NOT EXISTS vector;"
   ```
5. Seed the DRAW training data (drops/recreates `conop_draw_pairs` and ingests everything under `MERGED_CONOPS_DRAWS/`):
   ```bash
   python generate_draw.py
   ```
6. Export your Ollama key and database credentials so the generator and API can connect to external services (store these in `.env` or `env.sh` if desired):
   ```bash
   export OLLAMA_API_KEY=9f4e1f135c35424f82fde6596ae12569.krawhX9x4C3ua3Qn2snMmucQ
   export DB_HOST=localhost DB_PORT=5432 DB_NAME=mrit_db DB_USER=arpithaprakash DB_PASSWORD=MRI-20
   ```
7. The `/api/conops/upload` endpoint accepts a `.pptx` file, stores it, parses it via `parse_conop.py`, converts it to PDF for preview, and (when the prerequisites above are met) generates a DRAW JSON by calling the Ollama-assisted `generate_draw.py` pipeline. Each successful `/api/conops/generate-draw` call writes two PDFs into `generated_draws/`:
   - `<deck>-draw-<uuid>.pdf` — the editable DD 2977 with live form fields, used for downloads/exports.
   - `<deck>-draw-<uuid>-preview.pdf` — a PyMuPDF-rendered copy baked to page graphics, used exclusively by the React preview iframe to ensure filled content appears in every browser.

#### Daily Backend Run (Quick Start)

After completing the one-time setup, an end-to-end test session only needs three commands:

```bash
cd /Users/arpithaprakash/mission-ready-in-20
source env.sh          # activates .venv and exports OLLAMA_API_KEY/DB_*
python generate_draw.py
uvicorn api_server:app --reload
```

`env.sh` ensures you are inside the correct virtual environment and that the Ollama key and database credentials are present. Rerun these commands whenever you start a new shell.

### Frontend

1. Install Node.js dependencies:
   ```bash
   cd conops-to-draw-main
   npm install
   ```
2. (Optional) Set `VITE_API_URL` in a `.env` file if the backend is not accessible via the default `http://127.0.0.1:8000`. When unset, the Vite dev server proxies `/api/*` and `/uploads/*` calls to `http://127.0.0.1:8000` automatically.
3. Start the development server:
   ```bash
   npm run dev
   ```
4. Use the UI to upload a CONOPS PowerPoint. The "Original CONOPS" panel will display the parsed content returned by the backend, and the file becomes non-editable in the viewer.

#### Frontend Quick Start

For day-to-day runs, stay inside `conops-to-draw-main/` and execute:

```bash
npm install   # first run only, safe to repeat
npm run dev
```

Keep this dev server running while you exercise the workflow in the browser.

### End-to-End Test Flow

Run through this checklist whenever you need to verify the full stack:

1. **Backend prerequisites (run once per shell):**
   ```bash
   cd /Users/arpithaprakash/mission-ready-in-20
   source .venv/bin/activate
   export PGHOST=localhost PGPORT=5432 PGUSER=arpithaprakash PGPASSWORD=MRI-20
   export DB_HOST=localhost DB_PORT=5432 DB_NAME=mrit_db DB_USER=arpithaprakash DB_PASSWORD=MRI-20
   psql -h "$PGHOST" -p "$PGPORT" -U "$PGUSER" -d mrit_db -c "CREATE EXTENSION IF NOT EXISTS vector;"
   python generate_draw.py
   export OLLAMA_API_KEY=9f4e1f135c35424f82fde6596ae12569.krawhX9x4C3ua3Qn2snMmucQ
   ```
2. **Start services:**

   ```bash
   # Terminal 1 (backend)
   uvicorn api_server:app --reload

   # Terminal 2 (frontend)
   cd conops-to-draw-main
   npm run dev
   ```

3. **Manual verification:** open the Vite URL (default `http://localhost:8080`), upload a `.pptx`, confirm the progress tracker moves to Step 2, the CONOPS PDF preview appears once LibreOffice finishes, and the AI-generated DRAW populates the right-hand panel (progress advances to Step 3). Use the Export button to download the generated draft and review it locally.

## Outputs

- `PARSED_CONOPS/` and `PARSED_DRAWS/` hold JSON outputs generated by the scripts above.
- `MERGED_CONOPS_DRAWS/` contains merged JSON files for each directory.
- Filenames begin with the four-digit directory ID followed by a slugged description and a suffix indicating `-conop`, `-draw`, or `-merged`.
- Open `skipped_documents_report.json` (or the custom path supplied via `--skip-report`) to review any failures.

## Troubleshooting

- **Missing records in database**: Check for invalid or empty JSON files in `MERGED_CONOPS_DRAWS/`. Review the script output for errors or skipped files. Use logging to identify files that failed to upload.
- **"PPTX contained no extractable text"**: The PowerPoint file is likely corrupted or saved in the legacy `.ppt` format. Re-save it as a valid `.pptx` and rerun the parser.
- **"Unable to extract text" (PDF)**: Provide a text-based PDF (not a scanned image) or ensure XFA dependencies listed in `requirements.txt` are installed.
- Re-run the batch script after fixing problematic files to generate the remaining JSON outputs.
