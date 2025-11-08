#!/usr/bin/env python3
"""Batch parser for CONOP PPTX files and DRAW PDFs.

This script walks through subdirectories beneath one or more base folders,
parses the first CONOP PPTX and first DRAW PDF found in each directory, and
writes the results to their respective output folders. Both parsed outputs
receive a `source_directory_id` field so downstream consumers can link the
paired artifacts back to the originating directory.
"""

import argparse
import json
from datetime import datetime
from pathlib import Path

from parse_conop import extract_text_from_pptx, parse_conop_sections
from parse_draw import (
    extract_text_multibackend,
    parse_dd2977,
    parse_dd2977_xfa,
    slugify,
)


def parse_draw_file(pdf_path: Path) -> tuple[dict | None, str | None]:
    """Parse a DRAW PDF into structured data, returning (payload, error)."""
    try:
        parsed = parse_dd2977_xfa(pdf_path)
    except Exception as exc:  # pragma: no cover - defensive logging
        return None, f"XFA parsing failed: {exc}"

    if parsed is None:
        try:
            text = extract_text_multibackend(pdf_path)
        except Exception as exc:  # pragma: no cover - defensive logging
            return None, f"Text extraction failed: {exc}"

        if "Please wait" in text and "Adobe Reader" in text:
            return None, "Detected XFA form but XFA dependencies are missing"
        if not text.strip():
            return None, "Unable to extract text"
        parsed = parse_dd2977(text)

    parsed["source_pdf"] = str(pdf_path.resolve())
    return parsed, None


def parse_conop_file(ppt_path: Path) -> tuple[dict | None, str | None]:
    """Parse a CONOP PPTX into structured data, returning (payload, error)."""
    try:
        text = extract_text_from_pptx(ppt_path)
    except Exception as exc:  # pragma: no cover - defensive logging
        return None, f"Failed to open PPTX: {exc}"

    if not text:
        return None, "PPTX contained no extractable text"

    sections = parse_conop_sections(text)
    payload = {
        "filename": ppt_path.name,
        "path": str(ppt_path.resolve()),
        "sections": sections,
    }
    return payload, None


def find_first_matching(directory: Path, suffix: str) -> Path | None:
    """Return the first file (sorted) in directory that ends with suffix."""
    candidates = sorted(
        p for p in directory.iterdir() if p.is_file() and p.name.lower().endswith(suffix)
    )
    return candidates[0] if candidates else None


def iter_directories(base_path: Path) -> list[Path]:
    return sorted([p for p in base_path.iterdir() if p.is_dir()])


def process_directory(
    directory: Path,
    dir_id: int,
    draws_outdir: Path,
    conops_outdir: Path,
    skipped: list[dict],
) -> None:
    """Parse the first CONOP and DRAW within a directory."""
    ppt_path = find_first_matching(directory, ".pptx")
    pdf_path = find_first_matching(directory, ".pdf")

    base_dir_name = directory.parent.name

    if ppt_path:
        payload, error = parse_conop_file(ppt_path)
        if payload:
            data = payload
            data["source_directory_id"] = dir_id
            data["source_directory_name"] = directory.name
            data["source_base_directory"] = base_dir_name
            file_slug = slugify(ppt_path.stem)
            outpath = conops_outdir / f"{dir_id:04d}-{file_slug}-conop.json"
            conops_outdir.mkdir(parents=True, exist_ok=True)
            outpath.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"Wrote CONOP: {outpath}")
        else:
            reason = error or "Unknown failure"
            print(f"Skipping CONOP {ppt_path}: {reason}")
            skipped.append(
                {
                    "type": "CONOP",
                    "path": str(ppt_path.resolve()),
                    "reason": reason,
                    "source_directory_id": dir_id,
                    "source_directory_name": directory.name,
                    "source_base_directory": base_dir_name,
                }
            )
    else:
        print(f"No CONOP PPTX found in {directory}")

    if pdf_path:
        payload, error = parse_draw_file(pdf_path)
        if payload:
            data = payload
            data["source_directory_id"] = dir_id
            data["source_directory_name"] = directory.name
            data["source_base_directory"] = base_dir_name
            file_slug = slugify(pdf_path.stem)
            outpath = draws_outdir / f"{dir_id:04d}-{file_slug}-draw.json"
            draws_outdir.mkdir(parents=True, exist_ok=True)
            outpath.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"Wrote DRAW: {outpath}")
        else:
            reason = error or "Unknown failure"
            print(f"Skipping DRAW {pdf_path}: {reason}")
            skipped.append(
                {
                    "type": "DRAW",
                    "path": str(pdf_path.resolve()),
                    "reason": reason,
                    "source_directory_id": dir_id,
                    "source_directory_name": directory.name,
                    "source_base_directory": base_dir_name,
                }
            )
    else:
        print(f"No DRAW PDF found in {directory}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Batch parse paired CONOP PPTX and DRAW PDF files by directory."
    )
    parser.add_argument(
        "base_dirs",
        nargs="*",
        default=[
            "1-2CR CONOPS&DRAWs - Copy",
            "3-2CR CONOPs&DRAWs",
        ],
        help="One or more base directories containing subdirectories with CONOPs and DRAWs.",
    )
    parser.add_argument(
        "--draws-outdir",
        default="PARSED_DRAWS",
        help="Directory where parsed DRAW JSON files will be written.",
    )
    parser.add_argument(
        "--conops-outdir",
        default="PARSED_CONOPS",
        help="Directory where parsed CONOP JSON files will be written.",
    )
    parser.add_argument(
        "--skip-report",
        default="skipped_documents_report.json",
        help="Path to write a JSON report summarizing unparsed documents.",
    )
    args = parser.parse_args()

    base_paths = [Path(p) for p in args.base_dirs]
    draws_outdir = Path(args.draws_outdir)
    conops_outdir = Path(args.conops_outdir)
    skip_report_path = Path(args.skip_report)

    skipped: list[dict] = []
    dir_id = 1
    for base_path in base_paths:
        if not base_path.exists():
            print(f"Base directory not found: {base_path}")
            continue

        directories = iter_directories(base_path)
        if not directories:
            print(f"No subdirectories found under {base_path}")
            continue

        for directory in directories:
            print(f"\nProcessing directory {directory} (ID {dir_id})")
            process_directory(directory, dir_id, draws_outdir, conops_outdir, skipped)
            dir_id += 1

    if dir_id == 1:
        print("No valid base directories processed.")
    else:
        skipped = [entry for entry in skipped if entry.get("path")]  # normalize
        if skipped:
            print("\nThe following documents could not be parsed:")
            for entry in skipped:
                print(
                    f" - {entry['type']}: {entry['path']}\n   Reason: {entry['reason']}"
                )
            report_payload = {
                "generated_at": datetime.utcnow().isoformat() + "Z",
                "total_directories_processed": dir_id - 1,
                "skipped_count": len(skipped),
                "skipped_documents": skipped,
            }
            skip_report_path.parent.mkdir(parents=True, exist_ok=True)
            skip_report_path.write_text(
                json.dumps(report_payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            print(f"\nSkipped document report written to {skip_report_path.resolve()}")
        else:
            print("\nAll detected documents were parsed successfully.")


if __name__ == "__main__":
    main()
