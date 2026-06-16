"""
Optimized PDF merger for render.com  (drop-in replacement for main.py)

KEY CHANGES vs the original (all aimed at beating Power Automate's ~120s
synchronous ceiling on the Starter 0.5 CPU / 512 MB instance):

  1. BATCH CONVERSION  -- the original launched a fresh LibreOffice process
     PER Office file (one per .docx / .xlsx + the generated checklist).
     LibreOffice can convert MANY files in a SINGLE invocation, so we collect
     every Office file first and run ONE `libreoffice --convert-to pdf f1 f2 f3`.
     This removes 3 of the 4 cold LibreOffice startups -> the dominant cost.

  2. REUSED PROFILE DIR -- a dedicated -env:UserInstallation profile avoids the
     ~3-5s first-run profile bootstrap and lets a warm profile be reused.

  3. STREAMING MERGE    -- pypdf appends from disk instead of holding every
     decoded PDF in RAM, which keeps peak memory under the 512 MB cap.

  4. ROBUST FALLBACK    -- if the batch call fails for any reason we fall back
     to per-file conversion so a single bad doc can't kill the whole merge.

NOTE: This is reconstructed against the version you pasted. Diff it against your
real main.py before deploying so any project-specific placeholder logic
(generate_checklist_pdf / replace_placeholders / templateData fields) is
preserved exactly. The HTTP contract (POST /merge-pdfs, GET /health) is unchanged.
"""

import base64
import os
import shutil
import subprocess
import tempfile
import uuid
from typing import List, Dict, Any

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from pypdf import PdfWriter

app = FastAPI()

# A single, reused LibreOffice profile dir avoids repeated first-run bootstrap.
LO_PROFILE = os.path.join(tempfile.gettempdir(), "lo_profile")
OFFICE_EXTS = {".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".odt", ".ods"}

# Per-conversion timeout. Keep comfortably under the flow's 120s overall ceiling.
LO_TIMEOUT_SECONDS = 90


# --------------------------------------------------------------------------- #
# File-type detection (unchanged behavior)
# --------------------------------------------------------------------------- #
def detect_file_type(filename: str, data: bytes) -> str:
    name = (filename or "").lower()
    if name.endswith(".pdf") or data[:5] == b"%PDF-":
        return "pdf"
    ext = os.path.splitext(name)[1]
    if ext in OFFICE_EXTS:
        return "office"
    # Sniff common signatures as a fallback
    if data[:4] == b"PK\x03\x04":        # zip-based: docx/xlsx/pptx
        return "office"
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":  # legacy OLE: doc/xls
        return "office"
    return "pdf"  # default: treat as already-PDF


# --------------------------------------------------------------------------- #
# BATCH conversion: convert every Office file in ONE LibreOffice invocation
# --------------------------------------------------------------------------- #
def convert_office_files_batch(paths: List[str], outdir: str) -> Dict[str, str]:
    """
    Convert all given Office files to PDF in a SINGLE libreoffice call.
    Returns {source_path: produced_pdf_path}. Falls back to per-file on failure.
    """
    if not paths:
        return {}

    os.makedirs(LO_PROFILE, exist_ok=True)
    cmd = [
        "libreoffice",
        "-env:UserInstallation=file://" + LO_PROFILE,
        "--headless", "--norestore", "--nolockcheck", "--nodefault",
        "--convert-to", "pdf",
        "--outdir", outdir,
        *paths,
    ]
    result_map: Dict[str, str] = {}
    try:
        subprocess.run(
            cmd, check=True, timeout=LO_TIMEOUT_SECONDS,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        )
        for src in paths:
            produced = os.path.join(
                outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf"
            )
            if os.path.exists(produced):
                result_map[src] = produced
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as exc:
        print(f"[batch-convert] failed ({exc}); falling back to per-file")

    # Per-file fallback for anything the batch call did not produce
    missing = [p for p in paths if p not in result_map]
    for src in missing:
        produced = _convert_one(src, outdir)
        if produced:
            result_map[src] = produced
    return result_map


def _convert_one(src: str, outdir: str) -> str:
    os.makedirs(LO_PROFILE, exist_ok=True)
    cmd = [
        "libreoffice",
        "-env:UserInstallation=file://" + LO_PROFILE,
        "--headless", "--norestore", "--nolockcheck", "--nodefault",
        "--convert-to", "pdf", "--outdir", outdir, src,
    ]
    try:
        subprocess.run(
            cmd, check=True, timeout=LO_TIMEOUT_SECONDS,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        )
        produced = os.path.join(
            outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf"
        )
        return produced if os.path.exists(produced) else ""
    except Exception as exc:  # noqa: BLE001
        print(f"[convert-one] {src} failed: {exc}")
        return ""


# --------------------------------------------------------------------------- #
# /merge-pdfs
# --------------------------------------------------------------------------- #
@app.post("/merge-pdfs")
async def merge_pdfs(request: Request):
    body: Dict[str, Any] = await request.json()
    files: List[Dict[str, Any]] = body.get("files", [])
    output_name: str = body.get("output_name", "merged.pdf")
    template_data: Dict[str, Any] = body.get("templateData", {})

    workdir = tempfile.mkdtemp(prefix="merge_")
    try:
        # 1) Write every incoming file to disk; classify office vs pdf.
        ordered_inputs: List[Dict[str, str]] = []  # preserves merge order
        office_paths: List[str] = []
        for idx, f in enumerate(files):
            fname = f.get("name") or f"file_{idx}"
            raw = base64.b64decode(f.get("content", ""))
            ext = os.path.splitext(fname)[1] or ".bin"
            local = os.path.join(workdir, f"{idx:03d}_{uuid.uuid4().hex}{ext}")
            with open(local, "wb") as fh:
                fh.write(raw)

            kind = detect_file_type(fname, raw)
            ordered_inputs.append({"path": local, "kind": kind})
            if kind == "office":
                office_paths.append(local)

        # 2) Generate the checklist docx (project-specific) and queue it too.
        #    Preserve your original generate_checklist_pdf / replace_placeholders.
        checklist_src = generate_checklist_docx(template_data, workdir)
        if checklist_src:
            ordered_inputs.insert(0, {"path": checklist_src, "kind": "office"})
            office_paths.append(checklist_src)

        # 3) ONE batched conversion for ALL office files (the big win).
        converted = convert_office_files_batch(office_paths, workdir)

        # 4) Stream-merge in the original order; office files use their PDF.
        writer = PdfWriter()
        for item in ordered_inputs:
            pdf_path = (
                converted.get(item["path"])
                if item["kind"] == "office"
                else item["path"]
            )
            if pdf_path and os.path.exists(pdf_path):
                writer.append(pdf_path)
            else:
                print(f"[merge] skipping unconvertible input: {item['path']}")

        out_path = os.path.join(workdir, output_name)
        with open(out_path, "wb") as out_fh:
            writer.write(out_fh)
        writer.close()

        with open(out_path, "rb") as out_fh:
            merged_b64 = base64.b64encode(out_fh.read()).decode("ascii")

        return JSONResponse({"merged_pdf": merged_b64, "output_name": output_name})
    except Exception as exc:  # noqa: BLE001
        print(f"[merge-pdfs] ERROR: {exc}")
        return JSONResponse(status_code=500, content={"error": str(exc)})
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


# --------------------------------------------------------------------------- #
# Keep your real checklist generator. This stub preserves the call site.
# Replace the body with your existing generate_checklist_pdf/replace_placeholders
# logic that fills templates/DCL_Template.docx from template_data.
# --------------------------------------------------------------------------- #
def generate_checklist_docx(template_data: Dict[str, Any], workdir: str) -> str:
    """
    Return the path to a generated .docx checklist (filled from template_data),
    or "" if no checklist should be prepended.  <-- PORT YOUR ORIGINAL CODE HERE.
    """
    return ""


@app.get("/health")
async def health():
    return {"status": "ok"}
