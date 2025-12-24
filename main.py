from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List
import base64
from io import BytesIO
from pypdf import PdfWriter
import logging
import openpyxl
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import tempfile
import os
import subprocess
import shutil
from PIL import Image
from docx import Document

# ==========================================
# Logging
# ==========================================
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ==========================================
# Check LibreOffice
# ==========================================
path = shutil.which("libreoffice")
if path:
    logger.info(f"✅ LibreOffice found at: {path}")
else:
    logger.warning("⚠ LibreOffice NOT FOUND — DOC/DOCX/XLS conversion will fail.")

# ==========================================
# FastAPI App
# ==========================================
app = FastAPI(
    title="DCL PDF Merger API",
    description="Convert ANY uploaded DCL files to PDF and merge",
    version="4.0.0"  # ← UPDATED VERSION
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    allow_credentials=True,
)

# ==========================================
# Pydantic Models (UPDATED)
# ==========================================
class DocumentFile(BaseModel):
    name: str
    content: str  # base64


class MergeRequest(BaseModel):
    files: List[DocumentFile]
    output_name: str = "merged_dcl.pdf"
    
    # ✅ NEW STRUCTURE - Matches extraction system
    templateData: dict | None = None  # All 67 fields in one object
    
    # ⚠️ DEPRECATED - Keep for backward compatibility only
    checklist: dict | None = None
    checklistMapped: dict | None = None
    templateMaster: dict | None = None


# ==========================================
# DOCX Placeholder Replacement
# ==========================================
def replace_placeholders(doc, data):
    
    def replace_in_paragraph(p):
        # Merge all runs into one text block
        full_text = "".join(run.text for run in p.runs)

        # Apply all replacements
        for key, value in data.items():
            full_text = full_text.replace("{{" + key + "}}", str(value))

        # Clear all runs and write final text
        for run in p.runs:
            run.text = ""
        if p.runs:
            p.runs[0].text = full_text
        else:
            p.add_run(full_text)

    def replace_in_table(table):
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p)
                for t2 in cell.tables:
                    replace_in_table(t2)

    # BODY paragraphs
    for p in doc.paragraphs:
        replace_in_paragraph(p)

    # BODY tables
    for t in doc.tables:
        replace_in_table(t)

    # FOOTER paragraphs
    for section in doc.sections:
        footer = section.footer
        for p in footer.paragraphs:
            replace_in_paragraph(p)
        for t in footer.tables:
            replace_in_table(t)

def generate_checklist_pdf(full_data: dict) -> bytes:
    template_path = "templates/DCL_Template.docx"

    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template missing: {template_path}")

    doc = Document(template_path)
    replace_placeholders(doc, full_data)

    temp_doc = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(temp_doc.name)

    pdf_bytes = convert_docx_to_pdf(open(temp_doc.name, "rb").read())
    os.remove(temp_doc.name)

    return pdf_bytes


@app.get("/debug-template")
def debug_template():
    path = "templates/DCL_Template.docx"
    return {
        "exists": os.path.exists(path),
        "absolute_path": os.path.abspath(path),
        "cwd": os.getcwd(),
        "dir_listing": os.listdir("templates") if os.path.exists("templates") else "templates folder missing"
    }


# ==========================================
# File Type Detection
# ==========================================
def detect_file_type(content_bytes: bytes, filename: str) -> str:
    if content_bytes.startswith(b"%PDF"):
        return "pdf"

    if content_bytes.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"

    if content_bytes[0:3] == b"\xff\xd8\xff":
        return "jpg"

    if content_bytes.startswith(b"GIF87a") or content_bytes.startswith(b"GIF89a"):
        return "gif"

    if content_bytes.startswith(b"RIFF") and content_bytes[8:12] == b"WEBP":
        return "webp"

    if content_bytes.startswith(b"PK\x03\x04"):
        header = content_bytes[:2000]
        if b"word/" in header:
            return "docx"
        if b"xl/" in header:
            return "xlsx"

    if content_bytes.startswith(b"\xd0\xcf\x11\xe0"):
        if filename.lower().endswith(".doc"):
            return "doc"
        if filename.lower().endswith(".xls"):
            return "xls"

    ext = filename.lower().split(".")[-1]
    if ext in ["pdf", "docx", "xlsx", "doc", "xls", "png", "jpg", "jpeg", "gif", "webp"]:
        return ext

    return "unknown"


# ==========================================
# Converters
# ==========================================
def convert_image_to_pdf(image_bytes: bytes) -> bytes:
    try:
        img = Image.open(BytesIO(image_bytes))

        if img.mode in ("RGBA", "LA"):
            bg = Image.new("RGB", img.size, (255, 255, 255))
            bg.paste(img, mask=img.split()[-1])
            img = bg
        else:
            img = img.convert("RGB")

        output = BytesIO()
        img.save(output, format="PDF")
        return output.getvalue()

    except Exception as e:
        raise ValueError(f"Failed to convert image: {e}")


def convert_docx_to_pdf(doc_bytes: bytes) -> bytes:
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp:
        temp.write(doc_bytes)
        doc_path = temp.name

    out_dir = tempfile.gettempdir()
    out_pdf = doc_path.replace(".docx", ".pdf")

    try:
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, doc_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=True,
        )

        with open(out_pdf, "rb") as f:
            return f.read()

    finally:
        if os.path.exists(doc_path):
            os.remove(doc_path)
        if os.path.exists(out_pdf):
            os.remove(out_pdf)


def convert_doc_to_pdf(doc_bytes: bytes) -> bytes:
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as temp:
        temp.write(doc_bytes)
        doc_path = temp.name

    out_dir = tempfile.gettempdir()
    out_pdf = doc_path.replace(".doc", ".pdf")

    try:
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, doc_path],
            check=True
        )
        with open(out_pdf, "rb") as f:
            return f.read()

    finally:
        if os.path.exists(doc_path):
            os.remove(doc_path)
        if os.path.exists(out_pdf):
            os.remove(out_pdf)


def convert_xlsx_to_pdf(xlsx_bytes: bytes) -> bytes:
    """Convert XLSX to PDF using LibreOffice (maintains all formatting)"""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp:
        temp.write(xlsx_bytes)
        xlsx_path = temp.name

    out_dir = tempfile.gettempdir()
    out_pdf = xlsx_path.replace(".xlsx", ".pdf")

    try:
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, xlsx_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=True,
            timeout=30  # Prevent hanging
        )

        if not os.path.exists(out_pdf):
            raise FileNotFoundError(f"LibreOffice failed to create PDF: {out_pdf}")

        with open(out_pdf, "rb") as f:
            return f.read()

    except subprocess.TimeoutExpired:
        raise ValueError("Excel conversion timed out (file too large or complex)")
    except subprocess.CalledProcessError as e:
        raise ValueError(f"LibreOffice conversion failed: {e.stderr.decode()}")
    finally:
        if os.path.exists(xlsx_path):
            os.remove(xlsx_path)
        if os.path.exists(out_pdf):
            os.remove(out_pdf)
            

def convert_xls_to_pdf(xls_bytes: bytes) -> bytes:
    """Convert XLS to PDF using LibreOffice"""
    with tempfile.NamedTemporaryFile(suffix=".xls", delete=False) as temp:
        temp.write(xls_bytes)
        xls_path = temp.name

    out_dir = tempfile.gettempdir()
    out_pdf = xls_path.replace(".xls", ".pdf")

    try:
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, xls_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=True,
            timeout=30
        )

        if not os.path.exists(out_pdf):
            raise FileNotFoundError(f"LibreOffice failed to create PDF: {out_pdf}")

        with open(out_pdf, "rb") as f:
            return f.read()

    except subprocess.TimeoutExpired:
        raise ValueError("Excel conversion timed out")
    except subprocess.CalledProcessError as e:
        raise ValueError(f"LibreOffice conversion failed: {e.stderr.decode()}")
    finally:
        if os.path.exists(xls_path):
            os.remove(xls_path)
        if os.path.exists(out_pdf):
            os.remove(out_pdf)

# ==========================================
# Merge Endpoint (UPDATED)
# ==========================================
@app.post("/merge-pdfs")
async def merge_pdfs(request: MergeRequest):
    if not request.files:
        raise HTTPException(400, "No files provided.")

    merger = PdfWriter()
    conversions = {}

    # ✅ UPDATED: Support both new and old structure
    if request.templateData:
        # NEW STRUCTURE (from extraction system)
        full_checklist_data = request.templateData
        logger.info("✅ Using NEW templateData structure (67 fields)")
    else:
        # OLD STRUCTURE (backward compatibility)
        header = request.checklist or {}
        activities = request.checklistMapped or {}
        template_master = request.templateMaster or {}
        
        full_checklist_data = {
            **header,
            **activities,
            **template_master
        }
        logger.info("⚠️ Using OLD structure (checklist + checklistMapped + templateMaster)")

    # 1️⃣ Generate and append checklist PDF
    if full_checklist_data:
        try:
            checklist_pdf = generate_checklist_pdf(full_checklist_data)
            merger.append(BytesIO(checklist_pdf))
            logging.info("✅ Checklist PDF added successfully.")
        except Exception as e:
            raise HTTPException(500, f"Checklist PDF generation failed: {e}")

    # 2️⃣ Convert and append documents
    for f in request.files:
        try:
            file_bytes = base64.b64decode(f.content)
        except:
            raise HTTPException(400, f"Invalid base64 for {f.name}")

        ftype = detect_file_type(file_bytes, f.name)

        if ftype == "pdf":
            output = file_bytes
            conversions[f.name] = "passed_pdf"

        elif ftype in ["png", "jpg", "jpeg", "gif", "webp"]:
            output = convert_image_to_pdf(file_bytes)
            conversions[f.name] = "image_to_pdf"

        elif ftype == "docx":
            output = convert_docx_to_pdf(file_bytes)
            conversions[f.name] = "docx_to_pdf"

        elif ftype == "doc":
            output = convert_doc_to_pdf(file_bytes)
            conversions[f.name] = "doc_to_pdf"

        elif ftype == "xlsx":
            output = convert_xlsx_to_pdf(file_bytes)
            conversions[f.name] = "xlsx_to_pdf_libreoffice"

        elif ftype == "xls":
            output = convert_xls_to_pdf(file_bytes)
            conversions[f.name] = "xls_to_pdf_libreoffice"

        else:
            raise HTTPException(400, f"Unsupported file type: {f.name}")

        merger.append(BytesIO(output))

    # 3️⃣ Final merge
    output_stream = BytesIO()
    merger.write(output_stream)
    merger.close()
    final_pdf = output_stream.getvalue()

    return {
        "success": True,
        "output_name": request.output_name,
        "size_bytes": len(final_pdf),
        "content": base64.b64encode(final_pdf).decode(),
        "files_merged": len(request.files),
        "conversions": conversions,
    }


@app.get("/health")
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
