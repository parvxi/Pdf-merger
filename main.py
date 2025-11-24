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
    version="3.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    allow_credentials=True,
)

# ==========================================
# Pydantic Models
# ==========================================
class DocumentFile(BaseModel):
    name: str
    content: str  # base64


class MergeRequest(BaseModel):
    files: List[DocumentFile]
    output_name: str = "merged_dcl.pdf"
    checklist: dict | None = None
    checklistMapped: dict | None = None


# ==========================================
# DOCX Placeholder Replacement
# ==========================================
def replace_placeholders(doc, data):

    def _replace_in_paragraphs(paragraphs):
        for p in paragraphs:
            for key, value in data.items():
                placeholder = "{{" + key + "}}"

                if placeholder in p.text:
                    # Fix split runs: rebuild full text
                    new_text = p.text.replace(placeholder, str(value))

                    # Remove all runs and set text once
                    for run in p.runs:
                        run.text = ""
                    p.runs[0].text = new_text

    def _replace_in_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    _replace_in_paragraphs(cell.paragraphs)

    # ----------------------------
    # BODY
    # ----------------------------
    _replace_in_paragraphs(doc.paragraphs)
    _replace_in_tables(doc.tables)

    # ----------------------------
    # FOOTER ONLY (header skipped)
    # ----------------------------
    for section in doc.sections:
        footer = section.footer

        _replace_in_paragraphs(footer.paragraphs)
        _replace_in_tables(footer.tables)


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
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_xlsx:
        temp_xlsx.write(xlsx_bytes)
        xlsx_path = temp_xlsx.name

    pdf_path = xlsx_path.replace(".xlsx", ".pdf")

    try:
        wb = openpyxl.load_workbook(xlsx_path)
        c = canvas.Canvas(pdf_path, pagesize=letter)
        width, height = letter

        for sheet in wb.worksheets:
            c.drawString(80, height - 50, f"Sheet: {sheet.title}")
            y = height - 100

            for row in sheet.iter_rows(max_row=40, max_col=10):
                row_data = [str(cell.value) if cell.value else "" for cell in row]
                c.drawString(40, y, " | ".join(row_data)[:150])
                y -= 20
                if y < 60:
                    c.showPage()
                    y = height - 60

            c.showPage()

        c.save()

        with open(pdf_path, "rb") as f:
            return f.read()

    finally:
        if os.path.exists(xlsx_path):
            os.remove(xlsx_path)
        if os.path.exists(pdf_path):
            os.remove(pdf_path)


# ==========================================
# Merge Endpoint
# ==========================================
@app.post("/merge-pdfs")
async def merge_pdfs(request: MergeRequest):
    if not request.files:
        raise HTTPException(400, "No files provided.")

    merger = PdfWriter()
    conversions = {}

    # 1️⃣ Merge checklist (header + activities)
    header = request.checklist or {}
    activities = request.checklistMapped or {}
    full_checklist_data = {**header, **activities}

    if full_checklist_data:
        try:
            checklist_pdf = generate_checklist_pdf(full_checklist_data)
            merger.append(BytesIO(checklist_pdf))
            logging.info("Checklist PDF added successfully.")
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
            conversions[f.name] = "xlsx_to_pdf"

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

