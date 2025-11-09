from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List
import base64
from io import BytesIO
from PyPDF2 import PdfMerger
import logging
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import tempfile
import os
import subprocess
import shutil

# ✅ Setup logging first (before any logger calls)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ✅ Check LibreOffice presence and version
path = shutil.which("libreoffice")
if path:
    logger.info(f"✅ LibreOffice found at: {path}")
    try:
        version = subprocess.run(["libreoffice", "--version"], capture_output=True, text=True)
        logger.info(f"📦 LibreOffice version: {version.stdout.strip()}")
    except Exception as e:
        logger.warning(f"⚠️ Could not get LibreOffice version: {e}")
else:
    logger.warning("⚠️ LibreOffice not found. DOCX conversion will fail.")

# ✅ Initialize FastAPI app
app = FastAPI(
    title="DCL PDF Merger API with Auto-Conversion",
    description="Convert Word/Excel to PDF and merge all documents for Petrolube DCL system",
    version="2.0.0"
)


# Enable CORS for Power Automate
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class DocumentFile(BaseModel):
    name: str
    content: str  # Base64 encoded content (any format)

class MergeRequest(BaseModel):
    files: List[DocumentFile]
    output_name: str = "merged_dcl.pdf"

def detect_file_type(content_bytes: bytes, filename: str) -> str:
    """
    Detect file type from content and filename
    Returns: 'pdf', 'docx', 'doc', 'xlsx', 'xls', or 'unknown'
    """
    # Check by magic bytes
    if content_bytes.startswith(b'%PDF'):
        return 'pdf'
    elif content_bytes.startswith(b'PK\x03\x04'):  # ZIP-based (docx, xlsx)
        if filename.lower().endswith('.docx'):
            return 'docx'
        elif filename.lower().endswith('.xlsx'):
            return 'xlsx'
        else:
            # Try to detect by content
            try:
                # Check for Word document markers
                if b'word/' in content_bytes[:1000]:
                    return 'docx'
                elif b'xl/' in content_bytes[:1000]:
                    return 'xlsx'
            except:
                pass
    elif content_bytes.startswith(b'\xd0\xcf\x11\xe0'):  # Old Office format
        if filename.lower().endswith('.doc'):
            return 'doc'
        elif filename.lower().endswith('.xls'):
            return 'xls'
    
    # Fallback to extension
    ext = filename.lower().split('.')[-1]
    if ext in ['pdf', 'docx', 'doc', 'xlsx', 'xls']:
        return ext
    
    return 'unknown'

import subprocess
import tempfile
import os

def convert_docx_to_pdf(docx_bytes: bytes) -> bytes:
    """
    Convert DOCX → PDF using LibreOffice headless mode.
    Produces identical formatting to Microsoft Word export.
    Works cross-platform on Linux (Render), macOS, Windows.
    """
    # Save DOCX to a temporary file
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp_docx:
        temp_docx.write(docx_bytes)
        temp_docx_path = temp_docx.name

    # Define output path
    output_dir = tempfile.gettempdir()
    pdf_path = os.path.join(output_dir, os.path.basename(temp_docx_path).replace(".docx", ".pdf"))

    try:
        # Call LibreOffice to perform the conversion
        result = subprocess.run(
            [
                "libreoffice",
                "--headless",            # no GUI
                "--convert-to", "pdf",   # convert format
                "--outdir", output_dir,  # output folder
                temp_docx_path
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=True
        )

        if not os.path.exists(pdf_path):
            raise RuntimeError(f"LibreOffice failed: {result.stderr.decode('utf-8')}")

        # Read and return PDF bytes
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()

        return pdf_bytes

    finally:
        # Clean up temp files
        for path in [temp_docx_path, pdf_path]:
            if os.path.exists(path):
                os.remove(path)

def convert_xlsx_to_pdf(xlsx_bytes: bytes) -> bytes:
    """
    Convert XLSX to PDF using openpyxl and reportlab
    Simple conversion - creates PDF with cell contents
    """
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_xlsx:
        temp_xlsx.write(xlsx_bytes)
        temp_xlsx_path = temp_xlsx.name
    
    temp_pdf_path = temp_xlsx_path.replace('.xlsx', '.pdf')
    
    try:
        # Load workbook
        wb = openpyxl.load_workbook(temp_xlsx_path)
        
        # Create PDF
        c = canvas.Canvas(temp_pdf_path, pagesize=letter)
        width, height = letter
        
        for sheet in wb.worksheets:
            c.drawString(100, height - 50, f"Sheet: {sheet.title}")
            y = height - 100
            
            for row in sheet.iter_rows(max_row=50, max_col=10):  # Limit size
                row_data = [str(cell.value) if cell.value else '' for cell in row]
                text = ' | '.join(row_data)
                if text.strip():
                    c.drawString(50, y, text[:100])  # Truncate long text
                    y -= 20
                    if y < 50:
                        c.showPage()
                        y = height - 50
            
            c.showPage()
        
        c.save()
        
        # Read the PDF
        with open(temp_pdf_path, 'rb') as pdf_file:
            pdf_bytes = pdf_file.read()
        
        return pdf_bytes
    finally:
        # Cleanup
        if os.path.exists(temp_xlsx_path):
            os.remove(temp_xlsx_path)
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)

@app.get("/")
def health_check():
    """
    Health check endpoint
    """
    return {
        "status": "healthy",
        "service": "DCL PDF Merger with Auto-Conversion",
        "version": "2.0.0",
        "features": {
            "supported_formats": ["PDF", "DOCX", "XLSX"],
            "auto_conversion": True,
            "merge_pdfs": True
        },
        "endpoints": {
            "health": "/",
            "docs": "/docs",
            "merge": "/merge-pdfs"
        }
    }

@app.post("/merge-pdfs")
async def merge_pdfs(request: MergeRequest):
    """
    Convert documents to PDF (if needed) and merge into one PDF
    
    Supported formats:
    - PDF (passed through)
    - DOCX (converted to PDF)
    - XLSX (converted to PDF)
    
    Request body:
    {
        "files": [
            {"name": "invoice.docx", "content": "base64_encoded_content"},
            {"name": "packing_list.pdf", "content": "base64_encoded_content"},
            {"name": "data.xlsx", "content": "base64_encoded_content"}
        ],
        "output_name": "merged_dcl.pdf"
    }
    
    Response:
    {
        "success": true,
        "output_name": "merged_dcl.pdf",
        "content": "base64_encoded_merged_pdf",
        "size_bytes": 12345,
        "files_merged": 3,
        "conversions": {
            "invoice.docx": "converted",
            "packing_list.pdf": "passed_through",
            "data.xlsx": "converted"
        }
    }
    """
    try:
        # Validate input
        if not request.files:
            raise HTTPException(
                status_code=400, 
                detail="No files provided. Please include at least one file."
            )
        
        if len(request.files) > 50:
            raise HTTPException(
                status_code=400,
                detail="Too many files. Maximum 50 files allowed per merge."
            )
        
        logger.info(f"🔄 Starting merge of {len(request.files)} files")
        
        # Create PDF merger
        merger = PdfMerger()
        conversions = {}
        
        # Process each file
        for idx, doc_file in enumerate(request.files):
            try:
                logger.info(f"📄 Processing file {idx + 1}/{len(request.files)}: {doc_file.name}")
                
                # Decode base64 to bytes
                try:
                    file_bytes = base64.b64decode(doc_file.content)
                except Exception as e:
                    raise ValueError(f"Invalid base64 encoding in file: {doc_file.name}")
                
                # Detect file type
                file_type = detect_file_type(file_bytes, doc_file.name)
                logger.info(f"🔍 Detected type: {file_type} for {doc_file.name}")
                
                # Convert or pass through
                if file_type == 'pdf':
                    # Validate it's a valid PDF
                    if not file_bytes.startswith(b'%PDF'):
                        raise ValueError(f"File {doc_file.name} is not a valid PDF")
                    pdf_bytes = file_bytes
                    conversions[doc_file.name] = "passed_through"
                    
                elif file_type == 'docx':
                    logger.info(f"📝 Converting DOCX to PDF: {doc_file.name}")
                    pdf_bytes = convert_docx_to_pdf(file_bytes)
                    conversions[doc_file.name] = "converted_from_docx"
                    
                elif file_type == 'xlsx':
                    logger.info(f"📊 Converting XLSX to PDF: {doc_file.name}")
                    pdf_bytes = convert_xlsx_to_pdf(file_bytes)
                    conversions[doc_file.name] = "converted_from_xlsx"
                    
                else:
                    raise ValueError(
                        f"Unsupported file type for {doc_file.name}. "
                        f"Supported formats: PDF, DOCX, XLSX"
                    )
                
                # Add to merger
                pdf_stream = BytesIO(pdf_bytes)
                merger.append(pdf_stream)
                
                logger.info(f"✅ Added: {doc_file.name} ({len(pdf_bytes)} bytes)")
                
            except ValueError as ve:
                logger.error(f"❌ Validation error for {doc_file.name}: {str(ve)}")
                raise HTTPException(status_code=400, detail=str(ve))
            except Exception as e:
                logger.error(f"❌ Error processing {doc_file.name}: {str(e)}")
                raise HTTPException(
                    status_code=400,
                    detail=f"Failed to process {doc_file.name}: {str(e)}"
                )
        
        # Write merged PDF to bytes
        logger.info("🔗 Merging all PDFs...")
        output_stream = BytesIO()
        merger.write(output_stream)
        merger.close()
        
        # Get bytes and encode to base64
        output_stream.seek(0)
        merged_bytes = output_stream.read()
        merged_base64 = base64.b64encode(merged_bytes).decode('utf-8')
        
        result = {
            "success": True,
            "output_name": request.output_name,
            "content": merged_base64,
            "size_bytes": len(merged_bytes),
            "files_merged": len(request.files),
            "conversions": conversions
        }
        
        logger.info(f"✅ Successfully processed and merged {len(request.files)} files ({len(merged_bytes)} bytes)")
        
        return result
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"❌ Unexpected error during merge: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Internal server error: {str(e)}"
        )

@app.get("/health")
def health():
    """Simple health check"""
    return {"status": "ok"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)





