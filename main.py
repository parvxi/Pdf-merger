from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List
import base64
from io import BytesIO
from PyPDF2 import PdfMerger
import logging

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="DCL PDF Merger API",
    description="Merge multiple PDFs into one for Petrolube DCL system",
    version="1.0.0"
)

# Enable CORS for Power Automate
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class PDFFile(BaseModel):
    name: str
    content: str  # Base64 encoded PDF

class MergeRequest(BaseModel):
    files: List[PDFFile]
    output_name: str = "merged_dcl.pdf"

@app.get("/")
def health_check():
    """
    Health check endpoint
    """
    return {
        "status": "healthy",
        "service": "DCL PDF Merger",
        "version": "1.0.0",
        "endpoints": {
            "health": "/",
            "docs": "/docs",
            "merge": "/merge-pdfs"
        }
    }

@app.post("/merge-pdfs")
async def merge_pdfs(request: MergeRequest):
    """
    Merge multiple PDF files into one
    
    Request body:
    {
        "files": [
            {"name": "doc1.pdf", "content": "base64_encoded_content"},
            {"name": "doc2.pdf", "content": "base64_encoded_content"}
        ],
        "output_name": "merged.pdf"
    }
    
    Response:
    {
        "success": true,
        "output_name": "merged.pdf",
        "content": "base64_encoded_merged_pdf",
        "size_bytes": 12345,
        "files_merged": 2
    }
    """
    try:
        # Validate input
        if not request.files:
            raise HTTPException(
                status_code=400, 
                detail="No files provided. Please include at least one PDF file."
            )
        
        if len(request.files) > 100:
            raise HTTPException(
                status_code=400,
                detail="Too many files. Maximum 100 files allowed per merge."
            )
        
        logger.info(f"🔄 Starting merge of {len(request.files)} PDF files")
        
        # Create PDF merger
        merger = PdfMerger()
        
        # Add each PDF
        for idx, pdf_file in enumerate(request.files):
            try:
                logger.info(f"📄 Processing file {idx + 1}/{len(request.files)}: {pdf_file.name}")
                
                # Decode base64 to bytes
                try:
                    pdf_bytes = base64.b64decode(pdf_file.content)
                except Exception as e:
                    raise ValueError(f"Invalid base64 encoding in file: {pdf_file.name}")
                
                # Validate it's a PDF
                if not pdf_bytes.startswith(b'%PDF'):
                    raise ValueError(f"File {pdf_file.name} is not a valid PDF")
                
                # Create stream and add to merger
                pdf_stream = BytesIO(pdf_bytes)
                merger.append(pdf_stream)
                
                logger.info(f"✅ Added: {pdf_file.name} ({len(pdf_bytes)} bytes)")
                
            except ValueError as ve:
                logger.error(f"❌ Validation error for {pdf_file.name}: {str(ve)}")
                raise HTTPException(status_code=400, detail=str(ve))
            except Exception as e:
                logger.error(f"❌ Error processing {pdf_file.name}: {str(e)}")
                raise HTTPException(
                    status_code=400,
                    detail=f"Failed to process {pdf_file.name}: {str(e)}"
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
            "files_merged": len(request.files)
        }
        
        logger.info(f"✅ Successfully merged {len(request.files)} PDFs ({len(merged_bytes)} bytes)")
        
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