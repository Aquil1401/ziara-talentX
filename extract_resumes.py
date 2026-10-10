#!/usr/bin/env python3
import os
from pathlib import Path
import fitz

OCR_AVAILABLE = False
try:
    from rapidocr_onnxruntime import RapidOCR
    ocr_engine = RapidOCR()
    OCR_AVAILABLE = True
except ImportError:
    ocr_engine = None

def ocr_pdf(pdf_path: Path, max_pages: int = 3) -> str:
    if not OCR_AVAILABLE or not ocr_engine:
        return ""
    text = []
    doc = fitz.open(pdf_path)
    for i, page in enumerate(doc):
        if i >= max_pages:
            break
        pix = page.get_pixmap(dpi=120)
        img_bytes = pix.tobytes("png")
        result, _ = ocr_engine(img_bytes)
        if result:
            text.extend([line[1] for line in result])
    doc.close()
    return "\n".join(text)
