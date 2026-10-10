#!/usr/bin/env python3
"""
Ziara TalentX - Candidate Contact Extractor
Author: Md Aquil (https://github.com/Aquil1401)
Organization: Ziara TechQ Labs (https://www.ziaratechqlabs.in/)
"""

import os
import sys
import argparse
from pathlib import Path
import fitz  # PyMuPDF

def extract_text_from_pdf(pdf_path: Path) -> str:
    """Extract text from native PDF."""
    text = ""
    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            text += page.get_text() + "\n"
        doc.close()
    except Exception as e:
        print(f"[!] Error reading {pdf_path}: {e}")
    return text
