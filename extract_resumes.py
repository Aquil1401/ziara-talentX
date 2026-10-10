#!/usr/bin/env python3
"""
Ziara TalentX - Candidate Contact Extractor
Author: Md Aquil (https://github.com/Aquil1401)
Organization: Ziara TechQ Labs (https://www.ziaratechqlabs.in/)
"""

import os
import sys
from pathlib import Path
import fitz

try:
    import docx
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

def extract_text_from_docx(docx_path: Path) -> str:
    if not DOCX_AVAILABLE:
        return ""
    try:
        doc = docx.Document(docx_path)
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    except Exception:
        return ""
