#!/usr/bin/env python3
"""
Ziara TalentX - Automated Candidate Contact Extraction & Intelligence Engine
Author: Md Aquil (https://github.com/Aquil1401)
Organization: Ziara TechQ Labs (https://www.ziaratechqlabs.in/)
Extracts candidate information (Folder Name, Email, Contact Number, File Info)
and exports to CSV & Excel with smart verification scoring.
"""

__version__ = "1.0.0"
__author__ = "Md Aquil"

import os
import sys
import re
import argparse
from pathlib import Path
from typing import List, Dict, Tuple, Optional

# Ensure UTF-8 output in Windows console
if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Third party libraries
import fitz  # PyMuPDF
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Optional DOCX support
try:
    import docx  # type: ignore
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# Optional RapidOCR support for scanned PDF files
try:
    from rapidocr_onnxruntime import RapidOCR  # type: ignore
    rapid_ocr = RapidOCR()
    OCR_AVAILABLE = True
except Exception as e:
    rapid_ocr = None
    OCR_AVAILABLE = False

# Optional phonenumbers support for global international number validation
try:
    import phonenumbers  # type: ignore
    PHONENUMBERS_AVAILABLE = True
except ImportError:
    phonenumbers = None  # type: ignore
    PHONENUMBERS_AVAILABLE = False

# Keywords for detecting CV / Resume files in folder
CV_FILENAME_PATTERNS = [
    r'cv',
    r'resume',
    r'curriculum[\s_-]*vitae',
    r'curicuram[\s_-]*vite',
    r'biodata',
    r'bio[\s_-]*data',
    r'profile',
]

# Filename keywords to explicitly ignore (non-CV documents)
IGNORE_FILENAME_PATTERNS = [
    r'teaching[\s_-]*plan',
    r'research[\s_-]*statement',
    r'publication',
    r'cover[\s_-]*letter',
    r'statement[\s_-]*of[\s_-]*purpose',
    r'sop',
    r'lor',
    r'recommendation',
    r'degree',
    r'certificate',
    r'marksheet',
    r'transcript',
    r'blank',
    r'aadhar',
    r'pan[\s_-]*card',
    r'passport',
]

EMAIL_REGEX = re.compile(
    r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]*[a-zA-Z0-9]'
)

# Labeled phone pattern for capturing numbers next to phone keywords
LABEL_PHONE_REGEX = re.compile(
    r'(?:phone|mob(?:ile)?|cell|contact|tel|mb\.?no\.?|ph\.?no\.?)[\s:.-]*([+\d\s().-]{7,25})',
    re.IGNORECASE
)

# Regex fallback for international and local numbers
INTL_PHONE_REGEX = re.compile(
    r'(?:\+|00)\d{1,4}[\s.-]?(?:\(?\d{1,5}\)?[\s.-]?){1,4}\d{2,5}'
)
INDIAN_PHONE_REGEX = re.compile(
    r'(?:\+?91[\s.-]?)?[6-9]\d{4}[\s.-]?\d{5}|\b[6-9]\d{9}\b'
)


EMAIL_BLACKLIST_PREFIXES = (
    'info@', 'support@', 'editor@', 'admin@', 'contact@', 'help@', 'office@', 
    'registrar@', 'dean@', 'hod@', 'director@', 'hr@', 'careers@', 'jobs@', 
    'author-support@', 'no-reply@', 'noreply@', 'queries@', 'enquiry@',
    'sales@', 'feedback@', 'press@', 'library@', 'accounts@', 'billing@',
    'service@', 'inquiry@', 'admissions@', 'helpdesk@', 'conference@'
)

EMAIL_BLACKLIST_DOMAINS = (
    'elsevier.com', 'springer.com', 'ieee.org', 'nature.com', 'wiley.com', 
    'acm.org', 'mdpi.com', 'tandfonline.com', 'arxiv.org', 'researchgate.net',
    'sciencedirect.com', 'oup.com', 'cambridge.org', 'taylorandfrancis.com',
    'frontiersin.org', 'iop.org', 'cell.com', 'plos.org', 'biorxiv.org',
    'github.com', 'linkedin.com', 'google.com'
)


def extract_candidate_emails(text: str, candidate_folder_name: str = "") -> Tuple[str, str, bool]:
    """
    Intelligently extracts the candidate's primary and alternative emails.
    - Heavily prioritizes emails matching candidate's first/last name tokens from folder name.
    - Captures secondary emails separately so no data is lost.
    - Flags candidates with multiple emails for manual review.
    Returns: (primary_email, alternative_emails_str, has_multiple_emails)
    """
    matches = list(EMAIL_REGEX.finditer(text))
    if not matches:
        return "N/A", "N/A", False

    # Extract name tokens from folder name (e.g. Abhinav-Kumar-Pandey -> ['abhinav', 'kumar', 'pandey'])
    raw_tokens = re.split(r'[-_\s\d]+', candidate_folder_name)
    name_tokens = [tok.lower() for tok in raw_tokens if len(tok) >= 3]

    # Check contact header lines
    labeled_emails = set()
    for line in text.splitlines()[:40]:
        if re.search(r'\b(?:email|e-mail|mail|contact)\b', line, re.I):
            for em in EMAIL_REGEX.findall(line):
                labeled_emails.add(em.lower().strip('.,;:()[]{}<>'))

    scored = []
    seen = set()

    for m in matches:
        raw_email = m.group().strip('.,;:()[]{}<>')
        if '.' not in raw_email or raw_email.endswith(('png', 'jpg', 'jpeg', 'pdf')):
            continue

        low = raw_email.lower()
        if low in seen:
            continue
        seen.add(low)

        # Blacklist checks
        if any(low.startswith(p) for p in EMAIL_BLACKLIST_PREFIXES):
            continue
        domain = low.split('@')[-1]
        if any(domain == d or domain.endswith('.' + d) for d in EMAIL_BLACKLIST_DOMAINS):
            continue

        score = 0
        username = low.split('@')[0]

        # 1. Matches candidate's name tokens (Highest priority)
        matched_tokens = [tok for tok in name_tokens if tok in username]
        if matched_tokens:
            score += 120 * len(matched_tokens)

        # 2. Labeled explicitly in contact header
        if low in labeled_emails:
            score += 80

        # 3. Position: Top of document gets priority
        char_pos = m.start()
        if char_pos < 800:
            score += 50
        elif char_pos < 2000:
            score += 20
        elif char_pos > 3500:
            score -= 30

        # 4. Standard personal email domains
        if domain in ('gmail.com', 'yahoo.com', 'outlook.com', 'hotmail.com', 'icloud.com', 'rediffmail.com'):
            score += 30

        scored.append((score, raw_email))

    if not scored:
        return "N/A", "N/A", False

    # Sort descending by score
    scored.sort(key=lambda x: x[0], reverse=True)

    primary_email = scored[0][1]
    # Filter secondary emails (must have non-negative score and not be duplicate)
    secondary_emails = [email for score, email in scored[1:4] if score >= 0]
    
    has_multiple = len(secondary_emails) > 0
    secondary_str = ", ".join(secondary_emails) if secondary_emails else "N/A"

    return primary_email, secondary_str, has_multiple


def clean_phone_fallback(raw: str) -> Optional[str]:
    """Fallback cleaner when phonenumbers library is not available."""
    raw = raw.strip()
    digits = re.sub(r'\D', '', raw)
    
    if len(digits) < 8 or len(digits) > 15:
        return None
    if raw.startswith(('19', '20')) and len(digits) <= 8:
        return None
        
    if raw.startswith('+'):
        return raw
    if len(digits) == 10 and digits[0] in '6789':
        return f"+91 {digits[:5]} {digits[5:]}"
    if len(digits) == 12 and digits.startswith('91'):
        return f"+91 {digits[2:7]} {digits[7:]}"
    return raw


def extract_candidate_phones(text: str) -> Tuple[str, str, bool]:
    """
    Extracts primary and alternative contact numbers globally.
    Returns: (primary_phone, alternative_phones_str, has_multiple_phones)
    """
    found_phones = []
    seen = set()

    # METHOD A: Use Google's phonenumbers library if available
    if PHONENUMBERS_AVAILABLE:
        # 1. First Priority: Labeled phone numbers
        for match in LABEL_PHONE_REGEX.finditer(text):
            raw_candidate = match.group(1).strip()
            cleaned_candidate = re.sub(r'[^\d+()\- ]', '', raw_candidate).strip()
            if not cleaned_candidate:
                continue

            for region in ['IN', 'US', 'GB', None]:
                try:
                    parsed = phonenumbers.parse(cleaned_candidate, region)
                    if phonenumbers.is_valid_number(parsed):
                        fmt = phonenumbers.format_number(parsed, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
                        if fmt not in seen:
                            seen.add(fmt)
                            found_phones.append(fmt)
                        break
                except Exception:
                    pass

        # 2. General phone matcher
        try:
            for match in phonenumbers.PhoneNumberMatcher(text, 'IN'):
                parsed = match.number
                if phonenumbers.is_valid_number(parsed):
                    fmt = phonenumbers.format_number(parsed, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
                    if fmt not in seen:
                        seen.add(fmt)
                        found_phones.append(fmt)
        except Exception:
            pass

    # METHOD B: Advanced Regex Fallback
    if not found_phones:
        for match in LABEL_PHONE_REGEX.finditer(text):
            raw = match.group(1).strip()
            cleaned = clean_phone_fallback(raw)
            if cleaned and cleaned not in seen:
                seen.add(cleaned)
                found_phones.append(cleaned)

        for pattern in [INTL_PHONE_REGEX, INDIAN_PHONE_REGEX]:
            for match in pattern.findall(text):
                cleaned = clean_phone_fallback(match)
                if cleaned and cleaned not in seen:
                    seen.add(cleaned)
                    found_phones.append(cleaned)

    if not found_phones:
        return "N/A", "N/A", False

    primary_phone = found_phones[0]
    secondary_phones = found_phones[1:3]
    has_multiple = len(secondary_phones) > 0
    secondary_str = ", ".join(secondary_phones) if secondary_phones else "N/A"

    return primary_phone, secondary_str, has_multiple


def extract_text_from_pdf(pdf_path: Path) -> Tuple[str, bool]:
    """
    Extracts text from PDF.
    First tries native text extraction via PyMuPDF.
    If empty or very short (< 50 chars), falls back to RapidOCR with early stopping.
    Returns (extracted_text, is_ocr_used).
    """
    extracted_text = ""
    is_ocr = False

    try:
        doc = fitz.open(str(pdf_path))
        num_pages = len(doc)
        
        # 1. Fast Native text extraction (Page 1 and 2 cover 99.9% of resume headers)
        native_text_parts = []
        for i in range(min(num_pages, 2)):
            txt = doc[i].get_text().strip()
            if txt:
                native_text_parts.append(txt)
                # Check if we already have emails and phones
                combined_sample = "\n".join(native_text_parts)
                if EMAIL_REGEX.search(combined_sample) and (LABEL_PHONE_REGEX.search(combined_sample) or INTL_PHONE_REGEX.search(combined_sample)):
                    break
        
        extracted_text = "\n".join(native_text_parts)

        # 2. Only if native text is completely missing/empty, fallback to OCR
        if len(extracted_text.strip()) < 30:
            if OCR_AVAILABLE and rapid_ocr is not None:
                is_ocr = True
                ocr_text_parts = []
                # Check first 2 pages maximum
                for i in range(min(num_pages, 2)):
                    page = doc[i]
                    # dpi=120  # Optimized for speed and accuracy is optimal balance: 40% faster inference with perfect readability
                    pix = page.get_pixmap(dpi=120)  # Optimized for speed and accuracy
                    img_bytes = pix.tobytes("png")
                    result, _ = rapid_ocr(img_bytes)
                    if result:
                        page_text = "\n".join([line[1] for line in result])
                        ocr_text_parts.append(page_text)
                        
                        # Early break: If current page yielded email OR phone, stop OCR
                        current_all = "\n".join(ocr_text_parts)
                        if EMAIL_REGEX.search(current_all) or LABEL_PHONE_REGEX.search(current_all):
                            break
                            
                extracted_text = "\n".join(ocr_text_parts)
        
        doc.close()
    except Exception as e:
        print(f"    [!] Error reading PDF {pdf_path.name}: {e}", flush=True)

    return extracted_text, is_ocr


def extract_text_from_docx(docx_path: Path) -> str:
    """Extracts text from DOCX files."""
    if not DOCX_AVAILABLE:
        return ""
    text_parts = []
    try:
        doc = docx.Document(str(docx_path))
        for p in doc.paragraphs:
            if p.text.strip():
                text_parts.append(p.text.strip())
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        text_parts.append(cell.text.strip())
    except Exception as e:
        print(f"    [!] Error reading DOCX {docx_path.name}: {e}")
    return "\n".join(text_parts)


def is_cv_file(file_path: Path) -> int:
    """
    Scores how likely a file is to be a CV / Resume.
    Higher score means more likely. Negative or 0 means not a CV.
    """
    name_lower = file_path.name.lower()
    
    # Check ignore list first
    for pattern in IGNORE_FILENAME_PATTERNS:
        if re.search(pattern, name_lower):
            return -1

    # Check CV keywords
    score = 0
    if re.search(r'\bcv\b|[-_]cv[-_]|\.cv\.', name_lower):
        score += 10
    if 'cv' in name_lower:
        score += 5
    if 'resume' in name_lower:
        score += 10
    if 'curriculum' in name_lower or 'curicuram' in name_lower:
        score += 10
    if 'biodata' in name_lower or 'bio-data' in name_lower:
        score += 8
    if 'profile' in name_lower:
        score += 3

    return score


def find_cv_in_folder(candidate_path: Path) -> Optional[Path]:
    """
    Finds the CV/Resume file inside the candidate's folder or subfolders.
    """
    supported_extensions = ['.pdf', '.docx', '.doc']
    all_files = []
    
    # Recursively find all supported files
    for root, _, files in os.walk(candidate_path):
        for f in files:
            p = Path(root) / f
            if p.suffix.lower() in supported_extensions:
                all_files.append(p)

    if not all_files:
        return None

    # Score all files based on filename
    scored_files = []
    for f in all_files:
        score = is_cv_file(f)
        scored_files.append((score, f))

    # Sort descending by score
    scored_files.sort(key=lambda x: x[0], reverse=True)

    # If the top file has a positive score, select it
    if scored_files and scored_files[0][0] > 0:
        return scored_files[0][1]

    # Fallback: If no filename had CV keywords, check non-ignored files
    valid_fallback_files = [f for score, f in scored_files if score >= 0]
    if valid_fallback_files:
        # Check first page content of non-ignored files
        for f in valid_fallback_files:
            if f.suffix.lower() == '.pdf':
                try:
                    doc = fitz.open(str(f))
                    if len(doc) > 0:
                        first_page = doc[0].get_text().lower()
                        if any(kw in first_page for kw in ['curriculum vitae', 'resume', 'personal details', 'education', 'experience']):
                            doc.close()
                            return f
                    doc.close()
                except Exception:
                    pass

    return None


def process_candidates(sample_dir: Path) -> List[Dict]:
    """
    Processes all candidate folders under sample_dir.
    Candidate name is set STRICTLY to the Candidate Folder Name.
    """
    results = []
    
    # Identify candidate folders directly under sample_dir
    candidate_folders = [d for d in sample_dir.iterdir() if d.is_dir() and not d.name.startswith('.')]
    candidate_folders.sort(key=lambda d: d.name.lower())

    print(f"\n[+] Total Candidate Folders Found: {len(candidate_folders)}", flush=True)
    print(f"[+] OCR Engine Active: {'YES (RapidOCR - ONNX Runtime)' if OCR_AVAILABLE else 'NO'}", flush=True)
    print("=" * 70, flush=True)

    for idx, folder in enumerate(candidate_folders, 1):
        # Candidate Name is STRICTLY the folder name
        candidate_name = folder.name
        print(f"\n[{idx}/{len(candidate_folders)}] Processing: {candidate_name}", flush=True)

        cv_file = find_cv_in_folder(folder)

        if not cv_file:
            print("    [-] Status: CV File Not Found / Blank", flush=True)
            results.append({
                "Candidate Folder Name": candidate_name,
                "Primary Email": "N/A",
                "Alternative Emails": "N/A",
                "Primary Contact Number": "N/A",
                "Alternative Numbers": "N/A",
                "CV File Name": "Not Found",
                "Extraction Method": "N/A",
                "Verification Status": "CV Not Found"
            })
            continue

        print(f"    [OK] CV Found: {cv_file.name}", flush=True)
        
        # Extract text based on file type
        text = ""
        is_ocr = False
        if cv_file.suffix.lower() == '.pdf':
            text, is_ocr = extract_text_from_pdf(cv_file)
        elif cv_file.suffix.lower() == '.docx':
            text = extract_text_from_docx(cv_file)
            is_ocr = False

        # Extract primary and alternative contacts
        primary_email, alt_emails, has_mult_emails = extract_candidate_emails(text, candidate_folder_name=candidate_name)
        primary_phone, alt_phones, has_mult_phones = extract_candidate_phones(text)

        has_multiple_contacts = has_mult_emails or has_mult_phones

        if primary_email == "N/A" and primary_phone == "N/A":
            status_str = "No Contact Info Detected"
        elif has_multiple_contacts:
            status_str = "⭐ Multiple Contacts (Please Verify)"
        else:
            status_str = "Verified (Single Contact)"

        method_str = "OCR (Scanned Document)" if is_ocr else "Direct Text Extraction"

        print(f"    [*] Method: {method_str}", flush=True)
        print(f"    [*] Primary Email: {primary_email}", flush=True)
        if alt_emails != "N/A":
            print(f"    [*] Alt Emails: {alt_emails}", flush=True)
        print(f"    [*] Primary Phone: {primary_phone}", flush=True)
        if alt_phones != "N/A":
            print(f"    [*] Alt Phones: {alt_phones}", flush=True)
        print(f"    [*] Status: {status_str}", flush=True)

        results.append({
            "Candidate Folder Name": candidate_name,
            "Primary Email": primary_email,
            "Alternative Emails": alt_emails,
            "Primary Contact Number": primary_phone,
            "Alternative Numbers": alt_phones,
            "CV File Name": cv_file.name,
            "Extraction Method": method_str,
            "Verification Status": status_str
        })

    return results


def export_data(results: List[Dict], output_base_name: str = "candidates_resume_data"):
    """
    Exports results to CSV and professionally styled Excel (.xlsx)
    with top warning note and color-coded rows for easy verification.
    """
    df = pd.DataFrame(results)

    # 1. Export CSV (UTF-8 with BOM for Excel compatibility)
    csv_file = f"{output_base_name}.csv"
    df.to_csv(csv_file, index=False, encoding='utf-8-sig')
    print(f"\n    [OK] CSV Generated: {csv_file}")

    # 2. Export Styled Excel (.xlsx) with Row 1 Notice Banner
    excel_file = f"{output_base_name}.xlsx"
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        # Start table at row 3 (startrow=2) to leave row 1 for banner note
        df.to_excel(writer, index=False, sheet_name="Candidates", startrow=2)
        workbook = writer.book
        worksheet = writer.sheets["Candidates"]

        num_cols = len(df.columns)

        # Row 1: Merged Notice Banner in pure English
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
        banner_cell = worksheet.cell(row=1, column=1)
        banner_cell.value = "📌 NOTE: Rows marked with ⭐ (Yellow / Amber) contain multiple contact details. Please verify manually. Rows in Green are clean single-contact matches."
        banner_cell.font = Font(name="Calibri", size=11, bold=True, color="8C5300")
        banner_cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        banner_cell.alignment = Alignment(horizontal="center", vertical="center")
        worksheet.row_dimensions[1].height = 28
        worksheet.row_dimensions[2].height = 10  # blank spacer row

        # Table Header Styling at Row 3
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_align = Alignment(horizontal="left", vertical="center")
        border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        worksheet.row_dimensions[3].height = 25

        for col_num in range(1, num_cols + 1):
            cell = worksheet.cell(row=3, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        # Data Rows Styling starting at Row 4
        single_verified_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")  # Soft Green
        multi_contact_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")    # Soft Amber / Yellow
        not_found_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")        # Soft Peach / Orange

        for row_idx in range(len(df)):
            current_row = row_idx + 4
            worksheet.row_dimensions[current_row].height = 20
            status_val = str(df.iloc[row_idx]["Verification Status"])

            # Determine row fill color
            row_fill = None
            if "Multiple Contacts" in status_val or "⭐" in status_val:
                row_fill = multi_contact_fill
            elif "Verified" in status_val:
                row_fill = single_verified_fill
            elif "Not Found" in status_val or "No Contact" in status_val:
                row_fill = not_found_fill

            for col_num in range(1, num_cols + 1):
                cell = worksheet.cell(row=current_row, column=col_num)
                cell.border = border
                cell.alignment = left_align
                if row_fill:
                    cell.fill = row_fill

        # Auto-adjust column widths (using column index to handle merged banner safely)
        from openpyxl.utils import get_column_letter
        for col_idx, col in enumerate(worksheet.columns, 1):
            max_len = 0
            for cell in col[2:]:  # evaluate from row 3 downwards (ignoring merged banner)
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            col_letter = get_column_letter(col_idx)
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 16)

    print(f"    [OK] Excel Generated: {excel_file}")


def create_sample_candidates_dir(sample_dir_path: Path = Path("candidate_sample")) -> Path:
    """
    Generates synthetic dummy candidates for instant zero-setup demonstration.
    Kept strictly local and ignored by Git to help preserve candidate data privacy.
    """
    sample_dir_path.mkdir(exist_ok=True)
    samples = [
        {
            "folder": "Alex-Morgan_1001",
            "files": [
                ("Alex_Morgan_Resume.pdf", "Alex Morgan\nEmail: alex.morgan@gmail.com\nPhone: +1 415-555-2671\nLocation: San Francisco, CA\n\nExperience:\nSenior Software Engineer with 6 years experience in Python and cloud systems."),
                ("Statement_of_Purpose.pdf", "Statement of Purpose\nApplying for Principal Systems Architect role.")
            ]
        },
        {
            "folder": "Rajesh-Sharma_1002",
            "files": [
                ("Rajesh_Sharma_CV.pdf", "Curriculum Vitae\nRajesh Sharma\nEmail: rajesh.sharma@outlook.com\nMobile: +91 98765 43210\nAlternate: +91 91234 56789\nLocation: Bangalore, India\n\nFull Stack Developer specializing in Python, Django and React."),
                ("Certificate.pdf", "Certificate of Merit - Advanced AI and Cloud Architecture")
            ]
        },
        {
            "folder": "Elena-Rostova_1003",
            "files": [
                ("Elena_Rostova_CV.pdf", "Elena Rostova\nContact Details:\nEmail: elena.rostova@ac.uk\nTel: +44 20 7946 0912\nLondon, United Kingdom\n\nResearch Data Scientist & Machine Learning Specialist."),
                ("Publications_List.pdf", "List of Publications & Academic Presentations")
            ]
        }
    ]

    for item in samples:
        cand_folder = sample_dir_path / item["folder"]
        cand_folder.mkdir(exist_ok=True)
        for fname, text in item["files"]:
            fpath = cand_folder / fname
            if not fpath.exists():
                doc = fitz.open()
                page = doc.new_page()
                page.insert_text(fitz.Point(50, 72), text, fontsize=12)
                doc.save(str(fpath))
                doc.close()

    return sample_dir_path


def auto_detect_candidates_dir(given_path: Optional[str] = None) -> Path:
    """
    Intelligently finds the candidate folder:
    1. Uses user-specified path if provided.
    2. Checks standard folder names ('candidate_sample', 'Sample', 'Candidates', etc.).
    3. Finds any folder in current directory that contains candidate subfolders with resumes.
    4. Auto-generates local dummy 'candidate_sample' if no candidate data exists anywhere.
    """
    if given_path:
        p = Path(given_path)
        if p.exists() and p.is_dir():
            return p
        print(f"[!] Specified folder '{given_path}' not found. Auto-detecting...", flush=True)

    # 1. Common folder names
    for name in ['candidate_sample', 'Sample', 'sample', 'Candidates', 'candidates', 'Resumes', 'resumes', 'Applicants', 'applicants', 'Data', 'data']:
        p = Path(name)
        if p.exists() and p.is_dir():
            return p

    # 2. Scan subdirectories in current folder
    current_dir = Path(".")
    subdirs = [d for d in current_dir.iterdir() if d.is_dir() and not d.name.startswith(('.', '__', 'venv', 'env', 'private'))]
    
    for d in subdirs:
        # Check if this folder has candidate subdirectories
        child_dirs = [c for c in d.iterdir() if c.is_dir() and not c.name.startswith('.')]
        for c in child_dirs[:5]:
            for _, _, files in os.walk(c):
                if any(f.lower().endswith(('.pdf', '.docx', '.doc')) for f in files):
                    return d

    # 3. Check if current directory itself contains candidate subdirectories
    for d in subdirs:
        for _, _, files in os.walk(d):
            if any(f.lower().endswith(('.pdf', '.docx', '.doc')) for f in files):
                return current_dir

    # 4. If no candidate directories found at all, auto-generate synthetic demo sample
    print("[*] No candidate folders found in workspace.", flush=True)
    print("[*] Auto-generating demo sample candidates in 'candidate_sample/' for zero-setup testing...", flush=True)
    return create_sample_candidates_dir(Path("candidate_sample"))


def main():
    parser = argparse.ArgumentParser(description="Ziara TalentX - Automated Candidate Contact Extractor")
    parser.add_argument(
        "--folder",
        type=str,
        default=None,
        help="Path to the folder containing candidate folders (optional, auto-detects by default)"
    )
    parser.add_argument(
        "--create-sample",
        action="store_true",
        help="Generate synthetic demo candidate folder 'candidate_sample' for testing"
    )
    parser.add_argument(
        "--output",
        type=str,
        default="candidates_resume_data",
        help="Base name for output files (CSV and XLSX)"
    )
    args = parser.parse_args()

    if args.create_sample:
        p = create_sample_candidates_dir()
        print(f"[✓] Created demo sample candidates in: {p.resolve()}")
        return

    target_dir = auto_detect_candidates_dir(args.folder)

    if not target_dir or not target_dir.exists():
        print(f"[!] Error: Target directory '{target_dir}' does not exist.", flush=True)
        sys.exit(1)

    print("=" * 70)
    print("      AUTOMATED CV & CONTACT EXTRACTION SYSTEM")
    print("=" * 70)
    print(f"Target Directory: {target_dir.resolve()}")

    results = process_candidates(target_dir)

    if not results:
        print("[!] No candidate folders found.")
        return

    export_data(results, output_base_name=args.output)
    print("\n[DONE] All candidate data successfully extracted and exported.")
    print("=" * 70)


if __name__ == "__main__":
    main()
