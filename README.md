# 🚀 Ziara TalentX

> **Automated Candidate Contact Extraction & Intelligence Engine for Recruiters, HR Teams & Institutions.**  
> Built & Maintained by [Md Aquil](https://github.com/Aquil1401) at [Ziara TechQ Labs](https://www.ziaratechqlabs.in/).  
> Extract verified candidate emails & global phone numbers from nested folders directly into clean Excel & CSV spreadsheets. Scanned PDF & OCR ready.

[![License: PolyForm Noncommercial](https://img.shields.io/badge/License-PolyForm%20Noncommercial-blue.svg)](LICENSE)
[![Organization: Ziara TechQ Labs](https://img.shields.io/badge/Organization-Ziara%20TechQ%20Labs-orange.svg)](https://www.ziaratechqlabs.in/)
[![Author: Md Aquil](https://img.shields.io/badge/Author-Md%20Aquil%20(Aquil1401)-purple.svg)](https://github.com/Aquil1401)
[![Email: ziaratechqlabs@gmail.com](https://img.shields.io/badge/Inquiries-ziaratechqlabs%40gmail.com-red.svg)](mailto:ziaratechqlabs@gmail.com)
[![Python: 3.10+](https://img.shields.io/badge/Python-3.10%2B-brightgreen.svg)](https://www.python.org/)
[![Architecture: Offline / Local Processing](https://img.shields.io/badge/Architecture-Offline%20%26%20Local%20Processing-blue.svg)]()

---

## 💼 Commercial Services & Enterprise Customization

Developed by **[Ziara TechQ Labs](https://www.ziaratechqlabs.in/)**, we provide customized recruitment automation, ATS integrations, and enterprise deployments for companies, colleges, and hiring agencies:

| Solution | Details | Action |
| :--- | :--- | :--- |
| 🏢 **Commercial License** | Deploy inside commercial platforms or proprietary agency workflows. | [Request License via Email](mailto:ziaratechqlabs@gmail.com) |
| 🖥️ **Custom Web / GUI App** | Turn this tool into a branded desktop application or private web dashboard. | [Contact Ziara TechQ Labs](mailto:ziaratechqlabs@gmail.com) |
| 📲 **WhatsApp / CRM Integration** | 1-click candidate WhatsApp reachout and direct database sync. | [Get Custom Build](https://www.ziaratechqlabs.in/) |
| 🧠 **AI Resume Scoring & Skills** | Automatically score and rank candidates against custom Job Descriptions. | [Schedule Demo](mailto:ziaratechqlabs@gmail.com) |

---

## ⚡ Quick Start: How to Run

No coding experience is required. Choose whichever method is easiest:

### Option 1: Drag & Drop (Recommended ⭐)
1. Find your folder containing candidate subfolders.
2. **Drag and drop that folder directly onto `run.bat`**.
3. Extraction starts immediately and generates your Excel spreadsheet.

---

### Option 2: Double-Click
1. **Double-click `run.bat`**.
2. Press **ENTER** to auto-detect candidate folders, or enter your custom folder path.
3. The results will be saved automatically.

---

## 📊 Output Spreadsheet Structure

Two structured files are generated upon completion:
1. **`candidates_resume_data.xlsx`** (Microsoft Excel file with colored headers and status tags)
2. **`candidates_resume_data.csv`** (Standard CSV file for ATS/database imports)

### Spreadsheet Columns:
* **Candidate Folder Name:** The exact folder identifier of the applicant.
* **Primary Email:** Candidate's personal email (smart-matched with candidate's name and contact headers).
* **Alternative Emails:** Secondary emails found (ensures no valid contacts are missed).
* **Primary Contact Number:** Standardized mobile/phone number formatted with country code (Supports India +91, US +1, UK +44, UAE +971, Europe, etc.).
* **Alternative Numbers:** Additional telephone/mobile numbers if multiple exist.
* **CV File Name:** Name of the detected resume file.
* **Extraction Method:** `Direct Text Extraction` or `OCR (Scanned Document)`.
* **Verification Status:**
  - `Verified (Single Contact)` ➔ Highlighted in **Soft Green**.
  - `⭐ Multiple Contacts (Please Verify)` ➔ Highlighted in **Soft Amber/Yellow** for manual review.
  - `CV Not Found` ➔ Highlighted in **Soft Peach/Orange**.

### 📌 Top Banner Notice in Excel:
At Row 1 of the generated spreadsheet, an executive note states:  
> *"📌 NOTE: Rows marked with ⭐ (Yellow / Amber) contain multiple contact details. Please verify manually. Rows in Green are clean single-contact matches."*

---

## 🛠️ One-Time Setup on a New Computer

If running on a new PC for the first time:

1. **Install Python:**
   - Download from [python.org/downloads](https://www.python.org/downloads/).
   - **IMPORTANT:** Check the box **"Add python.exe to PATH"** on the first installer screen.
2. **Launch `run.bat`:**
   - It will automatically install all required dependencies (`PyMuPDF`, `rapidocr-onnxruntime`, `phonenumbers`, `pandas`, `openpyxl`) in the background.

---

## 🧪 Testing with Sample Demo Data

This repository is designed with privacy best practices in mind:
* 📁 **Pre-configured .gitignore:** Common candidate directories (`candidate_sample/`, `candidates/`, `Sample/`) and document files (`*.pdf`, `*.docx`) are listed in `.gitignore` to help prevent unintentional local file commits to Git repositories.
* ⚡ **Auto Demo Creation:** If no candidate files are found when running, Ziara TalentX generates local synthetic demo candidates in `candidate_sample/` for an immediate test drive:
  - **Alex-Morgan_1001** (US format with Resume & Statement of Purpose)
  - **Rajesh-Sharma_1002** (Indian format with CV & Course Certificate)
  - **Elena-Rostova_1003** (UK format with Curriculum Vitae & Publications List)

To manually create demo test data at any time:
```bash
python extract_resumes.py --create-sample
```

---

## ⚖️ License & Legal Disclaimer

* **License:** This project is licensed under the **PolyForm Noncommercial License 1.0.0**. You may freely use, inspect, and modify this software for **personal, research, and non-commercial purposes**. Commercial use, hosting as a paid SaaS, or reselling is strictly prohibited without a commercial license. See [LICENSE](LICENSE).
* **Compliance & Ethics:** Please adhere to data privacy laws (DPDP Act 2023, GDPR, CCPA). Read [LEGAL.md](LEGAL.md).

---

## 📬 Contact & Commercial Licensing Approval

For commercial licensing grants, agency deployments, custom ATS integrations, or white-label solutions:
* ✉️ **Official Business Email:** [ziaratechqlabs@gmail.com](mailto:ziaratechqlabs@gmail.com)
* 🌐 **Company Website:** [Ziara TechQ Labs](https://www.ziaratechqlabs.in/)
* 💻 **GitHub Profile:** [@Aquil1401](https://github.com/Aquil1401)

### 📝 Commercial Authorization Process:
1. Send an email to **`ziaratechqlabs@gmail.com`** stating your business requirements and usage scale.
2. Ziara TechQ Labs will issue a formal **Commercial License Grant & Invoice**.
3. Upon approval, your organization receives full commercial rights to deploy and customize Ziara TalentX.
