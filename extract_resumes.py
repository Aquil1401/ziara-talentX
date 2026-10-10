#!/usr/bin/env python3
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

def export_to_excel(df, output_file: str = "candidates_resume_data.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Candidate Directory"
    # base excel styling
    wb.save(output_file)
