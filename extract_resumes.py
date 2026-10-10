#!/usr/bin/env python3
from openpyxl.styles import PatternFill

FILL_VERIFIED = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
FILL_REVIEW = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
FILL_MISSING = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
BANNER_NOTE = "NOTE: Rows marked with (Yellow / Amber) contain multiple contact details. Please verify manually."
