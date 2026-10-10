#!/usr/bin/env python3
import pandas as pd
from typing import List, Dict

def export_to_csv(data: List[Dict], output_file: str = "candidates_resume_data.csv"):
    df = pd.DataFrame(data)
    df.to_csv(output_file, index=False, encoding="utf-8-sig")
    print(f"[OK] CSV exported to {output_file}")
