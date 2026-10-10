#!/usr/bin/env python3
import re
from typing import List, Tuple

def score_email(email: str, candidate_folder_name: str, text: str) -> int:
    score = 0
    name_tokens = [t.lower() for t in re.split(r'[-_\s.]+', candidate_folder_name) if len(t) >= 3 and not t.isdigit()]
    email_local = email.split('@')[0].lower()
    for token in name_tokens:
        if token in email_local:
            score += 120
    return score
