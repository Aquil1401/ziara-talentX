#!/usr/bin/env python3
import re
from typing import List

EMAIL_REGEX = re.compile(
    r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+',
    re.IGNORECASE
)

def extract_emails(text: str) -> List[str]:
    raw_emails = EMAIL_REGEX.findall(text)
    cleaned = []
    for em in raw_emails:
        em = em.strip(".,;:()<>[]'\"").lower()
        if em and em not in cleaned:
            cleaned.append(em)
    return cleaned
