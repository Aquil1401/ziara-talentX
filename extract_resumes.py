#!/usr/bin/env python3
from typing import List, Dict

def select_best_contacts(emails: List[str], phones: List[str], candidate_name: str) -> Dict[str, str]:
    primary_email = emails[0] if emails else "N/A"
    alt_emails = "; ".join(emails[1:]) if len(emails) > 1 else "None"
    primary_phone = phones[0] if phones else "N/A"
    alt_phones = "; ".join(phones[1:]) if len(phones) > 1 else "None"
    return {
        "Primary Email": primary_email,
        "Alternative Emails": alt_emails,
        "Primary Phone": primary_phone,
        "Alternative Phone": alt_phones
    }
