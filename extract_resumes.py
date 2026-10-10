#!/usr/bin/env python3
import phonenumbers
from typing import List

def extract_phone_numbers(text: str, default_region: str = "IN") -> List[str]:
    phones = []
    try:
        for match in phonenumbers.PhoneNumberMatcher(text, default_region):
            formatted = phonenumbers.format_number(
                match.number, phonenumbers.PhoneNumberFormat.INTERNATIONAL
            )
            if formatted not in phones:
                phones.append(formatted)
    except Exception:
        pass
    return phones
