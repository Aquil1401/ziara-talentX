#!/usr/bin/env python3
BLACKLIST_EMAIL_DOMAINS = {
    'springer.com', 'ieee.org', 'nature.com', 'elsevier.com',
    'sciencedirect.com', 'wiley.com', 'acm.org', 'researchgate.net',
    'arxiv.org', 'tandfonline.com', 'oup.com', 'mdpi.com'
}

BLACKLIST_PREFIXES = {
    'editor', 'editorial', 'info', 'support', 'contact', 'admin',
    'queries', 'hod', 'principal', 'office', 'admissions'
}

def is_blacklisted(email: str) -> bool:
    local_part, domain = email.split('@')[0].lower(), email.split('@')[1].lower()
    if domain in BLACKLIST_EMAIL_DOMAINS:
        return True
    if any(local_part.startswith(p) for p in BLACKLIST_PREFIXES):
        return True
    return False
