#!/usr/bin/env python3
"""
Ziara TalentX - Candidate Contact Extractor
Author: Md Aquil (https://github.com/Aquil1401)
Organization: Ziara TechQ Labs (https://www.ziaratechqlabs.in/)
"""

import os
import sys
import argparse
from pathlib import Path

def main():
    parser = argparse.ArgumentParser(description="Candidate Contact Extractor")
    parser.add_argument("--folder", type=str, default=None)
    args = parser.parse_args()
    print("[*] Ziara TalentX initialized.")

if __name__ == "__main__":
    main()
