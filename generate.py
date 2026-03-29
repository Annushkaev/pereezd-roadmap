#!/usr/bin/env python3
"""Wrapper: adjusts paths and runs generate_dashboard.py for the standalone repo."""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent

# Patch paths before importing
import generate_roadmap
generate_roadmap.ROADMAP_DIR = ROOT / "source"

import generate_dashboard
generate_dashboard.ENTRY_PATH = ROOT / "data_entry.xlsx"
generate_dashboard.HTML_PATH = ROOT / "docs" / "index.html"

if __name__ == "__main__":
    generate_dashboard.main()
