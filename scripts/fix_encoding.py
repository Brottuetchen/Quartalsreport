#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script to ensure all Markdown files are properly encoded in UTF-8 without BOM.
This script will re-save all .md files to fix any encoding issues.
"""

import sys
from pathlib import Path
from typing import List


def fix_file_encoding(file_path: Path) -> bool:
    """Read and re-save file to ensure proper UTF-8 encoding without BOM."""
    try:
        # Read the file
        with open(file_path, 'r', encoding='utf-8-sig') as f:  # utf-8-sig removes BOM if present
            content = f.read()

        # Write back without BOM
        with open(file_path, 'w', encoding='utf-8', newline='\n') as f:
            f.write(content)

        return True
    except Exception as e:
        print(f"✗ Fehler bei {file_path}: {e}")
        return False


def fix_all_markdown_files(root_dir: Path) -> List[Path]:
    """Fix encoding of all Markdown files in the project."""
    md_files = list(root_dir.glob('**/*.md'))
    exclude_dirs = {'.git', 'node_modules', 'venv', '.venv', 'data'}

    fixed_files = []

    for md_file in md_files:
        # Skip excluded directories
        if any(excluded in md_file.parts for excluded in exclude_dirs):
            continue

        print(f"Verarbeite: {md_file.relative_to(root_dir)}")

        if fix_file_encoding(md_file):
            fixed_files.append(md_file)
            print(f"  ✓ Erfolgreich korrigiert")
        else:
            print(f"  ✗ Fehler beim Korrigieren")

    return fixed_files


def main():
    """Main function."""
    project_root = Path(__file__).parent.parent
    print(f"Korrigiere Encoding aller Markdown-Dateien in: {project_root}\n")

    fixed_files = fix_all_markdown_files(project_root)

    print(f"\n✅ {len(fixed_files)} Dateien wurden erfolgreich verarbeitet:")
    for file_path in fixed_files:
        print(f"  - {file_path.relative_to(project_root)}")

    sys.exit(0)


if __name__ == "__main__":
    main()
