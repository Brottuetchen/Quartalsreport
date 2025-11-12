#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script to check and ensure UTF-8 encoding for all text files in the project.
"""

import sys
from pathlib import Path
from typing import List, Tuple


def check_file_encoding(file_path: Path) -> Tuple[bool, str]:
    """Check if a file is valid UTF-8."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            f.read()
        return True, "OK"
    except UnicodeDecodeError as e:
        return False, f"UTF-8 decode error: {e}"
    except Exception as e:
        return False, f"Error: {e}"


def check_project_encoding(root_dir: Path) -> List[Tuple[Path, str]]:
    """Check encoding of all relevant text files in the project."""
    extensions = {'.py', '.md', '.html', '.css', '.js', '.txt', '.json', '.yml', '.yaml'}
    exclude_dirs = {'__pycache__', '.git', 'node_modules', 'venv', '.venv', 'data'}

    issues = []

    for file_path in root_dir.rglob('*'):
        # Skip directories
        if file_path.is_dir():
            continue

        # Skip excluded directories
        if any(excluded in file_path.parts for excluded in exclude_dirs):
            continue

        # Check only text files
        if file_path.suffix in extensions:
            is_valid, message = check_file_encoding(file_path)
            if not is_valid:
                issues.append((file_path, message))
            else:
                print(f"✓ {file_path.relative_to(root_dir)}")

    return issues


def main():
    """Main function."""
    project_root = Path(__file__).parent.parent
    print(f"Checking encoding of text files in: {project_root}\n")

    issues = check_project_encoding(project_root)

    if issues:
        print("\n❌ Encoding issues found:")
        for file_path, message in issues:
            print(f"  {file_path.relative_to(project_root)}: {message}")
        sys.exit(1)
    else:
        print("\n✅ All text files are properly encoded in UTF-8!")
        sys.exit(0)


if __name__ == "__main__":
    main()
