"""Cleanup script to remove temporary and test files."""

import shutil
from pathlib import Path


def cleanup_repo():
    """Remove temporary, test, and debug files from the repository."""

    repo_root = Path(__file__).parent
    files_removed = []
    dirs_removed = []

    print("Cleaning up repository...")
    print("=" * 60)

    # Test/Debug Python files
    test_patterns = [
        "test_*.py",
        "check_*.py",
        "validate_*.py",
        "analyze_*.py",
    ]

    for pattern in test_patterns:
        for file in repo_root.glob(pattern):
            if file.name != "cleanup.py":  # Don't delete this script
                print(f"  Removing: {file.name}")
                file.unlink()
                files_removed.append(file.name)

    # Temporary Excel files
    bsp_dir = repo_root / "BSP"
    if bsp_dir.exists():
        temp_excel_patterns = [
            "*_fixed.xlsm",
            "*_vorlauf.xlsm",
            "Test_*.xlsm",
            "~$*.xlsx",
            "~$*.xlsm",
        ]

        for pattern in temp_excel_patterns:
            for file in bsp_dir.glob(pattern):
                print(f"  Removing: BSP/{file.name}")
                file.unlink()
                files_removed.append(f"BSP/{file.name}")

    # Python cache directories
    for pycache in repo_root.rglob("__pycache__"):
        print(f"  Removing: {pycache.relative_to(repo_root)}")
        shutil.rmtree(pycache)
        dirs_removed.append(str(pycache.relative_to(repo_root)))

    # Temporary output files
    temp_files = [
        "excel_check_output.txt",
    ]

    for filename in temp_files:
        file = repo_root / filename
        if file.exists():
            print(f"  Removing: {filename}")
            file.unlink()
            files_removed.append(filename)

    # Windows-specific scripts (replaced by Docker)
    windows_scripts = [
        "run_portable.ps1",
        "run_portable.cmd",
    ]

    for filename in windows_scripts:
        file = repo_root / filename
        if file.exists():
            print(f"  Removing: {filename}")
            file.unlink()
            files_removed.append(filename)

    # Old app.js (replaced by app-flexible.js)
    old_app_js = repo_root / "webapp" / "static" / "app.js"
    if old_app_js.exists():
        # Check if it's different from app-flexible.js
        app_flexible = repo_root / "webapp" / "static" / "app-flexible.js"
        if app_flexible.exists():
            print(f"  Note: Keep old app.js for reference (now superseded by app-flexible.js)")
            # Don't delete it, just note it

    print("=" * 60)
    print(f"[OK] Cleanup complete!")
    print(f"   Files removed: {len(files_removed)}")
    print(f"   Directories removed: {len(dirs_removed)}")

    if not files_removed and not dirs_removed:
        print("   Repository was already clean!")

    print("\nRepository structure is now clean and organized.")


if __name__ == "__main__":
    cleanup_repo()
