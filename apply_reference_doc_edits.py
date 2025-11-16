#!/usr/bin/env python3
"""
Extracts the contents of reference-doc.docx and syncs changes to the reference-doc/ directory.

This script:
1. Extracts reference-doc.docx to a temporary directory
2. Syncs changes to the reference-doc/ directory
3. Protects certain metadata files from being overwritten
4. Formats the XML files for version control
"""

import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


def get_relative_paths(base_dir: Path) -> set[Path]:
    """
    Get a set of relative file paths from a base directory.

    Args:
        base_dir: Path object representing the base directory

    Returns:
        Set of relative file paths as Path objects
    """
    base_dir = base_dir.resolve()

    return {
        file_path.relative_to(base_dir)
        for file_path in base_dir.rglob("*")
        if file_path.is_file()
    }


def main() -> None:
    """Main entry point."""
    # Set paths
    cwd = Path.cwd()
    docx_path = cwd / "reference-doc.docx"
    target_dir = cwd / "reference-doc"

    # Ensure the .docx file exists
    if not docx_path.exists():
        print("Error: The file reference-doc.docx does not exist.", file=sys.stderr)
        sys.exit(1)

    # Create temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir_path = Path(temp_dir)
        temp_extract_dir = temp_dir_path / "reference-doc-temp"

        # Extract DOCX (which is a ZIP file) to temp directory
        print(f"Extracting {docx_path} to temporary directory...")
        with zipfile.ZipFile(docx_path, "r") as zip_ref:
            zip_ref.extractall(temp_extract_dir)

        # Build sets of relative file paths
        source_files = get_relative_paths(temp_extract_dir)
        target_files = get_relative_paths(target_dir) if target_dir.exists() else set()

        # Delete files in target_dir that are not in source_files
        for rel_path in target_files - source_files:
            full_path = target_dir / rel_path
            print(
                f"Deleting {rel_path} since it is not present in the updated reference-doc.docx"
            )
            full_path.unlink()

            # Remove empty parent directories
            parent = full_path.parent
            while parent != target_dir:
                try:
                    if any(parent.iterdir()):
                        break
                    parent.rmdir()
                    parent = parent.parent
                except OSError:
                    break

        # Protected files that contain non-functional changes on every save
        protected_files = {
            Path("docProps/app.xml"),
            Path("docProps/core.xml"),
            Path("word/settings.xml"),
            Path("word/glossary/settings.xml"),
        }

        # Copy new and updated files from source to target
        for rel_path in source_files:
            source_path = temp_extract_dir / rel_path
            target_path = target_dir / rel_path

            # Skip protected files if they already exist in target
            if rel_path in protected_files and target_path.exists():
                print(
                    f"Skipping {rel_path} as it likely contains only non-functional changes"
                )
                continue

            # Ensure target directory exists
            target_path.parent.mkdir(parents=True, exist_ok=True)

            # Copy file
            print(f"Copying {rel_path}")
            shutil.copy2(source_path, target_path)

    # Format XML files
    print("Formatting XML files...")
    result = subprocess.run(
        [sys.executable, "format_openxml.py"],
        capture_output=True,
        text=True,
        check=False,
    )

    if result.returncode != 0:
        print(f"Error formatting XML files: {result.stderr}", file=sys.stderr)
        sys.exit(1)

    print("Successfully extracted contents of reference-doc.docx into reference-doc/")


if __name__ == "__main__":
    main()
