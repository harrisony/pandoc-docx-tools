#!/usr/bin/env python3
"""
Creates a DOCX file from the reference-doc/ directory.

This script:
1. Creates a ZIP archive from the reference-doc/ directory
2. Applies selective compression (no compression for images, optimal for XML)
3. Renames the ZIP to DOCX
4. Opens the file in the default application (e.g., Microsoft Word)
"""

import os
import platform
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import NoReturn


def open_file(file_path: Path) -> None:
    """
    Open a file with the default application for the current platform.

    Args:
        file_path: Path to the file to open
    """
    system = platform.system()

    try:
        if system == "Darwin":  # macOS
            subprocess.run(["open", str(file_path)], check=True)
        elif system == "Windows":
            os.startfile(str(file_path))  # type: ignore[attr-defined]
        else:  # Linux and other Unix-like systems
            subprocess.run(["xdg-open", str(file_path)], check=True)
    except (subprocess.CalledProcessError, OSError) as e:
        print(f"Warning: Could not open file: {e}", file=sys.stderr)


def get_compression_type(entry_name: str) -> int:
    """
    Determine the compression type for a ZIP entry.

    Media files are stored uncompressed as they already have built-in compression.

    Args:
        entry_name: The ZIP entry name (path)

    Returns:
        ZIP_STORED for media files, ZIP_DEFLATED for others
    """
    return (
        zipfile.ZIP_STORED
        if entry_name.startswith("word/media/")
        else zipfile.ZIP_DEFLATED
    )


def cleanup_existing_files(*paths: Path) -> None:
    """
    Remove existing files if they exist.

    Args:
        *paths: Variable number of Path objects to remove
    """
    for path in paths:
        path.unlink(missing_ok=True)


def create_docx_archive(source_dir: Path, zip_path: Path) -> None:
    """
    Create a ZIP archive from the source directory with selective compression.

    Args:
        source_dir: Directory containing the document structure
        zip_path: Path where the ZIP file will be created
    """
    with zipfile.ZipFile(zip_path, "w") as zip_file:
        # Get all files recursively
        for file_path in source_dir.rglob("*"):
            if not file_path.is_file():
                continue

            # Get relative path in ZIP using Unix-style forward slashes (ZIP standard)
            relative_path = file_path.relative_to(source_dir)
            entry_name = relative_path.as_posix()

            # Determine compression type
            compression = get_compression_type(entry_name)

            # Add file to ZIP with appropriate compression
            print(f"Adding {entry_name}")
            zip_file.write(file_path, arcname=entry_name, compress_type=compression)


def validate_source_directory(source_dir: Path) -> None:
    """
    Validate that the source directory exists.

    Args:
        source_dir: Directory to validate

    Raises:
        SystemExit: If directory doesn't exist
    """
    if not source_dir.is_dir():
        print(f"Error: Directory '{source_dir}' does not exist.", file=sys.stderr)
        sys.exit(1)


def main() -> NoReturn:
    """Main entry point."""
    # Set paths relative to current directory
    cwd = Path.cwd()
    source_dir = cwd / "reference-doc"
    zip_path = cwd / "reference-doc.zip"
    docx_path = cwd / "reference-doc.docx"

    # Validate source directory exists
    validate_source_directory(source_dir)

    # Cleanup any existing files
    cleanup_existing_files(zip_path, docx_path)

    print(f"Creating DOCX from {source_dir}...")

    # Create the ZIP archive
    create_docx_archive(source_dir, zip_path)

    # Rename the .zip to .docx
    zip_path.rename(docx_path)

    print(f"Successfully created {docx_path}")

    # Open with default application
    open_file(docx_path)

    sys.exit(0)


if __name__ == "__main__":
    main()
