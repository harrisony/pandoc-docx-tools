#!/usr/bin/env python3
"""
Compiles a Markdown file to DOCX format using Pandoc with specific options.

This script:
1. Validates the input Markdown file
2. Runs Pandoc with configured options for numbered sections, TOC, citations, etc.
3. Opens the resulting DOCX file in the default application (e.g., Microsoft Word)
"""

import argparse
import platform
import subprocess
import sys
from pathlib import Path
from typing import NoReturn


# Constants
MARKDOWN_EXTENSION = ".md"
DOCX_EXTENSION = ".docx"
REFERENCE_DOC = "reference-doc.docx"
TEMPLATE_FILE = "template.openxml"


def open_file(file_path: Path) -> None:
    """
    Open a file with the default application for the current platform.

    Args:
        file_path: Path to the file to open

    Raises:
        FileNotFoundError: If the file doesn't exist
        subprocess.CalledProcessError: If the open command fails
    """
    if not file_path.exists():
        raise FileNotFoundError(f"Cannot open non-existent file: {file_path}")

    system = platform.system()

    try:
        if system == "Darwin":  # macOS
            subprocess.run(["open", str(file_path)], check=True)
        elif system == "Windows":
            # Use subprocess instead of os.startfile for consistency
            subprocess.run(["cmd", "/c", "start", "", str(file_path)], check=True)
        else:  # Linux and other Unix-like systems
            subprocess.run(["xdg-open", str(file_path)], check=True)
    except (subprocess.CalledProcessError, FileNotFoundError, OSError) as e:
        print(f"Warning: Could not open file: {e}", file=sys.stderr)


def validate_markdown_file(md_file: Path) -> None:
    """
    Validate that the markdown file exists and has the correct extension.

    Args:
        md_file: Path to the markdown file

    Raises:
        SystemExit: If validation fails
    """
    if not md_file.exists():
        error_exit(f"File '{md_file}' does not exist.")

    if md_file.suffix.lower() != MARKDOWN_EXTENSION:
        error_exit(f"File '{md_file}' is not a Markdown file ({MARKDOWN_EXTENSION} expected).")


def validate_reference_doc(reference_doc: Path) -> None:
    """
    Validate that the reference document exists.

    Args:
        reference_doc: Path to the reference document

    Raises:
        SystemExit: If validation fails
    """
    if not reference_doc.exists():
        error_exit(
            f"File '{reference_doc}' does not exist. "
            "Run './edit_reference_doc.py' and try again."
        )


def error_exit(message: str) -> NoReturn:
    """
    Print an error message and exit with status code 1.

    Args:
        message: Error message to print
    """
    print(f"Error: {message}", file=sys.stderr)
    sys.exit(1)


def build_pandoc_command(
    md_file: Path,
    docx_file: Path,
    resource_path: Path,
    reference_doc: Path,
    template: Path,
) -> list[str]:
    """
    Build the Pandoc command with all required options.

    Args:
        md_file: Input markdown file
        docx_file: Output DOCX file
        resource_path: Directory containing resources (images, etc.)
        reference_doc: Reference document for styling
        template: OpenXML template file

    Returns:
        List of command arguments for subprocess
    """
    return [
        "pandoc",
        f"--reference-doc={reference_doc}",
        f"--template={template}",
        "--number-sections",
        "--toc",
        "--lot",
        "--lof",
        "--citeproc",
        "--metadata=link-citations:true",
        "-t",
        "docx+native_numbering",
        f"--resource-path={resource_path}",
        str(md_file),
        "-o",
        str(docx_file),
    ]


def compile_markdown(md_file: Path) -> None:
    """
    Compile a markdown file to DOCX using Pandoc.

    Args:
        md_file: Path to the markdown file to compile

    Raises:
        SystemExit: If compilation fails
    """
    # Validate inputs
    validate_markdown_file(md_file)
    reference_doc = Path(REFERENCE_DOC)
    validate_reference_doc(reference_doc)

    # Prepare paths
    docx_file = md_file.with_suffix(DOCX_EXTENSION)
    resource_path = md_file.parent
    template = Path(TEMPLATE_FILE)

    # Build and run Pandoc command
    pandoc_cmd = build_pandoc_command(
        md_file, docx_file, resource_path, reference_doc, template
    )

    print(f"Compiling {md_file} to {docx_file}...")

    try:
        subprocess.run(pandoc_cmd, check=True)
        print(f"Successfully compiled to {docx_file}")
        open_file(docx_file)
    except subprocess.CalledProcessError:
        error_exit("Pandoc compilation failed.")
    except FileNotFoundError:
        error_exit("Pandoc is not installed or not found in PATH.")


def main() -> None:
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Compile a Markdown file to DOCX using Pandoc.",
        usage="%(prog)s <filename.md>",
    )
    parser.add_argument("md_file", help="Path to the Markdown file to compile")

    args = parser.parse_args()
    md_file = Path(args.md_file)

    compile_markdown(md_file)


if __name__ == "__main__":
    main()
