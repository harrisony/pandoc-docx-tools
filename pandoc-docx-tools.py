#!/usr/bin/env python3
"""
Pandoc DOCX Tools - A unified tool for managing Pandoc DOCX reference documents.

OVERVIEW
========

This tool enables you to write content in Markdown (version-control friendly) while
maintaining corporate styling requirements in DOCX output. It separates content from
presentation, allowing proper version control while meeting formatting demands.

COMMANDS
========

compile <file.md>
    Compile Markdown files to DOCX format using Pandoc with configured options for
    numbered sections, table of contents, citations, and native Word numbering.

edit
    Create reference-doc.docx from the reference-doc/ directory.
    Use this after manually editing the unzipped DOCX structure.

apply
    Extract reference-doc.docx to the reference-doc/ directory.
    Use this to modify the reference document in Word, then sync changes back.

format [path]
    Format XML files for version control by removing superfluous attributes and
    elements that change on every save (rsid, paraId, textId, etc.).
    Default path: ./reference-doc


WORKFLOW
========

1. Initial Setup:
   - Create or obtain a reference-doc.docx with your desired styles
   - Run: ./pandoc-docx-tools.py apply
   - This extracts and formats the DOCX for version control

2. Editing Styles in Word:
   - Modify reference-doc.docx in Microsoft Word
   - Update styles via: Home → Styles pane → Filter → Modify
   - Run: ./pandoc-docx-tools.py apply
   - Commit the changes to version control

3. Direct XML Editing:
   - Edit files in the reference-doc/ directory
   - Run: ./pandoc-docx-tools.py edit
   - This packages the directory back into reference-doc.docx

4. Compiling Documents:
   - Write content in Markdown
   - Run: ./pandoc-docx-tools.py compile document.md
   - Output uses styles from reference-doc.docx


TECHNICAL DETAILS
=================

DOCX Structure:
    DOCX files are ZIP archives containing XML files and media. The smallest unit
    that can be styled is a "run" (<w:r>), containing text in <w:t> elements,
    wrapped in paragraphs (<w:p>).

Reference Document (--reference-doc):
    A template DOCX whose styles and document properties (margins, headers, footers)
    are applied to output. The content is ignored; only styles matter.

OOXML Template (--template):
    An XML template controlling document structure and content placement using
    Pandoc variables like $body$, $toc$, etc.

Compatibility Settings:
    The <w:compat> element in /word/settings.xml is critical. Its absence causes
    Word to treat tables as legacy format, misaligning them with left margins
    rather than borders.

Image Handling:
    Image dimensions use English Metric Units (EMUs):
    - 360,000 EMUs per centimeter
    - 914,400 EMUs per inch
    Pandoc 3.7+ allows reusing image relationship IDs from reference documents.

Protected Files:
    The 'apply' command protects certain metadata files from overwriting as they
    contain non-functional changes that occur on every Word save:
    - docProps/app.xml
    - docProps/core.xml
    - word/settings.xml
    - word/glossary/settings.xml


REFERENCES
==========

- Article: https://rnwest.engineer/auto-generate-docx-with-pandoc/
- OpenXML Explained: https://www.brandwares.com/downloads/Open-XML-Explained.pdf
- OpenXML Reference: http://officeopenxml.com/

Author: R. N. West (PowerShell original)
Converted to Python and unified by Claude Code
"""

import argparse
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import NoReturn, Optional


# Constants
MARKDOWN_EXTENSION = ".md"
DOCX_EXTENSION = ".docx"
REFERENCE_DOC = "reference-doc.docx"
REFERENCE_DIR = "reference-doc"
TEMPLATE_FILE = "template.openxml"


# =============================================================================
# Shared Utility Functions
# =============================================================================


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


def error_exit(message: str) -> NoReturn:
    """
    Print an error message and exit with status code 1.

    Args:
        message: Error message to print
    """
    print(f"Error: {message}", file=sys.stderr)
    sys.exit(1)


# =============================================================================
# Compile Command Functions
# =============================================================================


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
            f"Run './pandoc-docx-tools.py edit' and try again."
        )


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


# =============================================================================
# Edit Command Functions
# =============================================================================


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
        error_exit(f"Directory '{source_dir}' does not exist.")


def edit_reference_doc(source_dir: Optional[Path] = None) -> None:
    """Create DOCX from a directory.

    Args:
        source_dir: Path to the directory to package (default: reference-doc/)
                    The output DOCX will be named after the directory + .docx extension.
    """
    # Set paths relative to current directory
    cwd = Path.cwd()
    if source_dir is None:
        source_dir = cwd / REFERENCE_DIR

    # Derive docx filename from directory name
    dir_name = source_dir.name
    zip_path = cwd / f"{dir_name}.zip"
    docx_path = cwd / f"{dir_name}.docx"

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


# =============================================================================
# Apply Command Functions
# =============================================================================


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


def apply_reference_doc_edits(docx_path: Optional[Path] = None) -> None:
    """Extract a .docx file and sync changes to a directory of the same name.

    Args:
        docx_path: Path to the .docx file to extract (default: reference-doc.docx)
                   The target directory will be the filename without the .docx extension.
    """
    # Set paths
    cwd = Path.cwd()
    if docx_path is None:
        docx_path = cwd / REFERENCE_DOC

    # Derive folder name from filename (remove .docx extension)
    folder_name = docx_path.stem
    target_dir = cwd / folder_name

    # Ensure the .docx file exists
    if not docx_path.exists():
        error_exit(f"The file {docx_path} does not exist.")

    # Create temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir_path = Path(temp_dir)
        temp_extract_dir = temp_dir_path / f"{folder_name}-temp"

        # Extract DOCX (which is a ZIP file) to temp directory
        print(f"Extracting {docx_path.name} to temporary directory...")
        with zipfile.ZipFile(docx_path, "r") as zip_ref:
            zip_ref.extractall(temp_extract_dir)

        # Build sets of relative file paths
        source_files = get_relative_paths(temp_extract_dir)
        target_files = get_relative_paths(target_dir) if target_dir.exists() else set()

        # Delete files in target_dir that are not in source_files
        for rel_path in target_files - source_files:
            full_path = target_dir / rel_path
            print(
                f"Deleting {rel_path} since it is not present in the updated {docx_path.name}"
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

    # Format XML files by calling the format function directly
    print("Formatting XML files...")
    try:
        format_openxml(target_dir)
    except Exception as e:
        error_exit(f"Error formatting XML files: {e}")

    print(f"Successfully extracted contents of {docx_path.name} into {folder_name}/")


# =============================================================================
# Format Command Functions
# =============================================================================


def remove_attributes(node: ET.Element) -> None:
    """
    Remove superfluous attributes that change on every save.

    These attributes appear to be superfluous, change on every save, and should
    therefore be removed to make version-control easier.
    """
    SUPERFLUOUS_ATTRS = {
        "w14:paraId",
        "w14:textId",
        "wp14:editId",
        "wp14:anchorId",
        "w:storeItemID",
    }

    # Filter out superfluous attributes and rsid attributes in one pass
    node.attrib = {
        key: value
        for key, value in node.attrib.items()
        if key not in SUPERFLUOUS_ATTRS
        and not (key.startswith("w:rsid") or ("{" in key and "rsid" in key))
    }

    # Recursively process child nodes
    for child in node:
        remove_attributes(child)


def _extract_namespace(root: ET.Element) -> Optional[str]:
    """Extract namespace URI from root element tag."""
    if root.tag.startswith("{"):
        return root.tag[1 : root.tag.index("}")]
    return None


def remove_elements(root: ET.Element) -> None:
    """
    Remove superfluous elements that change on every save.

    These elements appear to be superfluous, change on every save, and should
    therefore be removed to make version-control easier.
    """
    SUPERFLUOUS_ELEMENTS = ["w:rsid", "w:id"]
    namespace_uri = _extract_namespace(root)

    # Build list of elements to remove with proper namespace handling
    for element_name in SUPERFLUOUS_ELEMENTS:
        if ":" in element_name:
            prefix, local_name = element_name.split(":", 1)
            tag_to_find = (
                f"{{{namespace_uri}}}{local_name}"
                if namespace_uri and prefix == "w"
                else local_name
            )
        else:
            tag_to_find = element_name

        # Find and remove elements
        for parent in root.iter():
            # Create list copy to avoid modification during iteration
            children_to_remove = [
                child
                for child in parent
                if child.tag == tag_to_find or child.tag.endswith(f"}}{local_name}")
            ]
            for child in children_to_remove:
                parent.remove(child)


def format_xml_with_namespaces(xml_string: str) -> str:
    """
    Format XML namespace declarations on the root element to be on new lines.
    """
    # Put the first xmlns declaration on a new line
    xml_string = re.sub(
        r'^(<[\w:]+) (xmlns(?::\w+)?="[^"]*")',
        r"\1\r\n  \2",
        xml_string,
        flags=re.MULTILINE,
    )

    # Split remaining xmlns declarations onto separate lines
    # This needs iteration due to overlapping patterns
    while True:
        new_string = re.sub(
            r'^  (xmlns(?::\w+)?="[^"]*") (xmlns(?::\w+)?="[^"]*")',
            r"  \1\r\n  \2",
            xml_string,
            flags=re.MULTILINE,
        )
        if new_string == xml_string:
            break
        xml_string = new_string

    # Put mc:Ignorable attributes on a new line
    xml_string = re.sub(
        r'^  (xmlns(?::\w+)?="[^"]*") (mc:Ignorable="[^"]*")',
        r"  \1\r\n  \2",
        xml_string,
        flags=re.MULTILINE,
    )

    return xml_string


def format_xml_file(file_path: Path) -> None:
    """Format a single XML or .rels file."""
    print(f"Formatting {file_path}")

    # Read and parse XML
    tree = ET.parse(file_path)
    root = tree.getroot()

    # Remove superfluous attributes and elements
    remove_attributes(root)
    remove_elements(root)

    # Convert to string with indentation
    ET.indent(tree, space="  ")
    xml_string = ET.tostring(root, encoding="unicode", xml_declaration=False)

    # Format XML namespace declarations on separate lines
    xml_string = format_xml_with_namespaces(xml_string)

    # Write formatted XML with UTF-8 encoding (no BOM) and trailing newline
    file_path.write_text(xml_string + "\r\n", encoding="utf-8", newline="")


def format_openxml(path: Path) -> None:
    """
    Format XML files in a directory for version control.

    Args:
        path: Path to the directory containing XML files
    """
    # Validate directory exists
    if not path.is_dir():
        error_exit(f"The directory '{path}' does not exist.")

    # Process all .xml and .rels files recursively
    xml_files = [*path.rglob("*.xml"), *path.rglob("*.rels")]

    for xml_file in xml_files:
        try:
            format_xml_file(xml_file)
        except Exception as e:
            error_exit(f"Error formatting {xml_file}: {e}")


# =============================================================================
# Command Handlers
# =============================================================================


def cmd_compile(args: argparse.Namespace) -> None:
    """Handle the compile subcommand."""
    md_file = Path(args.md_file)
    compile_markdown(md_file)


def cmd_edit(args: argparse.Namespace) -> None:
    """Handle the edit subcommand."""
    source_dir = Path(args.directory) if args.directory else None
    edit_reference_doc(source_dir)


def cmd_apply(args: argparse.Namespace) -> None:
    """Handle the apply subcommand."""
    docx_path = Path(args.docx) if args.docx else None
    apply_reference_doc_edits(docx_path)


def cmd_format(args: argparse.Namespace) -> None:
    """Handle the format subcommand."""
    path = Path(args.path)
    format_openxml(path)


# =============================================================================
# Main Entry Point
# =============================================================================


def main() -> None:
    """Main entry point with subcommand support."""
    description = """
Pandoc DOCX Tools - Manage Pandoc DOCX reference documents

Write content in Markdown (version-control friendly) while maintaining corporate
styling requirements in DOCX output. Separates content from presentation.
"""

    epilog = """
WORKFLOW
========

1. Initial Setup:
   - Create or obtain a reference-doc.docx with your desired styles
   - Run: %(prog)s apply
   - This extracts and formats the DOCX for version control

2. Editing Styles in Word:
   - Modify reference-doc.docx in Microsoft Word
   - Update styles via: Home → Styles pane → Filter → Modify
   - Run: %(prog)s apply
   - Commit the changes to version control

3. Direct XML Editing:
   - Edit files in the reference-doc/ directory
   - Run: %(prog)s edit
   - This packages the directory back into reference-doc.docx

4. Compiling Documents:
   - Write content in Markdown
   - Run: %(prog)s compile document.md
   - Output uses styles from reference-doc.docx


TECHNICAL DETAILS
=================

Reference Document (--reference-doc):
    A template DOCX whose styles and document properties (margins, headers, footers)
    are applied to output. The content is ignored; only styles matter.

OOXML Template (--template):
    An XML template controlling document structure and content placement using
    Pandoc variables like $body$, $toc$, etc.


REFERENCES
==========

Article: https://rnwest.engineer/auto-generate-docx-with-pandoc/
OpenXML Explained: https://www.brandwares.com/downloads/Open-XML-Explained.pdf
OpenXML Reference: http://officeopenxml.com/

Example: %(prog)s compile document.md
"""

    parser = argparse.ArgumentParser(
        description=description,
        epilog=epilog,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    subparsers = parser.add_subparsers(
        title="commands",
        description="Available commands",
        dest="command",
        required=True,
    )

    # Compile command
    compile_parser = subparsers.add_parser(
        "compile",
        help="Compile a Markdown file to DOCX using Pandoc",
        description="""
Compile Markdown to DOCX format using Pandoc with configured options including:
- Numbered sections (--number-sections)
- Table of contents (--toc)
- List of tables and figures (--lot, --lof)
- Citations with bibliography (--citeproc)
- Native Word numbering (docx+native_numbering)

The output uses styles from reference-doc.docx and structure from template.openxml.
The compiled DOCX file will be automatically opened in your default application.
        """,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    compile_parser.add_argument(
        "md_file",
        help="Path to the Markdown file to compile",
    )
    compile_parser.set_defaults(func=cmd_compile)

    # Edit command
    edit_parser = subparsers.add_parser(
        "edit",
        help="Create DOCX from a directory",
        description="""
Package a directory into a DOCX file.

Use this command after manually editing the unzipped DOCX structure in a directory.
The output DOCX filename is derived from the directory name.

For example, packaging 'my-template/' creates 'my-template.docx'.

The resulting DOCX will be opened in your default application.
        """,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    edit_parser.add_argument(
        "directory",
        nargs="?",
        default=None,
        help="Path to the directory to package (default: reference-doc/)",
    )
    edit_parser.set_defaults(func=cmd_edit)

    # Apply command
    apply_parser = subparsers.add_parser(
        "apply",
        help="Extract DOCX to a directory of the same name",
        description="""
Extract a DOCX file to a directory of the same name and format for version control.

Use this command to:
1. Extract a DOCX file after editing styles in Microsoft Word
2. Sync changes from the DOCX back to the directory structure
3. Prepare the files for committing to version control

The directory name is derived from the DOCX filename (without the .docx extension).
For example, 'my-template.docx' syncs to 'my-template/'.
""",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    apply_parser.add_argument(
        "docx",
        nargs="?",
        default=None,
        help="Path to the DOCX file to extract (default: reference-doc.docx)",
    )
    apply_parser.set_defaults(func=cmd_apply)

    # Format command
    format_parser = subparsers.add_parser(
        "format",
        help="Format XML files for version control",
        description="""
Format XML files in a directory to improve readability and version control.

This makes the XML files suitable for meaningful version control diffs.
        """,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    format_parser.add_argument(
        "path",
        nargs="?",
        default="./reference-doc",
        help="Path to the folder containing XML files (default: ./reference-doc)",
    )
    format_parser.set_defaults(func=cmd_format)

    # Parse arguments and call the appropriate function
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
