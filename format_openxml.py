#!/usr/bin/env python3
"""
Formats the contents of an unzipped DOCX to improve readability and enable effective version control.

This script recursively processes XML and .rels files in a given folder.
Each XML element is placed on a new line. The indent is set at two spaces.
XML namespace declarations on the root element are placed on separate lines.
Superfluous element attributes such as w14:paraId and w:rsidR are removed.

Author: R. N. West (PowerShell original)
Converted to Python by Claude Code

References:
- https://www.brandwares.com/downloads/Open-XML-Explained.pdf is a good introduction to OpenXML.
- See http://officeopenxml.com/ for an OpenXML element reference.
"""

import argparse
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional


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


def main() -> None:
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Formats the contents of an unzipped DOCX to improve readability and enable effective version control.",
        epilog="Example: %(prog)s ./reference-doc",
    )
    parser.add_argument(
        "path",
        nargs="?",
        default="./reference-doc",
        help="The path to the folder containing the XML files to format (default: ./reference-doc)",
    )

    args = parser.parse_args()
    path = Path(args.path)

    # Validate directory exists
    if not path.is_dir():
        print(
            f"Warning: The directory '{path}' does not exist. Exiting script.",
            file=sys.stderr,
        )
        sys.exit(1)

    # Process all .xml and .rels files recursively
    xml_files = [*path.rglob("*.xml"), *path.rglob("*.rels")]

    for xml_file in xml_files:
        try:
            format_xml_file(xml_file)
        except Exception as e:
            print(f"Error formatting {xml_file}: {e}", file=sys.stderr)
            sys.exit(1)


if __name__ == "__main__":
    main()
