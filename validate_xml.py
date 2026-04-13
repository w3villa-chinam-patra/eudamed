#!/usr/bin/env python3
"""
XML Validator for EUDAMED XML files
Validates XML files against the XSD schemas
"""

import sys
import os
from pathlib import Path
from lxml import etree

def validate_xml(xml_file_path, xsd_file_path=None):
    """
    Validate an XML file against an XSD schema.
    
    Args:
        xml_file_path: Path to the XML file to validate
        xsd_file_path: Path to the XSD schema file (optional, will use default if not provided)
    """
    # Get the script directory (workspace root)
    script_dir = Path(__file__).parent.absolute()
    
    # Resolve XML file path
    xml_path = Path(xml_file_path)
    if not xml_path.is_absolute():
        xml_path = script_dir / xml_path
    
    if not xml_path.exists():
        print(f"✗ Error: XML file not found: {xml_path}")
        return False
    
    # Resolve XSD file path
    if xsd_file_path is None:
        # Default to Message.xsd in the XSD schemas directory
        xsd_path = script_dir / "XSD schemas" / "service" / "Message.xsd"
        print(f"XSD file path>>>>>>>>>>>>>>>.: {xsd_path}")

    else:
        print(f"XSD file path>>>>>>>>>>>>>>>.: {xsd_file_path}")
        xsd_path = Path(xsd_file_path)
        if not xsd_path.is_absolute():
            xsd_path = script_dir / xsd_path
    
    if not xsd_path.exists():
        print(f"✗ Error: XSD schema file not found: {xsd_path}")
        return False
    
    print(f"Validating: {xml_path.name}")
    print(f"Against schema: {xsd_path.relative_to(script_dir)}")
    print("-" * 60)
    
    try:
        # Change to XSD schemas directory to handle relative imports correctly
        original_cwd = os.getcwd()
        xsd_dir = xsd_path.parent.parent.parent  # Go up to "XSD schemas" directory
        
        try:
            os.chdir(xsd_dir)
            
            # Parse XSD schema
            print("Loading XSD schema...")
            xsd_doc = etree.parse(str(xsd_path.relative_to(xsd_dir)))
            xsd = etree.XMLSchema(xsd_doc)
            
            # Parse XML file (use absolute path)
            print("Loading XML file...")
            xml_doc = etree.parse(str(xml_path))
            
            # Validate
            print("Validating XML against schema...")
            is_valid = xsd.validate(xml_doc)
            
            if is_valid:
                print("\n" + "=" * 60)
                print("✓ SUCCESS: XML is valid!")
                print("=" * 60)
                return True
            else:
                print("\n" + "=" * 60)
                print("✗ FAILED: XML is invalid!")
                print("=" * 60)
                print("\nValidation errors:")
                print("-" * 60)
                for error in xsd.error_log:
                    print(f"  Line {error.line}, Column {error.column}:")
                    print(f"    {error.message}")
                    print()
                return False
                
        finally:
            os.chdir(original_cwd)
            
    except etree.XMLSyntaxError as e:
        print("\n" + "=" * 60)
        print("✗ FAILED: XML syntax error!")
        print("=" * 60)
        print(f"\nError: {e}")
        if hasattr(e, 'error_log'):
            for error in e.error_log:
                print(f"  Line {error.line}, Column {error.column}: {error.message}")
        return False
        
    except etree.XMLSchemaParseError as e:
        print("\n" + "=" * 60)
        print("✗ FAILED: XSD schema parsing error!")
        print("=" * 60)
        print(f"\nError: {e}")
        return False
        
    except Exception as e:
        print("\n" + "=" * 60)
        print("✗ FAILED: Unexpected error!")
        print("=" * 60)
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Main function to handle command line arguments."""
    if len(sys.argv) < 2:
        print("Usage: python validate_xml.py <xml_file> [xsd_file]")
        print("\nExamples:")
        print("  python validate_xml.py 'EOs - XML samples/SAMPLE_DTX_UDI_012.01.xml'")
        print("  python validate_xml.py 'EOs - XML samples/my_file.xml' 'XSD schemas/service/Message.xsd'")
        sys.exit(1)
    
    xml_file = sys.argv[1]
    xsd_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    success = validate_xml(xml_file, xsd_file)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()

