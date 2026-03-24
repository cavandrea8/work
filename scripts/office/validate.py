#!/usr/bin/env python3
"""Validate a .docx file structure."""
import sys
import os
import zipfile

def validate(docx_path):
    if not os.path.exists(docx_path):
        print(f"Error: {docx_path} not found")
        return False
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            # Check for required files
            required = ['[Content_Types].xml', 'word/document.xml', 
                       'word/_rels/document.xml.rels']
            for req in required:
                if req not in zip_ref.namelist():
                    print(f"Missing required file: {req}")
                    return False
            
            # Try to parse document.xml
            import xml.etree.ElementTree as ET
            doc_xml = zip_ref.read('word/document.xml')
            ET.fromstring(doc_xml)
            
            print(f"Validation PASSED: {docx_path}")
            return True
    except Exception as e:
        print(f"Validation FAILED: {e}")
        return False

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python validate.py <file.docx>")
        sys.exit(1)
    
    success = validate(sys.argv[1])
    sys.exit(0 if success else 1)
