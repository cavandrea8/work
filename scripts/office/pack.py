#!/usr/bin/env python3
"""Pack XML components back into a .docx file."""
import sys
import zipfile
import os
import shutil

def pack_docx(source_folder, output_path, original_docx=None):
    if not os.path.exists(source_folder):
        print(f"Error: Folder {source_folder} not found")
        sys.exit(1)
    
    # Create the docx file (which is just a zip)
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for root, dirs, files in os.walk(source_folder):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, source_folder)
                zip_ref.write(file_path, arcname)
    
    print(f"Packed {source_folder} to {output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python pack.py <folder/> <output.docx> [--original <file.docx>]")
        sys.exit(1)
    
    source_folder = sys.argv[1]
    output_path = sys.argv[2]
    pack_docx(source_folder, output_path)
