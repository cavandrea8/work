#!/usr/bin/env python3
"""Pack XML components back into a .docx file."""
import sys
import os
import zipfile
import shutil

def pack(input_folder, output_docx, original_docx=None):
    if not os.path.exists(input_folder):
        print(f"Error: {input_folder} not found")
        sys.exit(1)
    
    # If original is provided, we preserve its binary parts
    temp_zip = output_docx + ".tmp.zip"
    
    with zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(input_folder):
            for file in files:
                full_path = os.path.join(root, file)
                arcname = os.path.relpath(full_path, input_folder)
                zipf.write(full_path, arcname)
    
    os.rename(temp_zip, output_docx)
    print(f"Packed {input_folder} to {output_docx}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python pack.py <folder/> <output.docx> [--original <file.docx>]")
        sys.exit(1)
    
    input_folder = sys.argv[1]
    output_docx = sys.argv[2]
    original_docx = None
    
    if len(sys.argv) >= 5 and sys.argv[3] == "--original":
        original_docx = sys.argv[4]
    
    pack(input_folder, output_docx, original_docx)
