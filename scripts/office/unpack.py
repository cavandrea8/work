#!/usr/bin/env python3
"""Unpack a .docx file into its XML components for editing."""
import sys
import os
import zipfile
import shutil

def unpack(docx_path, output_folder):
    if not os.path.exists(docx_path):
        print(f"Error: {docx_path} not found")
        sys.exit(1)
    
    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)
    os.makedirs(output_folder)
    
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(output_folder)
    
    print(f"Unpacked {docx_path} to {output_folder}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python unpack.py <file.docx> <folder/>")
        sys.exit(1)
    unpack(sys.argv[1], sys.argv[2])
