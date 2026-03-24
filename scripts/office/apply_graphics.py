#!/usr/bin/env python3
"""
Apply complete graphic identity from Manuale_SGI_Tresun to LEG-SGI-01 document.
This script modifies XML files directly to ensure 360° graphic consistency.
"""

import os
import shutil
import re

# Paths
MANUALE_UNPACKED = "/workspace/temp/manuale_unpack"
LEGSGI_UNPACKED = "/workspace/temp/legsgi_unpack"
OUTPUT_FOLDER = "/workspace/outputs"

def copy_file(src, dst):
    """Copy a file from src to dst."""
    shutil.copy2(src, dst)
    print(f"Copied: {src} -> {dst}")

def apply_styles():
    """Copy styles.xml from Manuale to LEG-SGI."""
    copy_file(
        f"{MANUALE_UNPACKED}/word/styles.xml",
        f"{LEGSGI_UNPACKED}/word/styles.xml"
    )

def apply_theme():
    """Copy theme1.xml from Manuale to LEG-SGI."""
    os.makedirs(f"{LEGSGI_UNPACKED}/word/theme", exist_ok=True)
    copy_file(
        f"{MANUALE_UNPACKED}/word/theme/theme1.xml",
        f"{LEGSGI_UNPACKED}/word/theme/theme1.xml"
    )

def apply_headers_footers():
    """Copy header4.xml and footer4.xml from Manuale to LEG-SGI."""
    # Copy header
    if os.path.exists(f"{MANUALE_UNPACKED}/word/header4.xml"):
        copy_file(
            f"{MANUALE_UNPACKED}/word/header4.xml",
            f"{LEGSGI_UNPACKED}/word/header4.xml"
        )
    
    # Copy footer
    if os.path.exists(f"{MANUALE_UNPACKED}/word/footer4.xml"):
        copy_file(
            f"{MANUALE_UNPACKED}/word/footer4.xml",
            f"{LEGSGI_UNPACKED}/word/footer4.xml"
        )

def apply_media():
    """Copy all media files from Manuale to LEG-SGI."""
    os.makedirs(f"{LEGSGI_UNPACKED}/word/media", exist_ok=True)
    
    manuale_media = f"{MANUALE_UNPACKED}/word/media"
    legsgi_media = f"{LEGSGI_UNPACKED}/word/media"
    
    if os.path.exists(manuale_media):
        for filename in os.listdir(manuale_media):
            src = os.path.join(manuale_media, filename)
            dst = os.path.join(legsgi_media, filename)
            if os.path.isfile(src):
                copy_file(src, dst)

def replace_colors_in_xml(filepath):
    """Replace color codes in an XML file to match Manuale palette."""
    if not os.path.exists(filepath):
        return
    
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Color replacements (Manuale palette)
    color_replacements = {
        '#26211F': '#263238',  # dark text → dark blue-grey
        '#6B6361': '#546E7A',  # secondary text → blue-grey
        '#3D3735': '#37474F',  # tertiary text → slate
        '#C19A6B': '#E65100',  # accent/gold → deep orange
        '#F5F0EB': '#F5F5F5',  # light bg → light grey
        '#FDFCFB': '#FFFFFF',  # near-white → pure white
    }
    
    for old_color, new_color in color_replacements.items():
        # Replace in w:val attributes
        content = content.replace(f'w:val=\"{old_color[1:]}\"', f'w:val=\"{new_color[1:]}\"')
        # Replace in w:color attributes
        content = content.replace(f'w:color=\"{old_color[1:]}\"', f'w:color=\"{new_color[1:]}\"')
        # Replace in w:fill attributes
        content = content.replace(f'w:fill=\"{old_color[1:]}\"', f'w:fill=\"{new_color[1:]}\"')
    
    # Replace Times New Roman with Calibri
    content = content.replace('Times New Roman', 'Calibri')
    content = content.replace('w:ascii=\"Times New Roman\"', 'w:ascii=\"Calibri\"')
    content = content.replace('w:hAnsi=\"Times New Roman\"', 'w:hAnsi=\"Calibri\"')
    content = content.replace('w:eastAsia=\"Times New Roman\"', 'w:eastAsia=\"Calibri\"')
    content = content.replace('w:cs=\"Times New Roman\"', 'w:cs=\"Calibri\"')
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"Applied color/font replacements to: {filepath}")

def apply_color_replacements_throughout():
    """Apply color replacements to all XML files in LEG-SGI unpacked folder."""
    word_folder = f"{LEGSGI_UNPACKED}/word"
    
    xml_files = []
    for root, dirs, files in os.walk(word_folder):
        for file in files:
            if file.endswith('.xml'):
                xml_files.append(os.path.join(root, file))
    
    for filepath in xml_files:
        replace_colors_in_xml(filepath)

def main():
    print("=" * 60)
    print("APPLYING MANUALE SGI GRAPHICS TO LEG-SGI-01")
    print("=" * 60)
    
    # Step 1: Copy styles
    print("\n[STEP 1] Applying styles...")
    apply_styles()
    
    # Step 2: Copy theme
    print("\n[STEP 2] Applying theme...")
    apply_theme()
    
    # Step 3: Copy headers/footers
    print("\n[STEP 3] Applying headers and footers...")
    apply_headers_footers()
    
    # Step 4: Copy media/images
    print("\n[STEP 4] Applying media/images...")
    apply_media()
    
    # Step 5: Apply color replacements throughout
    print("\n[STEP 5] Applying color/font replacements...")
    apply_color_replacements_throughout()
    
    print("\n" + "=" * 60)
    print("GRAPHIC TRANSFORMATION COMPLETE")
    print("=" * 60)
    print(f"\nUnpacked folder ready: {LEGSGI_UNPACKED}")
    print("\nNext step: Run pack.py to create the final .docx file")

if __name__ == "__main__":
    main()
