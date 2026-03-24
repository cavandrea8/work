#!/usr/bin/env python3
"""
Script per copiare tutti gli stili da un documento Word a un altro.
Copia gli stili da Manuale_SGI_Tresun.docx a LEG-SGI-01_Registro_Requisiti_Legali_Tresun_DEFINITIVO_FORMATTATO.docx
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import shutil
import os

def copy_styles(source_docx, target_docx, backup=True):
    """
    Copia tutti gli stili dal documento source al documento target.
    
    Args:
        source_docx: percorso del documento sorgente (stili da copiare)
        target_docx: percorso del documento destinazione (dove copiare gli stili)
        backup: se True, crea un backup del file destinazione
    """
    
    # Verifica che i file esistano
    if not os.path.exists(source_docx):
        print(f"❌ Errore: File non trovato - {source_docx}")
        return False
    
    if not os.path.exists(target_docx):
        print(f"❌ Errore: File non trovato - {target_docx}")
        return False
    
    # Crea backup del file destinazione
    if backup:
        backup_path = target_docx.replace('.docx', '_BACKUP.docx')
        shutil.copy2(target_docx, backup_path)
        print(f"✓ Backup creato: {backup_path}")
    
    try:
        # Carica i documenti
        print(f"📂 Caricamento documento sorgente: {source_docx}")
        source_doc = Document(source_docx)
        
        print(f"📂 Caricamento documento destinazione: {target_docx}")
        target_doc = Document(target_docx)
        
        # Copia gli stili di paragrafo
        print("📋 Copia stili di paragrafo...")
        styles_copied = 0
        
        for source_style in source_doc.styles:
            style_name = source_style.name
            
            # Controlla se lo stile esiste già nel documento target
            try:
                target_style = target_doc.styles[style_name]
            except KeyError:
                # Lo stile non esiste, lo creeremo
                target_style = None
            
            # Copia le proprietà dello stile
            if source_style.type == 1:  # Stile di paragrafo
                if target_style:
                    # Aggiorna lo stile esistente
                    try:
                        if source_style.font.name:
                            target_style.font.name = source_style.font.name
                        if source_style.font.size:
                            target_style.font.size = source_style.font.size
                        if source_style.font.bold:
                            target_style.font.bold = source_style.font.bold
                        if source_style.font.italic:
                            target_style.font.italic = source_style.font.italic
                        if source_style.font.color.rgb:
                            target_style.font.color.rgb = source_style.font.color.rgb
                        
                        # Copia proprietà paragrafo
                        target_style.paragraph_format.alignment = source_style.paragraph_format.alignment
                        target_style.paragraph_format.left_indent = source_style.paragraph_format.left_indent
                        target_style.paragraph_format.right_indent = source_style.paragraph_format.right_indent
                        target_style.paragraph_format.first_line_indent = source_style.paragraph_format.first_line_indent
                        target_style.paragraph_format.space_before = source_style.paragraph_format.space_before
                        target_style.paragraph_format.space_after = source_style.paragraph_format.space_after
                        target_style.paragraph_format.line_spacing = source_style.paragraph_format.line_spacing
                        
                        print(f"  ✓ Stile aggiornato: {style_name}")
                        styles_copied += 1
                    except Exception as e:
                        print(f"  ⚠ Errore nell'aggiornamento dello stile '{style_name}': {str(e)}")
        
        # Salva il documento modificato
        print(f"\n💾 Salvataggio documento: {target_docx}")
        target_doc.save(target_docx)
        
        print(f"\n✅ Operazione completata!")
        print(f"   Stili copiati/aggiornati: {styles_copied}")
        print(f"   Documento salvato: {target_docx}")
        
        return True
        
    except Exception as e:
        print(f"\n❌ Errore durante l'operazione: {str(e)}")
        return False

if __name__ == "__main__":
    # Percorsi dei documenti
    source = "Manuale_SGI_Tresun.docx"
    target = "LEG-SGI-01_Registro_Requisiti_Legali_Tresun_DEFINITIVO_FORMATTATO.docx"
    
    print("=" * 60)
    print("COPIA STILI TRA DOCUMENTI WORD")
    print("=" * 60)
    print(f"\nSorgente: {source}")
    print(f"Destinazione: {target}\n")
    
    # Esegui la copia degli stili
    success = copy_styles(source, target, backup=True)
    
    if not success:
        print("\n⚠ Verifica che i file siano nella stessa directory dello script")
        print("   o modifica i percorsi nel file copy_styles.py")