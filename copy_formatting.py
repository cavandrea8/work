import docx

# Function to copy formatting elements from one document to another

def copy_formatting(source_file, target_file):
    source_doc = docx.Document(source_file)
    target_doc = docx.Document(target_file)
    
    # Loop through each element in the source document
    for element in source_doc.element.body:
        # Copy formatting elements such as styles, headers, footers, etc.
        # Note: This is a simplified example; actual implementation may vary.
        if element.tag.endswith('p'):  # Paragraph formatting
            new_paragraph = target_doc.add_paragraph()
            new_paragraph.style = source_doc.styles[element.style]
            new_paragraph.text = element.text
        elif element.tag.endswith('tbl'):  # Table formatting
            new_table = target_doc.add_table(rows=0, cols=0)
            # Copy each cell formatting
            for row in element.rows:
                new_row = new_table.add_row()
                for cell in row.cells:
                    new_cell = new_row.cells.add_cell()
                    new_cell.text = cell.text
                    new_cell.style = cell.style
    
    # Save the target document
    target_doc.save(target_file)

# Usage
copy_formatting('Manuale_SGI_Tresun.docx', 'LEG-SGI-01_Registro_Requisiti_Legali_Tresun_DEFINITIVO_FORMATTATO.docx')
