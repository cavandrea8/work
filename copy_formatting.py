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
            style_name = element.style
            if style_name and style_name in source_doc.styles:
                new_paragraph.style = source_doc.styles[style_name]
            new_paragraph.text = element.text
        elif element.tag.endswith('tbl'):  # Table formatting
            # Get the table object from the XML element
            table_obj = docx.table.Table(element, source_doc)
            new_table = target_doc.add_table(rows=len(table_obj.rows), cols=len(table_obj.columns))
            new_table.style = table_obj.style
            # Copy each cell formatting
            for row_idx, row in enumerate(table_obj.rows):
                for cell_idx, cell in enumerate(row.cells):
                    new_cell = new_table.cell(row_idx, cell_idx)
                    new_cell.text = cell.text
                    try:
                        new_cell.style = cell.style
                    except:
                        pass
    
    # Save the target document
    target_doc.save(target_file)

# Usage
copy_formatting('Manuale_SGI_Tresun.docx', 'LEG-SGI-01_Registro_Requisiti_Legali_Tresun_DEFINITIVO_FORMATTATO.docx')
