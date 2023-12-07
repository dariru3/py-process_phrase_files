from docx import Document
import os
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

def change_table_cell(cell, background_color=None):
    """Changes the background color of a table cell."""
    if background_color:
        shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), background_color))
        cell._tc.get_or_add_tcPr().append(shading_elm)

def process_word_file(file_path):
    doc = Document(file_path)

    # Delete the first three tables
    for _ in range(3):
        if len(doc.tables) > 0:
            tbl = doc.tables[0]
            tbl._element.getparent().remove(tbl._element)

    # Process the fourth table (now the first table in the document)
    if len(doc.tables) > 0:
        original_table = doc.tables[0]

        # Create a new table with 2 columns (for the 4th and 5th columns of the original table)
        new_table = doc.add_table(rows=0, cols=2) # UPDATE: 4 cols

        for row in original_table.rows:
            new_row = new_table.add_row()
            # Copy the content from the 4th and 5th columns of the original table; UPDATE: col 3-6
            new_row.cells[0].text = row.cells[3].text  # 4th column
            new_row.cells[1].text = row.cells[5].text  # 5th column

            # If the new 2nd column cell has text, set the new 1st column cell to gray; UPDATE: color both cells
            if new_row.cells[1].text.strip():
                change_table_cell(new_row.cells[0], background_color="D9D9D9")  # Gray color

        # Remove the original table
        original_table._element.getparent().remove(original_table._element)

    # Save the document
    new_file_path = file_path.replace('.docx', '_processed.docx')
    doc.save(new_file_path)

    print(f"Processed file saved as: {new_file_path}")

def process_all_word_files_in_folder(folder_path):
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.docx'):
            file_path = os.path.join(folder_path, file_name)
            process_word_file(file_path)

# Example usage
input_folder = 'input_files'
process_all_word_files_in_folder(input_folder)
