from docx import Document
from docx.shared import Mm
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import os

def change_cell_color(cells, background_color=None):
    for cell in cells:
        if background_color:
            shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), background_color))
            cell._tc.get_or_add_tcPr().append(shading_elm)

def delete_first_n_tables(doc, n):
    for _ in range(n):
        if len(doc.tables) > 0:
            table = doc.tables[0]
            table._element.getparent().remove(table._element)

def create_and_format_table(doc, row_widths, num_of_columns):
    new_table = doc.add_table(rows=0, cols=num_of_columns)
    new_table.style = 'Table Grid'
    for i, width in enumerate(row_widths):
        new_table.columns[i].width = Mm(width)
    return new_table

def copy_content_to_table(original_table, new_table, columns_to_copy):
    for row in original_table.rows:
        new_row = new_table.add_row()
        new_cells = new_row.cells
        for i, col_index in enumerate(columns_to_copy):
            new_cells[i].text = row.cells[col_index].text

def apply_conditional_formatting(new_table, condition_column_index, format_column_index, background_color):
    for row in new_table.rows:
        if row.cells[condition_column_index].text.strip():
            cells_to_color = [row.cells[format_column_index], row.cells[condition_column_index]]
            change_cell_color(cells_to_color, background_color)

def process_word_file(file_path, output_folder):
    doc = Document(file_path)

    delete_first_n_tables(doc=doc, n=3)

    row_widths = [9, 90, 110, 10]
    columns_to_copy = [2, 3, 5, 6]

    original_table = doc.tables[0]
    new_table = create_and_format_table(doc, row_widths, num_of_columns=4)
    copy_content_to_table(original_table, new_table, columns_to_copy)
    apply_conditional_formatting(new_table, condition_column_index=2, format_column_index=1, background_color="D9D9D9")

    # Remove the original table
    original_table._element.getparent().remove(original_table._element)

    # Construct new file path
    base_name = os.path.basename(file_path)
    name_part, extension = os.path.splitext(base_name)
    new_filename = name_part + '_processed' + extension
    processed_file_path = os.path.join(output_folder, new_filename)

    # Check for output folder
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Save the document
    doc.save(processed_file_path)

    print(f"Processed file saved as: {processed_file_path}")

def process_all_word_files_in_folder(folder_path, output_folder):
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.docx'):
            file_path = os.path.join(folder_path, file_name)
            process_word_file(file_path, output_folder)

# Example usage
process_all_word_files_in_folder('input_files', 'output_files')
