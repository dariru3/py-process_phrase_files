from docx import Document
import help_format_tables as help
import os

def delete_first_n_tables(doc, n):
    for _ in range(n):
        if len(doc.tables) > 0:
            table = doc.tables[0]
            table._element.getparent().remove(table._element)

def copy_content_to_table(original_table, new_table, columns_to_copy):
    for row in original_table.rows:
        new_row = new_table.add_row()
        new_cells = new_row.cells
        for i, col_index in enumerate(columns_to_copy):
            new_cells[i].text = row.cells[col_index].text

def process_word_file(file_path, output_folder):
    doc = Document(file_path)

    delete_first_n_tables(doc=doc, n=3)

    columns_to_copy = [2, 3, 4, 5, 6] # may need adjusting

    original_table = doc.tables[0]
    new_table = doc.add_table(rows=0, cols=5)
    help.format_table(new_table, comments=True)

    copy_content_to_table(original_table, new_table, columns_to_copy)
    help.apply_conditional_formatting(new_table)

    # Remove the original table
    original_table._element.getparent().remove(original_table._element)

    # Construct new file path
    base_name = os.path.basename(file_path)
    name_part, extension = os.path.splitext(base_name)
    name_part = name_part.replace('_processed', '') # remove '_processed' if added by Excel process
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
