from docx import Document
import help_format_tables as help
from table_to_df import table_to_df
import os
import re

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

    columns_to_copy = [2, 3, 5, 6, 7] # [2, 3, 4, 5, 6] # may need adjusting

    original_table = doc.tables[0]
    new_table = doc.add_table(rows=0, cols=5)
    help.format_table(new_table, comments=True)

    copy_content_to_table(original_table, new_table, columns_to_copy)
    help.apply_conditional_formatting(new_table)

    df_word_table = table_to_df(new_table)
    # csv_file = 'output_files/df_word_table.csv'
    # df_word_table.to_csv(csv_file, index=False)
    # print(f"CSV files has been saved to {csv_file}.")

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

    if validate_table_contents(new_table):
        # Save the document
        doc.save(processed_file_path)
        print(f"Processed file saved as: {processed_file_path}")
    else:
        print(f'At least one row does not meet the specified criteria in file {base_name}')

    return df_word_table

def contains_japanese(text):
    # Regular expression for matching Japanese characters
    pattern = r'[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]'
    return re.search(pattern, text) is not None

def validate_table_contents(new_table):
    valid_rows = True
    for row_number, row in enumerate(new_table.rows[1:9], start=2):  # Adjusted indexing for Python's 0-based index
        column_2_text = row.cells[1].text  # Column 2 (0-based index)
        column_3_text = row.cells[2].text  # Column 3 (0-based index)

        # Check conditions
        # if column_2_text and not contains_japanese(column_2_text):
        #     print(f"Row {row_number}, Column 2 does not meet the criteria.")
        #     print(column_2_text)
        #     valid_rows = False
        if column_3_text and contains_japanese(column_3_text):
            print(f"Row {row_number}, Column 3 does not meet the criteria.")
            print("Column 2")
            print(column_2_text)
            print("Column 3")
            print(column_3_text)
            valid_rows = False
    
    return valid_rows

# You can call validate_table_contents(new_table) after copying the content
# If the function returns False, it means at least one row does not meet the specified criteria.


def process_all_word_files_in_folder(folder_path, output_folder):
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.docx'):
            file_path = os.path.join(folder_path, file_name)
            process_word_file(file_path, output_folder)

# Example usage
process_all_word_files_in_folder('input_files', 'output_files')
