from docx import Document
from table_to_df import table_to_df
import os
import re
from config_loader import CONFIG
p_settings = CONFIG["ProcessingSettings"]

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

def process_word_file(file_path, output_folder, attempts=1):
    final_col_length = len(CONFIG["GeneralSettings"]["Column_Headers"])
    if attempts == 1:
        print("Processing .DOCX file...")
    max_attempts = p_settings["MaxAttempts"]
    doc = Document(file_path)

    tables_to_delete = p_settings["DeleteFirstNTables"]
    delete_first_n_tables(doc=doc, n=tables_to_delete)

    columns_to_copy = adjust_columns_by_attempts(attempts)

    original_table = doc.tables[0]
    new_table = doc.add_table(rows=0, cols=final_col_length)

    copy_content_to_table(original_table, new_table, columns_to_copy)

    if validate_table_contents(new_table):
        df_table = table_to_df(new_table)
        return df_table
    else:
        if attempts < max_attempts:
            print(f'Attempt {attempts} failed, trying again...')
            return process_word_file(file_path, output_folder, attempts + 1)
        else:
            print(f'Maximum attempts reached for file {file_path}. File processing aborted.')
            return None

def contains_japanese(text):
    # Regular expression for matching Japanese characters
    pattern = p_settings["JapanesePattern"]
    return re.search(pattern, text) is not None

def validate_table_contents(new_table):
    valid_rows = True
    for row_number, row in enumerate(new_table.rows[1:11], start=2):
        column_3_target_text = row.cells[2].text

        if column_3_target_text and contains_japanese(column_3_target_text):
            valid_rows = False
    
    return valid_rows

def adjust_columns_by_attempts(attempts):
    attempt_1_col = p_settings["Mapping_1"]
    attempt_2_col = p_settings["Mapping_2"]
    attempts_mapping = {
        1: ("First attempt", attempt_1_col),
        2: ("Second attempt", attempt_2_col),
    }

    message, columns = attempts_mapping.get(attempts, ("Second attempt failed", None))
    print(message)
    return columns

# start of UNUSED
def save_as_word_file(file_path, output_folder, doc):
    # Construct new file path
    base_name = os.path.basename(file_path)
    name_part, extension = os.path.splitext(base_name)
    name_part = name_part.replace('_processed', '') # remove '_processed' if added by Excel process
    new_filename = name_part + '_processed' + extension
    processed_file_path = os.path.join(output_folder, new_filename)
    # Save the document
    doc.save(processed_file_path)
    print(f"Processed Word file saved as {processed_file_path}")
    return name_part

def save_as_csv_file(dataframe, filename):
    csv_file = f'output_files/{filename}.csv'
    dataframe.to_csv(csv_file, index=False)
    print(f"Word table saved as CSV file: {csv_file}.")

def process_all_word_files_in_folder(folder_path, output_folder):
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.docx'):
            file_path = os.path.join(folder_path, file_name)
            process_word_file(file_path, output_folder)

# end of UNUSED