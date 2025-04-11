from docx import Document
from .table_to_df import table_to_df
from .save_formatting import extract_formatting_from_column
from .process_mxliff import remove_tags
from .config_loader import CONFIG

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
            original_text = row.cells[col_index].text
            # print(f"Col index: {col_index}, content: {original_text}")
            cleansed_text = remove_tags(original_text)
            new_cells[i].text = cleansed_text

def process_word_file(file_path, output_folder, attempts=1):
    p_settings = CONFIG["ProcessingSettings"]
    final_col_length = len(CONFIG["GeneralSettings"]["Column_Headers"])
    if attempts == 1:
        print("Processing .DOCX file...")
    doc = Document(file_path)

    # Save English text formatting
    formatting_info = extract_formatting_from_column(doc=doc, table_num=3, col_num=5)

    tables_to_delete = p_settings["DeleteFirstNTables"]
    delete_first_n_tables(doc=doc, n=tables_to_delete)

    columns_to_copy = [0, 3, 5, 6, 7]

    original_table = doc.tables[0]
    new_table = doc.add_table(rows=0, cols=final_col_length)

    copy_content_to_table(original_table, new_table, columns_to_copy)
    df_table = table_to_df(new_table)
    return df_table, formatting_info
