from docx import Document
import os
import pandas as pd
from datetime import datetime
from .process_mxliff import parse_mxliff_to_df
from .format_helper import (
    format_table,
    apply_conditional_formatting,
    set_column_language,
    set_landscape_orientation,
    format_font_lines
)
from .merge_df import merge_dfs
from .process_word import process_word_file
from .save_formatting import reapply_formatting_to_column
from .config_loader import CONFIG

def delete_column_in_table(table, column_index):
    grid = table._tbl.find("w:tblGrid", table._tbl.nsmap)
    for cell in table.column_cells(column_index):
        cell._tc.getparent().remove(cell._tc)
    col_elem = grid[column_index]
    grid.remove(col_elem)

def get_file_pairs(folder_path):
    docx_files = {}
    mxliff_files = {}
    for filename in os.listdir(folder_path):
        base_name, ext = os.path.splitext(filename)
        if ext == ".docx":
            docx_files[base_name] = filename
        elif ext == ".mxliff":
            mxliff_files[base_name] = filename

    # Pairing files with the same base name
    pairs = []
    for base_name, docx_file in docx_files.items():
        mxliff_file = mxliff_files.get(base_name)
        if mxliff_file:
            pairs.append((docx_file, mxliff_file))
    return pairs

def dataframe_to_word_table(docx_file, df, output_folder, formatting_info):
    doc = Document()
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.autofit = False

    # Rename column headers
    df.rename(columns={'Index': 'p', 'Source': 'Japanese', 'Target': 'English'}, inplace=True)

    # Add header row
    for i, column in enumerate(df.columns):
        table.cell(0, i).text = str(column)

    # Add data rows
    for index, row in df.iterrows():
        cells = table.add_row().cells
        for i, value in enumerate(row):
            if pd.isnull(value) or value == "None":
                cells[i].text = ""
            else:
                cells[i].text = str(value)

    # TODO: combine table helpers and document helpers
    format_table(table)
    apply_conditional_formatting(table)
    set_column_language(table, 1, 'ja-JP')
    set_landscape_orientation(doc)
    format_font_lines(doc)

    # Reapply formatting to Enlglish text
    reapply_formatting_to_column(table=table, table_num=0, col_num=2, formatting_info=formatting_info)

    # Drop the 'Match' column after all formatting is done
    match_column_index = CONFIG["ConditionalFormattingSettings"]["MatchColumnIndex"]
    delete_column_in_table(table, match_column_index)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    output_file_path = os.path.join(output_folder, f"{os.path.splitext(docx_file)[0]}_merged.docx")
    doc.core_properties.created = datetime.now()
    doc.core_properties.modified = datetime.now()
    doc.save(output_file_path)
    print(f"Merged tables saved as Word document: {output_file_path}.")

def process_files(docx_file, mxliff_file, input_folder, output_folder):
    df_word, formatting_info = None, None
    # Process the Word and MXLIFF files
    processed_data = process_word_file(os.path.join(input_folder, docx_file), output_folder)
    if processed_data is not None:
        df_word, formatting_info = processed_data
    if df_word is not None and not df_word.empty:
        df_mxliff = parse_mxliff_to_df(os.path.join(input_folder, mxliff_file))
    else:
        print("Failed to process Word file.")
        return

    # Merge the DataFrames
    merged_df = merge_dfs(df_word, df_mxliff)

    # Save the merged DataFrame to a Word document
    dataframe_to_word_table(docx_file, merged_df, output_folder, formatting_info)

def print_debug(message_string, table):
    '''
    UNUSED
    '''
    print(f"\n========== {message_string} ==========")
    for i, row in enumerate(table.rows):
        if i in [6, 11]:
            print([cell.text for cell in row.cells])
