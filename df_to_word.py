from docx import Document
import os
import pandas as pd
from process_mxliff import parse_mxliff_to_df
import help_format_tables as help
from merge_df import merge_dfs
from process_word import process_word_file
from config_loader import CONFIG

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

def process_files(docx_file, mxliff_file, output_folder):
    # Process the Word and MXLIFF files
    df_word = process_word_file(os.path.join(input_folder, docx_file), output_folder)
    df_mxliff = parse_mxliff_to_df(os.path.join(input_folder, mxliff_file))

    # Merge the DataFrames
    merged_df = merge_dfs(df_word, df_mxliff)

    # Save the merged DataFrame to a Word document
    output_file_path = os.path.join(output_folder, f"{os.path.splitext(docx_file)[0]}_merged.docx")
    dataframe_to_word_table(merged_df, output_file_path)

def dataframe_to_word_table(df, output_file_path):
    doc = Document()
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.autofit = False

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
    
    help.format_table(table)
    help.apply_conditional_formatting(table)
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    doc.save(output_file_path)
    print(f"Merged tables saved as Word document: {output_file_path}.")

if __name__ == "__main__":
    g_settings = CONFIG["GeneralSettings"]
    input_folder = g_settings["InputFolderPath"] # "input_files/"
    output_folder = g_settings["OutputFolderPath"] # "output_files/"

    pairs = get_file_pairs(input_folder)
    for docx_file, mxliff_file in pairs:
        print(f"File pair:\n{docx_file}\n{mxliff_file}")
        process_files(docx_file, mxliff_file, output_folder)