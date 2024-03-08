from docx import Document
from process_mxliff import parse_mxliff_to_df
import help_format_tables as help
from merge_df import merge_dfs
from process_word import process_word_file

word_input_file = "input_files/ダリルさん-ja-en-T.docx"
word_output_folder = "output_files/"
mxliff_file = "input_files/ダリルさん-ja-en-T.mxliff"

# Word file has comments
df1 = process_word_file(word_input_file, word_output_folder)
# .mxliff file has repeating text
df2 = parse_mxliff_to_df(mxliff_file)

df = merge_dfs(df1, df2)

output_file_path = "output_files/merge_to_word_output.docx"

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
            if value is None or value == "None":
                cells[i].text = ""
            else:
                cells[i].text = str(value)
    
    help.format_table(table, comments=False)
    help.apply_conditional_formatting(table)
    
    doc.save(output_file_path)
    print(f"Merged tables saved as Word document: {output_file_path}.")

if __name__ == "__main__":
    dataframe_to_word_table(df, output_file_path)