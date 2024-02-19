from docx import Document
from process_mxliff import parse_mxliff_to_df
from config import mxliff_filepath

mxliff_file = mxliff_filepath
output_file_path = "output_files/df_to_word_output.docx"

def dataframe_to_word_table(df, output_file_path):
    doc = Document()
    table = doc.add_table(rows=1, cols=len(df.columns) + 1)
    
    # Add header row
    table.cell(0,0).text = 'Index'
    for i, column in enumerate(df.columns):
        table.cell(0, i + 1).text = str(column)
    
    # Add data rows
    for index, row in df.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(index)
        for i, value in enumerate(row):
            cells[i + 1].text = str(value)
    
    doc.save(output_file_path)
    print(f"Word document has been saved to {output_file_path}.")

df = parse_mxliff_to_df(mxliff_file)
dataframe_to_word_table(df, output_file_path)