from docx import Document
from docx.shared import Mm
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from process_mxliff import parse_mxliff_to_df
from config import mxliff_filepath

mxliff_file = mxliff_filepath
df = parse_mxliff_to_df(mxliff_file)

output_file_path = "output_files/df_to_word_output.docx"

def format_table(table):
    table.style = 'Table Grid'
    row_widths = [11, 60, 70, 11] # missing comments column!
    for i, width in enumerate(row_widths):
        for cell in table.columns[i].cells:
            cell.width = Mm(width)

def change_cell_color(cells, background_color=None):
    for cell in cells:
        if background_color:
            shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), background_color))
            cell._tc.get_or_add_tcPr().append(shading_elm)

def apply_conditional_formatting(table):
    condition_column_index = 2
    format_column_index = 1
    background_color= "D9D9D9"

    for row in table.rows:
        # Check if the cell has text.
        condition_met = row.cells[condition_column_index].text.strip() != ""
        # Additionally, check if the next cell to the right has either "99", "100", or "101".
        next_cell_value = row.cells[condition_column_index + 1].text.strip()
        next_cell_condition_met = next_cell_value in ["99", "100", "101"]

        indexes_to_color = [format_column_index, condition_column_index, condition_column_index +1]

        if condition_met and next_cell_condition_met:
            cells_to_color = [row.cells[i] for i in indexes_to_color]
            change_cell_color(cells_to_color, background_color)

def dataframe_to_word_table(df, output_file_path):
    doc = Document()
    df.index = df.index + 1
    df.index.name = 'Index'
    df.reset_index(inplace=True)
    
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
    
    format_table(table)
    apply_conditional_formatting(table)
    
    doc.save(output_file_path)
    print(f"Word document has been saved to {output_file_path}.")

dataframe_to_word_table(df, output_file_path)