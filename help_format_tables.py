from docx.shared import Mm
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

def change_cell_color(cells, background_color=None):
    for cell in cells:
        if background_color:
            shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), background_color))
            cell._tc.get_or_add_tcPr().append(shading_elm)

def format_table(table, comments=True):
    table.style = 'Table Grid'
    if comments:
        row_widths = [9, 90, 110, 10, 50]
    else:
        row_widths = [11, 60, 70, 11]
    
    for i, width in enumerate(row_widths):
        if comments:
            table.columns[i].width = Mm(width)
        else:
            for cell in table.columns[i].cells:
                cell.width = Mm(width)

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