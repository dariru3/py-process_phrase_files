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
    '''
    Change row cell colors to gray if either condition is met.
    There is text in the target cell and ...
    Condition 1: the match cell as either "100" or "101"
    Condition 2: the text in the comment cell is "lock" or "locked"
    '''
    target_column_index = 2
    match_column_index = 3
    comment_column_index = 4
    comment_to_gray = ['lock', 'locked']
    match_to_gray = ['100', '101']
    background_color= "D9D9D9"

    for row in table.rows:
        target_value = row.cells[target_column_index].text.strip()
        match_value = row.cells[match_column_index].text.strip()
        comment_value = row.cells[comment_column_index].text.lower().strip()

        condition_1_met = match_value in match_to_gray
        condition_2_met = comment_value in comment_to_gray

        if target_value and (condition_1_met or condition_2_met):
            cells_to_color = [cell for cell in row.cells]
            change_cell_color(cells_to_color, background_color)