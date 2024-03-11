from docx.shared import Mm
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from config_loader import CONFIG

def change_cell_color(cells, background_color=None):
    for cell in cells:
        if background_color:
            shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), background_color))
            cell._tc.get_or_add_tcPr().append(shading_elm)

def format_table(table, comments=True):
    f_settings = CONFIG["TableFormattingSettings"]
    table.style = 'Table Grid'
    if comments:
        row_widths = f_settings["RowWidthsWithComments"] # [9, 90, 110, 10, 50]
    else:
        row_widths = f_settings["RowWidthsWithoutComments"] # [11, 60, 70, 11]
    
    for i, width in enumerate(row_widths):
        if comments:
            table.columns[i].width = Mm(width)
        else:
            for cell in table.columns[i].cells:
                cell.width = Mm(width)

def apply_conditional_formatting(table):
    '''
    Change row cell colors to gray if either condition is met.
    Condition 1: There is text in the target cell and the match is either "100" or "101"
    Condition 2: There is text in the target cell and the comment is "lock" or "locked"
    '''
    cf_settings = CONFIG["ConditionalFormattingSettings"]
    target_column_index = cf_settings["TargetColumnIndex"] # 2
    match_column_index = cf_settings["MatchColumnIndex"] #3
    comment_column_index = cf_settings["CommentColumnIndex"] # 4
    comment_to_gray = cf_settings["CommentToGray"] # ['lock', 'locked']
    match_to_gray = cf_settings["MatchToGray"] # ['100', '101']
    background_color= cf_settings["BackgroundColor"] # "D9D9D9"

    for row in table.rows:
        target_value = row.cells[target_column_index].text.strip()
        match_value = row.cells[match_column_index].text.strip()
        comment_value = row.cells[comment_column_index].text.lower().strip()

        condition_1_met = target_value and match_value in match_to_gray
        condition_2_met = target_value and comment_value in comment_to_gray

        if condition_1_met or condition_2_met:
            cells_to_color = [cell for cell in row.cells]
            change_cell_color(cells_to_color, background_color)