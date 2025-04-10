import pandas as pd
from config_loader import CONFIG

def table_to_df(table):
    column_headers = CONFIG["GeneralSettings"]["Column_Headers"]
    data = []
    for row in table.rows:
        row_data = []
        for cell in row.cells:
            row_data.append(cell.text.strip())
        data.append(row_data)
    return pd.DataFrame(data, columns=column_headers)