import pandas as pd

def table_to_df(table):
    data = []
    for row in table.rows:
        row_data = []
        for cell in row.cells:
            row_data.append(cell.text.strip())
        data.append(row_data)
    return pd.DataFrame(data, columns=['Index', 'Source', 'Target', 'Match', 'Comment'])