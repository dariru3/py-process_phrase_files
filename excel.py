import xlwings as xw
import os

def process_excel(file_path):
    # Start an instance of Excel and open the workbook
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sheet = wb.sheets[0]

    # Determine the last row and column
    last_row = sheet.range('C' + str(sheet.cells.last_cell.row)).end('up').row
    last_column = sheet.range('XFD1').end('left').column

    # Create a new workbook
    new_wb = xw.Book()
    new_sheet = new_wb.sheets[0]

    # Read and copy the data, skipping the first two rows and columns
    for row in range(3, last_row + 1):
        values = sheet.range((row, 3), (row, last_column)).value
        new_sheet.range((row - 2, 1), (row - 2, last_column - 2)).value = values

    # Save the processed workbook
    processed_file_path = file_path.replace('.xlsx', '_processed.xlsx')
    new_wb.save(processed_file_path)
    new_wb.close()

    # Quit the Excel application
    app.quit()

    print(f"Processed file saved as: {processed_file_path}")

def process_all_excel_in_folder(folder_path):
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(folder_path, file_name)
            process_excel(file_path)

# Example usage
process_all_excel_in_folder('input_files')
