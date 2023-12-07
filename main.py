import xlwings as xw
import os

def process_excel(file_path):
    # Start an instance of Excel and open the workbook
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sheet = wb.sheets[0]

    # Delete the first two columns
    sheet.range('A:B').delete()

    # Delete the first two rows
    sheet.range('1:2').delete()

    # Save the changes to the workbook
    wb.save()

    # Close the workbook and quit the Excel application
    wb.close()
    app.quit()

    print(f"Processed file: {file_path}")

def process_all_excel_in_folder(folder_path):
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(folder_path, file_name)
            process_excel(file_path)

# Example usage
process_all_excel_in_folder('input_files')
