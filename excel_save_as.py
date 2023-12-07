import xlwings as xw
import os

def process_excel(file_path, output_folder):
    # Start an instance of Excel and open the workbook
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sheet = wb.sheets[0]

    # Delete the first two columns
    sheet.range('A:B').delete()

    # Delete the first two rows
    sheet.range('1:2').delete()

    # Construct the new file path with '_processed' appended before the file extension
    base_name = os.path.basename(file_path)
    name_part, extension = os.path.splitext(base_name)
    new_file_name = name_part + '_processed' + extension
    new_file_path = os.path.join(output_folder, new_file_name)

    # Save the changes to a new workbook
    wb.save(new_file_path)

    # Close the workbook and quit the Excel application
    wb.close()
    app.quit()

    print(f"Processed file saved as: {new_file_path}")

def process_all_excel_in_folder(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for file_name in os.listdir(input_folder):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(input_folder, file_name)
            process_excel(file_path, output_folder)

# Example usage
input_folder = 'input_files'
output_folder = 'output_csv_files'
process_all_excel_in_folder(input_folder, output_folder)
