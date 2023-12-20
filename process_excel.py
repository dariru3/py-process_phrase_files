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

    # Construct new file path
    base_name = os.path.basename(file_path)
    name_part, extension = os.path.splitext(base_name)
    new_filename = name_part + '_processed' + extension
    processed_file_path = os.path.join(output_folder, new_filename)

    # Check for output folder
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Save the changes to the workbook
    wb.save(processed_file_path)

    # Close the workbook and quit the Excel application
    wb.close()
    app.quit()

    print(f"Processed file: {processed_file_path}")

def process_all_excel_in_folder(folder_path, output_folder):
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(folder_path, file_name)
            process_excel(file_path, output_folder)

# Example usage
process_all_excel_in_folder('input_files', 'output_files')
