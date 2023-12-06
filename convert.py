import pandas as pd
import os

def convert_excel_to_csv(file_path, output_folder):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Construct the CSV file path
    base_name = os.path.basename(file_path)
    csv_file_name = os.path.splitext(base_name)[0] + '.csv'
    csv_file_path = os.path.join(output_folder, csv_file_name)

    # Save as CSV
    df.to_csv(csv_file_path, index=False)

    return csv_file_path

def convert_all_excel_in_folder_to_csv(input_folder, output_folder):
    for file_name in os.listdir(input_folder):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(input_folder, file_name)
            convert_excel_to_csv(file_path, output_folder)

# Example usage
convert_all_excel_in_folder_to_csv('input_files', 'output_csv_files')
