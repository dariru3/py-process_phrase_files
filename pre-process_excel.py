import pandas as pd
import os

def process_excel(file_path, output_folder):
    # Load the Excel file into a pandas DataFrame, skipping the first two rows
    df = pd.read_excel(file_path, engine='openpyxl')

    # Delete the first two rows and the first two columns
    df = df.iloc[1:, 2:]

    # Construct new file path
    base_name = os.path.basename(file_path)
    name_part, extension = os.path.splitext(base_name)
    new_filename = name_part + '_pre-processed' + extension
    processed_file_path = os.path.join(output_folder, new_filename)

    # Check for output folder and create if necessary
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Save the processed DataFrame back to an Excel file
    df.to_excel(processed_file_path, index=False, header=False, engine='openpyxl')

    print(f"Processed file: {processed_file_path}")

def process_all_excel_in_folder(folder_path, output_folder):
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(folder_path, file_name)
            process_excel(file_path, output_folder)

# Example usage - make sure to adjust 'input_files' and 'output_files' paths accordingly
input_folder_path = 'input_files'
output_folder_path = 'output_files'

if __name__ == "__main__":
    process_all_excel_in_folder(input_folder_path, output_folder_path)
