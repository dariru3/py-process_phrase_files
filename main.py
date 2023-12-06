import os
import pandas as pd

def process_excel(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Remove the first two rows and columns
    processed_df = df.iloc[2:, 2:]

    # Save the processed DataFrame back to an Excel file
    processed_file_path = file_path.replace('.xlsx', '_processed.xlsx')
    processed_df.to_excel(processed_file_path, index=False)

    print(f"Processed file saved as: {processed_file_path}")

def process_all_excel_in_folder(folder_path):
    # List all files in the given folder
    for file_name in os.listdir(folder_path):
        # Check if the file is an Excel file
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(folder_path, file_name)
            process_excel(file_path)

# Example usage
process_all_excel_in_folder('input_files')
