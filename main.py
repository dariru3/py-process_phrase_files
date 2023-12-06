import os
import pandas as pd

def process_excel_as_csv(file_path, encoding):
    try:
        # Read the Excel file as a CSV, treating all data as string, with the specified encoding
        df = pd.read_csv(file_path, header=None, dtype=str, sep='\t', encoding=encoding)

        # Remove the first two rows and columns
        processed_df = df.iloc[2:, 2:]

        # Save the processed DataFrame back to a CSV file
        processed_file_path = file_path.replace('.xlsx', '_processed.csv')
        processed_df.to_csv(processed_file_path, index=False, header=False)

        print(f"Processed file saved as: {processed_file_path}")
    except UnicodeDecodeError as e:
        print(f"Error processing {file_path} with encoding {encoding}: {e}")

def process_all_excel_as_csv_in_folder(folder_path):
    # List of possible encodings to try
    encodings = ['utf-8', 'iso-8859-1', 'cp1252', 'utf-16']

    # List all files in the given folder
    for file_name in os.listdir(folder_path):
        # Check if the file is an Excel file
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(folder_path, file_name)
            for encoding in encodings:
                try:
                    process_excel_as_csv(file_path, encoding)
                    break  # If successful, break out of the encoding loop
                except UnicodeDecodeError:
                    continue  # Try the next encoding

# Example usage
process_all_excel_as_csv_in_folder('input_files')
