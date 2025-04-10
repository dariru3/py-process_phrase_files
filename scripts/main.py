import os
from src.config_loader import CONFIG
from src.df_to_word import get_file_pairs, process_files

def filter_unprocessed_pairs(pairs, output_folder):
    unprocessed_pairs = []
    for docx_file, mxliff_file in pairs:
        # Extract the base name from the docx file (base name for mxliff file should be the same)
        base_name, _ = os.path.splitext(os.path.basename(docx_file))

        # Check if the merged file exists
        merged_filename = f"{base_name}_merged.docx"
        merged_file_path = os.path.join(output_folder, merged_filename)
        if not os.path.exists(merged_file_path):
            unprocessed_pairs.append((docx_file, mxliff_file))
        else:
            print(f"Skipped processing for {base_name} because the merged file already exists.")
    return unprocessed_pairs

def main():
    g_settings = CONFIG["GeneralSettings"]
    input_folder = g_settings["InputFolderPath"] # "input_files/"
    output_folder = g_settings["OutputFolderPath"] # "output_files/"

    pairs = get_file_pairs(input_folder)
    unprocessed_pairs = filter_unprocessed_pairs(pairs, output_folder)
    for docx_file, mxliff_file in unprocessed_pairs:
        print(f"File pair:\n{docx_file}\n{mxliff_file}")
        process_files(docx_file, mxliff_file,input_folder, output_folder)

# Use `python3 -m scripts.main` to run file from console
if __name__ == "__main__":
    main()
