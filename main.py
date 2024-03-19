from config_loader import CONFIG
from df_to_word import get_file_pairs, process_files

def main():
    g_settings = CONFIG["GeneralSettings"]
    input_folder = g_settings["InputFolderPath"] # "input_files/"
    output_folder = g_settings["OutputFolderPath"] # "output_files/"

    pairs = get_file_pairs(input_folder)
    for docx_file, mxliff_file in pairs:
        print(f"File pair:\n{docx_file}\n{mxliff_file}")
        process_files(docx_file, mxliff_file,input_folder, output_folder)

if __name__ == "__main__":
    main()