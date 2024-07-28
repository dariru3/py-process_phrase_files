# Cell 1

import os
from google.colab import drive

drive.mount('/content/drive')

# Path to your 'MagicBox' folder
magic_box_path = '/content/drive/MyDrive/MagicBox/'
config_file_download_link = "https://drive.google.com/file/d/1E0gUuUhd8XsI-xRwMpDm994wt2kQaYTl/view?usp=sharing"

# Define the paths to the folders and file
output_folder_path = '/content/drive/MyDrive/MagicBox/Output_Folder'
config_folder_path = '/content/drive/MyDrive/MagicBox/configFiles'
config_file_path = os.path.join(config_folder_path, 'config.json')

def ensure_folder_exists(folder_path):
  if not os.path.exists(folder_path):
    os.makedirs(folder_path)
    print(f"Created folder: {folder_path}")

import os

def check_for_files(directory_path):
    has_docx = False
    has_mxliff = False

    # Iterate over the items in the directory
    for item in os.listdir(directory_path):
        item_path = os.path.join(directory_path, item)
        if os.path.isfile(item_path):  # Ensure the item is a file
            if item.endswith('.docx'):
                has_docx = True
            elif item.endswith('.mxliff'):
                has_mxliff = True

        # If both file types are found, no need to continue checking
        if has_docx and has_mxliff:
            break

    # Communicate the findings directly within the function
    if has_docx and has_mxliff:
        print("\nCheck 1: The directory contains at least one .docx file and one .mxliff file.")
        print("\nFiles in 'Magic Box':")
    else:
        print("Check 1:")
        if not has_docx:
            print("No .docx file found. Add file(s) to MagicBox.")
        if not has_mxliff:
            print("No .mxliff file found. Add files(s) to MagicBox.")

# Check for files and folders
check_for_files(magic_box_path)
ensure_folder_exists(output_folder_path)
ensure_folder_exists(config_folder_path)

# Loop through the contents of the 'MagicBox' folder
for item in os.listdir(magic_box_path):
    # Construct the full path to the item
    item_path = os.path.join(magic_box_path, item)
    if not os.path.isdir(item_path):
      print(f"{item}")

# Check if Output_Folder is empty
if not os.listdir(output_folder_path):
    print("\nCheck 2: Output_Folder is empty.")
else:
    print("\nCheck 2: Output_Folder is NOT empty. Consider deleting unnecessary files.")

# Check if config.json exists in configFolder
if os.path.exists(config_file_path):
    print("\nCheck 3: Configuration file OK\n")
    !pip install python-docx
else:
    print("\nCheck 3: Configuration file missing from MagicBox/configFolder!")
    print(f"\nDownload the configuration file from: {config_file_download_link}")
    print("\nAfter downloading, please upload it to the 'configFiles' folder within your 'MagicBox' directory.\n")
