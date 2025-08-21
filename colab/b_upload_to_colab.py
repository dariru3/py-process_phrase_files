'''
Copy/paste this script into the cell before *Step 1* in Google Colab
'''
# @title Step 1: Upload Files {display-mode: "form"}
!pip install python-docx
import os
from google.colab import files
import shutil

# Path to 'MagicBox' folder
magic_box_path = '/content/MagicBox/'
# Path to output folder
output_folder_path = '/content/MagicBox/Output_Folder'

def ensure_folder_exists(folder_path):
  if not os.path.exists(folder_path):
    os.makedirs(folder_path)

def clear_folder(folder_path):
    if os.path.exists(folder_path):
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f'Error: failed to clear {file_path}. Reason: {e}')

def upload_files(upload_path):
    # Clear the folder before uploading new files
    clear_folder(upload_path)
    ensure_folder_exists(upload_path)

    print("\nPlease upload one .docx and one .mxliff file.")
    uploaded = files.upload()

    docx_files = [fn for fn in uploaded.keys() if fn.endswith('.docx')]
    mxliff_files = [fn for fn in uploaded.keys() if fn.endswith('.mxliff')]

    # Validate file counts
    if len(docx_files) != 1 or len(mxliff_files) != 1:
        print("Error: Please upload exactly one .docx and one .mxliff file.")
        # Clean up any uploaded files
        for fn in uploaded.keys():
            os.remove(fn)
        return False

    docx_file = docx_files[0]
    mxliff_file = mxliff_files[0]

    # Move the uploaded files to the magic_box_path
    shutil.move(docx_file, os.path.join(upload_path, docx_file))
    shutil.move(mxliff_file, os.path.join(upload_path, mxliff_file))

    return True

if __name__ == "__main__":
    ensure_folder_exists(magic_box_path)
    ensure_folder_exists(output_folder_path)

    if upload_files(magic_box_path):
        print("\nFile upload successful.")
    else:
        print("\nFile upload failed. Please follow the instructions and try again.")
