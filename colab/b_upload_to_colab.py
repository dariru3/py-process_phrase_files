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

    print("\nUpload one or more .docx and .mxliff files.")
    print("Pairs are matched by the same base filename.\n")
    uploaded = files.upload()

    if not uploaded:
        print("No files uploaded.")
        return False

    # Build maps of base name -> filename for each extension
    docx_map = {os.path.splitext(fn)[0]: fn for fn in uploaded.keys() if fn.lower().endswith('.docx')}
    mxliff_map = {os.path.splitext(fn)[0]: fn for fn in uploaded.keys() if fn.lower().endswith('.mxliff')}

    paired_basenames = sorted(set(docx_map.keys()) & set(mxliff_map.keys()))

    if len(paired_basenames) == 0:
        print("Error: No valid pairs found. Ensure each .docx has a matching .mxliff with the same base name.")
        # Clean up any uploaded files from working directory
        for fn in uploaded.keys():
            try:
                os.remove(fn)
            except FileNotFoundError:
                pass
        return False

    # Move only valid pairs to the upload path
    print(f"Found {len(paired_basenames)} pair(s):")
    for base in paired_basenames:
        docx_file = docx_map[base]
        mxliff_file = mxliff_map[base]
        print(f"- {docx_file}  <->  {mxliff_file}")
        shutil.move(docx_file, os.path.join(upload_path, docx_file))
        shutil.move(mxliff_file, os.path.join(upload_path, mxliff_file))

    # Clean up any non-paired uploads left in the working directory
    paired_files = set(docx_map[b] for b in paired_basenames) | set(mxliff_map[b] for b in paired_basenames)
    for fn in uploaded.keys():
        if fn not in paired_files:
            try:
                os.remove(fn)
            except FileNotFoundError:
                pass

    return True

if __name__ == "__main__":
    ensure_folder_exists(magic_box_path)
    ensure_folder_exists(output_folder_path)

    if upload_files(magic_box_path):
        print("\nFile upload successful.")
    else:
        print("\nFile upload failed. Please follow the instructions and try again.")
