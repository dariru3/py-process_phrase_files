'''
Copy/paste script below into cell before main cell in Google Colab
'''
# @title Step 1: Connect to Google Drive > MagicBox {display-mode: "form"}
!pip install python-docx
import os
from google.colab import drive
drive.mount('/content/drive')

# Path to 'MagicBox' folder
magic_box_path = '/content/drive/MyDrive/MagicBox/'
# Path to output folder
output_folder_path = '/content/drive/MyDrive/MagicBox/Output_Folder'

def check_point(checkpoint_count):
  check_counter = [1, 2, 3]
  return f"Check {check_counter[checkpoint_count - 1]} of {len(check_counter)}"

def ensure_folder_exists(folder_path):
  if not os.path.exists(folder_path):
    print(f"Missing folder: {folder_path}")
    os.makedirs(folder_path)
    print(f"Created folder: {folder_path}")
  else:
    print(f"Folder confirmed: {folder_path}")

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
  file_confirmation_message(has_docx, has_mxliff)

  return has_docx and has_mxliff

def file_confirmation_message(has_docx, has_mxliff):
  if has_docx and has_mxliff:
    print(f"OK. MagicBox contains at least one .docx file and one .mxliff file.")
    return True
  else:
    print(f"Warning:")
    if not has_docx:
        print("No .docx file found. Add file(s) to MagicBox.")
    if not has_mxliff:
        print("No .mxliff file found. Add files(s) to MagicBox.")
    return False

if __name__ == "__main__":
  # Check for folders and files
  print(f"\n{check_point(1)}") # Checkpoint 1
  ensure_folder_exists(magic_box_path)
  ensure_folder_exists(output_folder_path)

  print(f"\n{check_point(2)}") # Checkpoint 2
  has_files = check_for_files(magic_box_path)

  # Loop through the contents of the 'MagicBox' folder and print filenames for confirmation
  if (has_files):
    print("Files and folders in MagicBox:")
    for item in os.listdir(magic_box_path):
      print(f"- {item}")

  # Check if Output_Folder is empty
  print(f"\n{check_point(3)}") # Checkpoint 3
  if not os.listdir(output_folder_path):
      print(f"OK: Output_Folder is empty.")
  else:
      print(f"Warning: Output_Folder is NOT empty. Consider deleting unnecessary files.")
