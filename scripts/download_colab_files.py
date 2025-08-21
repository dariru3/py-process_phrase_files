'''
TODO: add this to the end of c_notebook?
'''
import os
from google.colab import files

def download_all_files(directory_path):
    """Downloads all files from the specified directory in Colab."""
    if not os.path.exists(directory_path):
        print(f"Directory not found: {directory_path}")
        return

    print(f"Downloading files from: {directory_path}")
    for filename in os.listdir(directory_path):
        file_path = os.path.join(directory_path, filename)
        if os.path.isfile(file_path):
            try:
                files.download(file_path)
                print(f"Downloaded: {filename}")
            except Exception as e:
                print(f"Error downloading {filename}: {e}")

# Specify the directory you want to download files from
directory_to_download = "/content/sample_data"
download_all_files(directory_to_download)
