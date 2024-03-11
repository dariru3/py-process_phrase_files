# Phrase Export Documents Processing and Merging Tool

This Python tool automates the processing and merging of paired `.docx` and `.mxliff` files, applying conditional formatting to the resulting Word tables. It's designed to streamline workflows that involve handling translations and comments across different document formats.

## Features

- **Automatic File Pairing**: Identifies and pairs `.docx` and `.mxliff` files based on their base names.
- **Conditional Formatting**: Applies gray background color to rows based on specific content criteria.
- **Document Merging**: Combines content from paired files into a single Word document with formatted tables.

## Prerequisites

- Python 3.6 or later
- `python-docx` library for reading and writing `.docx` files
- `pandas` library for data manipulation

## Installation

1. Ensure Python 3.6 or later is installed on your system.
2. Install required Python packages:
   ```sh
   pip install python-docx pandas
   ```

## Usage

1. Place `.docx` and `.mxliff` files in the `input_files/` directory. Files should be named such that each `.docx` file has a corresponding `.mxliff` file with the same base name.
2. Run the script:
   ```sh
   python df_to_word.py
   ```
3. Processed files will be saved in the `output_files/` directory, with `_merged.docx` appended to the base name.

## Structure

- **`df_to_word.py`**: The main script that orchestrates the file processing, merging, and output.
- **`process_word.py`**: Handles `.docx` file processing.
- **`process_mxliff.py`**: Handles `.mxliff` file processing.
- **`merge_df.py`**: Merges data frames obtained from `.docx` and `.mxliff` files.
- **`table_to_df.py`**: Converts tables in `.docx` files to pandas DataFrames.
- **`help_format_tables.py`**: Contains helper functions for table formatting and conditional coloring.

## License

Specify your project's license here.