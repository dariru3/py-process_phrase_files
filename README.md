## py-process_phrase_files

# Phrase Integration Processing Scripts

## Overview

This project comprises two Python scripts designed to automate the preprocessing of Excel files for Phrase (formerly Memsource) and the post-processing of Word files outputted from Phrase. These scripts are intended to streamline the workflow associated with translating and localizing content by formatting Excel and Word files according to the requirements of Phrase.

## Scripts Description

1. **Excel Processing Script**
   - This script processes `.xlsx` files by deleting the first two columns and the first two rows from each sheet. It is intended to prepare Excel files for import into Phrase.

2. **Word Processing Script**
   - This script processes `.docx` files by removing certain tables and modifying the remaining table's structure and cell formatting. It is designed to handle the output files from Phrase, aligning them with the desired document structure.

## Requirements

- Python 3.x
- Microsoft Excel installed on the machine (required for `xlwings`).
- Libraries: `xlwings`, `python-docx`
  - Install these libraries using `pip install xlwings python-docx`.

## Excel Processing Script Usage

1. Place the Excel files to be processed in a folder (e.g., `input_files`).
2. Run the script `process_all_excel_in_folder('input_files')`.
3. Processed Excel files will have their first two columns and two rows deleted.

## Word Processing Script Usage

1. Place the Word files to be processed in a folder (e.g., `input_files`).
2. Run the script `process_all_word_files_in_folder('input_files')`.
3. Processed Word files will have specific tables removed and modified, as per the script's logic.

## Important Notes

- Ensure you have backups of your original files before running these scripts, as the changes they make are irreversible.
- These scripts are part of a larger workflow involving Phrase. Adjustments may be necessary if there are changes in file formats or requirements.
