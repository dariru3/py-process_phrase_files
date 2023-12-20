## py-process_phrase_files

# Phrase Integration Processing Scripts

## Overview

This project includes two Python scripts designed to automate the preprocessing of Excel files for Phrase (formerly Memsource) and the post-processing of Word files outputted from Phrase. These scripts streamline the workflow associated with translating and localizing content by formatting Excel to meet Phrase's requirements and Word files to share externally.

## Scripts Description

1. **Excel Processing Script**
   - Processes `.xlsx` files by deleting the first two columns and rows from each sheet, preparing them for import into Phrase.
   - Processed files are saved with a "_processed" suffix in the filename to an output folder.

2. **Word Processing Script**
   - Processes `.docx` files by removing certain tables and modifying the remaining table's structure and cell formatting.
   - Designed to handle Phrase output files, aligning them with the desired document structure.
   - Processed files are saved with a "_processed" suffix in the filename to an output folder.

## Requirements

- Python 3.x
- Microsoft Excel installed on the machine (required for `xlwings`).
- Libraries: `xlwings`, `python-docx`
  - Install these libraries using `pip install xlwings python-docx`.

## Excel Processing Script Usage

1. Place Excel files in an input folder (e.g., `input_files`).
2. Run `process_all_excel_in_folder('input_files', 'output_files')`.
3. Processed files are saved in an output folder (e.g., `output_files`) with "_processed" appended to their filenames.

## Word Processing Script Usage

1. Place Word files in an input folder (e.g., `input_files`).
2. Run `process_all_word_files_in_folder('input_files', 'output_files')`.
3. Processed files are saved in an output folder (e.g., `output_files`) with "_processed" appended to their filenames.