# Phrase Files Formatter (Magic Box)

A small toolchain to merge paired `.docx` and `.mxliff` files into a single, formatted Word document per pair. It auto-pairs files by base name, preserves key text formatting, and applies conditional styles for quick review.

## Features

- **Auto pairing**: Matches `.docx` and `.mxliff` by the same base filename.
- **Merge to Word**: Produces one `<base>_merged.docx` per valid pair.
- **Formatting preservation**: Reapplies bold/italic, color, superscript/subscript where available for Japanese/English columns.
- **Conditional styling**: Grays rows based on configured match quality and/or comment hints.

## Project Layout

- `src/` — Processing pipeline
  - `config_loader.py`: Central configuration (paths, formatting rules).
  - `process_word.py`: Extracts tables and formatting from input `.docx`.
  - `process_mxliff.py`: Parses `.mxliff` to a DataFrame.
  - `merge_df.py`: Aligns and merges Word/XLIFF data.
  - `df_to_word.py`: Builds the output Word table and saves files.
  - `format_helper.py`, `save_formatting.py`, `table_to_df.py`: Formatting helpers.
- `scripts/`
  - `main.py`: CLI entry point for local runs.
  - `combine_for_notebook.py`: Builds a single Colab-friendly script from `src/`.
- `colab/`
  - `a_header_cell.md`: Instructions
  - `b_upload_to_colab.py`: Step 1 (upload pairs in Colab)
  - `c_notebook_script.py`: Step 2 (process pairs in Colab)
- `data/`
  - `input_files/`: Place source pairs here for local runs.
  - `output_files/`: Merged results are written here.
- `tests/`: Basic sanity tests for formatting and content mirroring.

## Requirements

- Python 3.8+

Install dependencies:

```
pip install -r requirements.txt
```

Alternative (explicit packages):

```
pip install python-docx pandas
```

## Local Usage

1. Put your `.docx` and `.mxliff` files into `data/input_files/` with matching base names (e.g., `Example.docx` and `Example.mxliff`).
2. Run the CLI:
   - `python3 -m scripts.main`
3. Find results in `data/output_files/` as `<base>_merged.docx`.

Notes

- Update input/output paths or formatting options in `src/config_loader.py` if your folders differ.
- Existing merged files are skipped to avoid reprocessing.

## Google Colab Usage

Use the prepared cells in `colab/` to run without setting up Python locally.

1. Create a new Colab notebook.
2. Add a text cell and paste the contents of `colab/a_header_cell.md`.
3. Add a code cell with the contents of `colab/b_upload_to_colab.py` and run “Step 1: Upload Files”.
4. Add a code cell with the contents of `colab/c_notebook_script.py` and run “Step 2: Process Files”.

Important

- When prompted by the browser, allow multiple automatic downloads for `colab.research.google.com` so Colab can save multiple merged files. See the note in `colab/a_header_cell.md` and the manual screenshot.

## Running Tests

Use the built-in unittest suite:

```
python3 -m unittest tests/test_main.py
```

## License

MIT — see `LICENSE` for details.
