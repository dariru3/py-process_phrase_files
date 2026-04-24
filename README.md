# Phrase Files Formatter (Magic Box)

A small toolchain to merge paired `.docx` and `.mxliff` files into a single, formatted Word document per pair. It auto-pairs files by base name, preserves key text formatting, and applies conditional styles for quick review.

## Features

- **Auto pairing**: Matches `.docx` and `.mxliff` by the same base filename.
- **Merge to Word**: Produces one `<base>_merged.docx` per valid pair.
- **Formatting preservation**: Reapplies bold/italic, color, superscript/subscript where available for Japanese/English columns.
- **Conditional styling**: Grays rows based on configured match quality and/or comment hints.

## Project Layout

- `src/` — Processing pipeline
  - `pipeline.py`: Shared runtime entrypoint used by the CLI and Colab.
  - `config_loader.py`: Central formatting and processing configuration.
  - `process_word.py`: Extracts tables and formatting from input `.docx`.
  - `process_mxliff.py`: Parses `.mxliff` to a DataFrame.
  - `merge_df.py`: Aligns and merges Word/XLIFF data.
  - `df_to_word.py`: Builds the output Word table and saves files.
  - `format_helper.py`, `save_formatting.py`, `table_to_df.py`: Formatting helpers.
- `main.py`: CLI entry point for local runs.
- `phrase_files_formatter.ipynb`: Open this notebook directly in Colab from GitHub.
- `tests/`: Generated-fixture coverage for parsing, merging, and formatting.

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

1. Put your `.docx` and `.mxliff` files into an input folder with matching base names (e.g., `Example.docx` and `Example.mxliff`).
2. Run the CLI:
   - `python3 main.py --input /path/to/input --output /path/to/output`
3. Find results in your output folder as `<base>_merged.docx`.

Notes

- Input and output folders are selected at runtime through CLI flags.
- Existing merged files are skipped by default to avoid reprocessing.
- Add `--force` to reprocess files that already have merged outputs.

## Google Colab Usage

Use the committed notebook at the repo root to run without setting up Python locally.

1. Open [`phrase_files_formatter.ipynb`](https://colab.research.google.com/github/dariru3/py-process_phrase_files/blob/main/phrase_files_formatter.ipynb) in Colab.
2. Run **Step 1: Install From GitHub** to install the current repo version.
3. Run **Step 2: Upload Files** to upload valid `.docx` / `.mxliff` pairs.
4. Run **Step 3: Process Files** to call the same shared pipeline used by the CLI.
5. Download outputs from Colab when processing is complete.

Important

- When prompted by the browser, allow multiple automatic downloads for `colab.research.google.com` so Colab can save multiple merged files.
- When testing the notebook from a feature branch, change the `git clone --branch ...` value in Step 1 from `"main"` to your branch name, then switch it back to `"main"` before merging.

## Running Tests

Use the built-in unittest suite:

```
python3 -m unittest
```

## License

MIT — see `LICENSE` for details.
