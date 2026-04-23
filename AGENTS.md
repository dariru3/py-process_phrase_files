# Repository Guidelines

## Project Structure & Modules
- Core pipeline lives in `src/`: `process_word.py` (extract `.docx` tables), `process_mxliff.py` (parse `.mxliff`), `merge_df.py` (aligns both), `df_to_word.py` (writes merged output), helpers like `format_helper.py` and `config_loader.py`.
- CLI entry is `main.py`; Colab flow uses the committed root notebook `phrase_files_formatter.ipynb`.
- Tests sit in `tests/`; runtime input and output folders are chosen explicitly when running the CLI.

## Setup, Build & Run
- Install deps: `pip install -r requirements.txt` (Python 3.8+).
- Local run: `python3 main.py --input /path/to/input --output /path/to/output`.
- Colab run: open `phrase_files_formatter.ipynb` from GitHub in Colab and run the notebook cells.

## Testing
- Suite uses `unittest`: `python3 -m unittest tests/test_pipeline.py`.
- Add new tests in `tests/` with names like `test_<feature>.py`; prefer table-driven cases for row/formatting combos.
- Keep coverage on merge correctness (row alignment) and formatting preservation (bold/italic, color, superscript/subscript).

## Coding Style & Naming
- Python only; default to PEP8 (4-space indent, snake_case functions/vars, UpperCamelCase classes).
- Favor pure functions in `src/` modules; keep file I/O centralized in `main.py` and `config_loader.py`.
- When adding format rules, keep constants together in `config_loader.py` to avoid scattering style knobs.

## Data, Config, and Safety
- Do not commit real client files; keep local sample inputs and outputs outside the repo or in ignored scratch locations.
- Configurable thresholds and formatting rules live in `src/config_loader.py`; do not hard-code them.
- Existing merged files are skipped; remove them manually if you need a clean re-run.

## Commit & Pull Request Guidelines
- Follow the existing Conventional Commit style seen in history (`docs:`, `chore:`, etc.); keep messages imperative and scoped.
- PRs should include: short description of the user-facing change, mention of config updates, test command output, and any before/after notes for formatting behavior.
