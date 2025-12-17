# Repository Guidelines

## Project Structure & Modules
- Core pipeline lives in `src/`: `process_word.py` (extract `.docx` tables), `process_mxliff.py` (parse `.mxliff`), `merge_df.py` (aligns both), `df_to_word.py` (writes merged output), helpers like `format_helper.py` and `config_loader.py`.
- CLI entry is `scripts/main.py`; Colab flow is under `colab/` (header, upload, notebook script).
- Sample I/O lives in `data/input_files/` and `data/output_files/`; tests sit in `tests/`.

## Setup, Build & Run
- Install deps: `pip install -r requirements.txt` (Python 3.8+).
- Local run: `python3 -m scripts.main` (reads from `data/input_files/`, writes `<base>_merged.docx` to `data/output_files/`).
- Colab run: follow the three cells in `colab/` (header, upload, process) for browser-based execution.

## Testing
- Suite uses `unittest`: `python3 -m unittest tests/test_main.py`.
- Add new tests in `tests/` with names like `test_<feature>.py`; prefer table-driven cases for row/formatting combos.
- Keep coverage on merge correctness (row alignment) and formatting preservation (bold/italic, color, superscript/subscript).

## Coding Style & Naming
- Python only; default to PEP8 (4-space indent, snake_case functions/vars, UpperCamelCase classes).
- Favor pure functions in `src/` modules; keep file I/O centralized in `scripts/main.py` and `config_loader.py`.
- When adding format rules, keep constants together in `config_loader.py` to avoid scattering style knobs.

## Data, Config, and Safety
- Do not commit real client files; use `data/input_files/` for local samples and clean up outputs after runs.
- Configurable paths/thresholds live in `src/config_loader.py`; update defaults there rather than hard-coding.
- Existing merged files are skipped; remove them manually if you need a clean re-run.

## Commit & Pull Request Guidelines
- Follow the existing Conventional Commit style seen in history (`docs:`, `chore:`, etc.); keep messages imperative and scoped.
- PRs should include: short description of the user-facing change, mention of config updates, test command output, and any before/after notes for formatting behavior.
