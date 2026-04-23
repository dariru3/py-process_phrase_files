import argparse
import logging
import os

from src.pipeline import run_pipeline


def build_parser():
    parser = argparse.ArgumentParser(
        description="Merge paired .docx and .mxliff files into formatted Word output."
    )
    parser.add_argument(
        "--input",
        required=True,
        help="Folder containing paired .docx and .mxliff files.",
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Folder where merged .docx files will be written.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Reprocess files even if an existing merged output is present.",
    )
    return parser


def main(argv=None):
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    parser = build_parser()
    args = parser.parse_args(argv)

    input_folder = os.path.abspath(args.input)
    output_folder = os.path.abspath(args.output)

    if not os.path.isdir(input_folder):
        parser.error(f"Input directory does not exist: {input_folder}")

    os.makedirs(output_folder, exist_ok=True)

    processed_count = run_pipeline(
        input_folder,
        output_folder,
        skip_existing=not args.force,
    )
    logging.info("Processed %s file pair(s).", processed_count)
    return processed_count

# Use `python3 main.py` to run from the console.
if __name__ == "__main__":
    main()
