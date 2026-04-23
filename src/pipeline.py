import logging
import os

from .df_to_word import get_file_pairs, process_files

logger = logging.getLogger(__name__)


def get_merged_output_path(docx_file, output_folder):
    base_name, _ = os.path.splitext(os.path.basename(docx_file))
    return os.path.join(output_folder, f"{base_name}_merged.docx")


def run_pipeline(input_folder, output_folder, skip_existing=True):
    pairs = get_file_pairs(input_folder)
    processed_count = 0

    for docx_file, mxliff_file in pairs:
        if skip_existing:
            merged_file_path = get_merged_output_path(docx_file, output_folder)
            if os.path.exists(merged_file_path):
                base_name, _ = os.path.splitext(os.path.basename(docx_file))
                logger.info(
                    "Skipped processing for %s because the merged file already exists.",
                    base_name,
                )
                continue

        logger.info("Processing file pair: %s <-> %s", docx_file, mxliff_file)
        process_files(docx_file, mxliff_file, input_folder, output_folder)
        processed_count += 1

    return processed_count
