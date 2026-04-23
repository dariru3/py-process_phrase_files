import os

from .df_to_word import get_file_pairs, process_files


def get_merged_output_path(docx_file, output_folder):
    base_name, _ = os.path.splitext(os.path.basename(docx_file))
    return os.path.join(output_folder, f"{base_name}_merged.docx")


def filter_unprocessed_pairs(pairs, output_folder):
    unprocessed_pairs = []
    for docx_file, mxliff_file in pairs:
        merged_file_path = get_merged_output_path(docx_file, output_folder)
        if os.path.exists(merged_file_path):
            base_name, _ = os.path.splitext(os.path.basename(docx_file))
            print(
                f"Skipped processing for {base_name} because the merged file already exists."
            )
            continue
        unprocessed_pairs.append((docx_file, mxliff_file))
    return unprocessed_pairs


def run_pipeline(input_folder, output_folder, skip_existing=True):
    pairs = get_file_pairs(input_folder)
    pairs_to_process = (
        filter_unprocessed_pairs(pairs, output_folder) if skip_existing else pairs
    )

    for docx_file, mxliff_file in pairs_to_process:
        print(f"File pair:\n{docx_file}\n{mxliff_file}")
        process_files(docx_file, mxliff_file, input_folder, output_folder)

    return len(pairs_to_process)
