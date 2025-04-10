import os

def get_file_content(filename):
    if os.path.isfile(filename):
        with open(filename, "r", encoding="utf-8") as file:
            return file.readlines()
    else:
        print(f"File {filename} does not exist and is skipped.")
        return []

def sort_imports(line, imports_set):
    stripped_line = line.strip()
    if stripped_line.startswith("from .") or stripped_line.startswith("from src."):
        return True
    if (stripped_line.startswith("from ") or stripped_line.startswith("import ")):
        imports_set.add(stripped_line)
        return True
    return False

def combine_scripts_for_notebook(file_names, output_file_path="colab/notebook_script.py"):
    combined_scripts = "" # Hold the combined content of all scripts
    imports_set = set() # Save all import lines

    # Loop through all files in the directory
    for filename in file_names:
        if filename.endswith(".py"):
            content = get_file_content(filename) or []
            combined_scripts += f"\n# Start of {filename}\n"

            for line in content:
                is_import = sort_imports(line, imports_set)
                if is_import == False:
                    combined_scripts += line

            combined_scripts += f"\n# End of {filename}\n\n"
        else:
            print(f"File {filename} is not a Python script and is skipped.")

    imports_list = sorted(imports_set)
    combined_imports = "\n".join(imports_list) + "\n"

    # Save the combined scripts to a new file
    with open(output_file_path, "w", encoding="utf-8") as output_file:
        output_file.write(combined_imports)
        output_file.write(combined_scripts)

    print(f"All scripts have been combined into {output_file_path}")

if __name__ == "__main__":
    src_path = "src/"
    script_path = "scripts/"
    file_list = [
        f"{src_path}config_loader.py",
        f"{src_path}save_formatting.py",
        f"{src_path}format_helper.py",
        f"{src_path}process_word.py",
        f"{src_path}process_mxliff.py",
        f"{src_path}table_to_df.py",
        f"{src_path}merge_df.py",
        f"{src_path}df_to_word.py",
        f"{script_path}main.py"
    ]

    combine_scripts_for_notebook(file_list)
