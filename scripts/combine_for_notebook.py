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
    colab_snippet = "# @title Step 2: Run Magic Box\n"
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

    combined_output = colab_snippet + combined_imports + combined_scripts

    # Save the combined scripts to a new file
    with open(output_file_path, "w", encoding="utf-8") as output_file:
        output_file.write(combined_output)

    print(f"All scripts have been combined into {output_file_path}")

if __name__ == "__main__":
    src_path = "src/"
    src_file_names = [
        "config_loader.py",
        "save_formatting.py",
        "format_helper.py",
        "process_word.py",
        "process_mxliff.py",
        "table_to_df.py",
        "merge_df.py",
        "df_to_word.py",
    ]
    script_path = "scripts/" # f"{script_path}main.py"

    file_list = [ f"{src_path}{file_name}" for file_name in src_file_names]
    file_list.append(f"{script_path}main.py")

    combine_scripts_for_notebook(file_list)
