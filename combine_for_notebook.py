import os

def combine_scripts_for_notebook(file_names, output_file_path = 'notebook_script.py'):
    # Hold the combined content of all scripts
    combined_scripts = ''
    # Save all import lines
    imports_set = set()

    # Loop through all files in the directory
    for filename in file_names():
        if filename.endswith('.py'):
            if os.path.isfile(filename):
                # Open and read the content of the file
                with open(filename, 'r', encoding='utf-8') as file:
                    content = file.readlines()
                    combined_scripts += f"\n# Start of {filename}\n"

                    for line in content:
                        stripped_line = line.strip()
                        if stripped_line.startswith('import ') or stripped_line.startswith('from '):
                            imports_set.add(stripped_line)
                        else:
                            combined_scripts += line

                    combined_scripts += f"\n# End of {filename}\n\n"
            else:
                print(f"File {filename} does not exist and is skipped.")
        else:
            print(f"File {filename} is not a Python script and is skipped.")

    imports_list = sorted(imports_set)
    combined_imports = '\n'.join(imports_list) + '\n\n'

    # Save the combined scripts to a new file
    with open(output_file_path, 'w', encoding='utf-8') as output_file:
        output_file.write(combined_imports)
        output_file.write(combined_scripts)

    print(f"All scripts have been combined into {output_file_path}")

if __name__ == "__main__":
    file_list = [
        'config_loader.py',
        'save_formatting.py',
        'format_helper.py',
        'process_word.py',
        'process_mxliff.py',
        'table_to_df.py',
        'merge_df.py',
        'df_to_word.py',
        'main.py',
    ]

    combine_scripts_for_notebook(file_list)
