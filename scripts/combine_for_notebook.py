import ast
import os
import pprint


def get_file_content(filename):
    if os.path.isfile(filename):
        with open(filename, "r", encoding="utf-8") as file:
            return file.readlines()
    else:
        print(f"File {filename} does not exist and is skipped.")
        return []


def sort_imports(line, from_imports, regular_imports):
    stripped_line = line.strip()
    if stripped_line.startswith("from .") or stripped_line.startswith("from src."):
        return True
    if stripped_line.startswith("from ") or stripped_line.startswith("import "):
        if stripped_line.startswith("from "):
            # Collect names per module to allow merging duplicate imports
            module, names = stripped_line[5:].split(" import ", 1)
            name_parts = [name.strip() for name in names.split(",")]
            from_imports.setdefault(module, set()).update(name_parts)
        else:
            regular_imports.add(stripped_line)
        return True
    return False


def process_config_loader(file_path, colab_input_path, colab_output_path):
    """
    Reads, updates, and returns the CONFIG dictionary assignment string from config_loader.py.
    """

    content = get_file_content(file_path)
    file_str = "".join(content)

    tree = ast.parse(file_str)
    config_dict_node = None
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id == "CONFIG"
            for target in node.targets
        ):
            config_dict_node = node.value
            break

    if config_dict_node is None:
        raise RuntimeError(
            "Could not find CONFIG dictionary assignment in config_loader.py"
        )

    config_dict = ast.literal_eval(config_dict_node)

    # Update paths
    if "GeneralSettings" in config_dict:
        config_dict["GeneralSettings"]["InputFolderPath"] = colab_input_path
        config_dict["GeneralSettings"]["OutputFolderPath"] = colab_output_path

    pretty_config = pprint.pformat(config_dict)
    return f"CONFIG = {pretty_config}\n"


def combine_scripts_for_notebook(
    file_names, output_file_path, colab_input_path, colab_output_path
):
    colab_snippet = "# @title Step 2: Process Files\n"
    combined_scripts = ""  # Hold the combined content of all scripts
    from_imports = {}
    regular_imports = set()

    # Loop through all files in the directory
    for filename in file_names:
        if filename.endswith("config_loader.py"):
            updated_config_code = process_config_loader(
                filename, colab_input_path, colab_output_path
            )
            combined_scripts += updated_config_code
        elif filename.endswith(".py"):
            content = get_file_content(filename) or []
            combined_scripts += f"\n# Start of {filename}\n"

            for line in content:
                is_import = sort_imports(line, from_imports, regular_imports)
                if not is_import:
                    combined_scripts += line

            combined_scripts += f"\n# End of {filename}\n\n"
        else:
            print(f"File {filename} is not a Python script and is skipped.")

    merged_from_imports = [
        f"from {module} import {', '.join(sorted(names))}"
        for module, names in sorted(from_imports.items())
    ]
    imports_list = sorted(regular_imports) + merged_from_imports
    combined_imports = "\n".join(imports_list) + "\n"

    combined_output = colab_snippet + combined_imports + combined_scripts

    # Save the combined scripts to a new file
    with open(output_file_path, "w", encoding="utf-8") as output_file:
        output_file.write(combined_output)

    print(f"All scripts have been combined into {output_file_path}")


if __name__ == "__main__":
    colab_input_path = "/content/MagicBox/"
    colab_output_path = "/content/MagicBox/Output_Folder/"

    output_file_path = "colab/c_notebook_script.py"

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

    file_list = [f"{src_path}{file_name}" for file_name in src_file_names]

    # Add main.py to end of list
    file_list.append("scripts/main.py")

    combine_scripts_for_notebook(
        file_list, output_file_path, colab_input_path, colab_output_path
    )
