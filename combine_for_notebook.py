import os

# Adjust these paths according to your needs
directory_path = 'temp'
output_file_path = 'temp/combined_script.py'

# This will hold the combined content of all scripts
combined_scripts = ''

# Loop through all files in the directory
for filename in os.listdir(directory_path):
    if filename.endswith('.py'):
        # Construct full file path
        file_path = os.path.join(directory_path, filename)
        # Open and read the content of the file
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
            combined_scripts += f"\n# Start of {filename}\n"
            combined_scripts += content
            combined_scripts += f"\n# End of {filename}\n\n"

# Save the combined scripts to a new file
with open(output_file_path, 'w', encoding='utf-8') as output_file:
    output_file.write(combined_scripts)

print(f"All scripts have been combined into {output_file_path}")
