import xml.etree.ElementTree as ET
import pandas as pd
from config import tmx_filepath

# Define the path to your TMX file
tmx_file = tmx_filepath

# Parse the TMX file
tree = ET.parse(tmx_file)
root = tree.getroot()

# Extracting the namespace map from the TMX file's root element
namespaces = {'xml': 'http://www.w3.org/XML/1998/namespace'}  # Common namespace for xml:lang attributes

# Initialize lists to hold the extracted data
sources = []
targets = []

# Loop through each translation unit in the TMX file
for tu in root.findall('body/tu'):
    source_text = ''
    target_text = ''
    
    # Extracting Japanese source segments
    ja_tuv = tu.find('tuv[@xml:lang="ja"]', namespaces=namespaces)
    if ja_tuv is not None:
        ja_seg = ja_tuv.find('seg')
        if ja_seg is not None:
            source_text = ja_seg.text
            # print(f'Source: {source_text}')
            
    # Extracting English target segments if available
    en_tuv = tu.find('tuv[@xml:lang="en"]', namespaces=namespaces)
    if en_tuv is not None:
        en_seg = en_tuv.find('seg')
        if en_seg is not None:
            target_text = en_seg.text
            # print(f'Target: {target_text}')
    
    # Append to lists regardless of whether both source and target texts are found
    sources.append(source_text if source_text is not None else '')
    targets.append(target_text if target_text is not None else '')

# Check if we have successfully extracted any data
if sources and targets:
    # Create a DataFrame from the extracted data
    df = pd.DataFrame({
        'Source (JA)': sources,
        'Target (EN)': targets
    })

    # Save the DataFrame to a CSV file
    csv_file = 'output_files/output.csv'
    df.to_csv(csv_file, index=False)

    print(f"CSV file has been saved to {csv_file}.")
else:
    print("No data was extracted. Please check the TMX file structure and namespaces.")
