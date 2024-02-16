import xml.etree.ElementTree as ET
import pandas as pd
from config import tmx_filepath

# Define the path to your TMX file
tmx_file = tmx_filepath

# Parse the TMX file
tree = ET.parse(tmx_file)
root = tree.getroot()

# Extracting the namespace map from the TMX file's root element
# TMX files and other XML documents often define namespaces, which need to be accounted for when searching for elements
namespaces = {'xml': 'http://www.w3.org/XML/1998/namespace'}  # Common namespace for xml:lang attributes

# Initialize lists to hold the extracted data
sources = []
targets = []

# Loop through each translation unit in the TMX file
for tu in root.findall('body/tu'):
    source_text = ''
    target_text = ''
    
    # Extracting source and target segments based on xml:lang attribute
    for tuv in tu.findall('tuv[@xml:lang="ja"]', namespaces=namespaces):
        seg = tuv.find('seg')
        if seg is not None:
            source_text = seg.text
    
    for tuv in tu.findall('tuv[@xml:lang="en"]', namespaces=namespaces):
        seg = tuv.find('seg')
        if seg is not None:
            target_text = seg.text
    
    # Only add to lists if both source and target text are found
    if source_text and target_text:
        sources.append(source_text)
        targets.append(target_text)

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
