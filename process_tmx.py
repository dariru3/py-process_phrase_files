import xml.etree.ElementTree as ET
import pandas as pd
from config import tmx_filepath

# Define the path to your TMX file
tmx_file = tmx_filepath

# Parse the TMX file
tree = ET.parse(tmx_file)
root = tree.getroot()

# Initialize lists to hold the extracted data
sources = []
targets = []

# Loop through each translation unit in the TMX file
for tu in root.findall('.//tu'):
    source_text = None
    target_text = None
    
    for tuv in tu.findall('.//tuv'):
        lang = tuv.get('{http://www.w3.org/XML/1998/namespace}lang')
        seg_text = tuv.find('seg').text
        if lang == 'EN':
            source_text = seg_text
        elif lang == 'JA':
            target_text = seg_text
    
    # Append the source and target texts to their respective lists
    if source_text is not None and target_text is not None:
        sources.append(source_text)
        targets.append(target_text)

# Create a DataFrame from the extracted data
df = pd.DataFrame({
    'Source (EN)': sources,
    'Target (JA)': targets
})

# Save the DataFrame to a CSV file
csv_file = 'output_files/output.csv'
df.to_csv(csv_file, index=False)

print(f"CSV file has been saved to {csv_file}.")
