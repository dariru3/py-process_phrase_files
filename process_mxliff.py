import xml.etree.ElementTree as ET
import pandas as pd
from config import mxliff_filepath

# Assuming 'mxliff_filepath' is a variable containing the path to your MXLIFF file
mxliff_file = mxliff_filepath

# Register the namespace to properly handle prefixed attributes
ET.register_namespace('m', 'urn:oasis:names:tc:xliff:document:1.2')

# Parse the MXLIFF file
tree = ET.parse(mxliff_file)
root = tree.getroot()

# Define the namespaces used in your MXLIFF file
namespaces = {
    'm': 'urn:oasis:names:tc:xliff:document:1.2',  # Adjust based on your file's namespaces
}

# Initialize lists to hold the extracted data
sources = []
targets = []
match_qualities = []

# Loop through each translation unit in the MXLIFF file
for trans_unit in list(root.findall('.//m:trans-unit', namespaces))[:5]:
    source_text = trans_unit.find('.//source', namespaces).text if trans_unit.find('.//source', namespaces) is not None else ''
    print(f'Source text: {source_text}')
    target_text = trans_unit.find('.//target', namespaces).text if trans_unit.find('.//target', namespaces) is not None else ''
    print(f'Target text: {target_text}')

    # Initialize match quality as 'N/A' or another suitable default value
    match_quality = 'N/A'
    
    # Check for alt-trans elements with origin="memsource-tm" and extract match-quality
    for alt_trans in trans_unit.findall('.//m:alt-trans', namespaces):
        if alt_trans.attrib.get('origin') == 'memsource-tm':
            match_quality = alt_trans.attrib.get('match-quality', 'N/A')
            print(f'Match: {match_quality}')
            break  # Assuming we only need the first matching alt-trans entry

    sources.append(source_text)
    targets.append(target_text)
    match_qualities.append(match_quality)

# Create a DataFrame from the extracted data
df = pd.DataFrame({
    'Source': sources,
    'Target': targets,
    'Match Quality (Memsource-TM)': match_qualities
})

# Save the DataFrame to a CSV file
csv_file = 'output_files/mxliff_output.csv'
df.to_csv(csv_file, index=False)

print(f"CSV file has been saved to {csv_file}.")
