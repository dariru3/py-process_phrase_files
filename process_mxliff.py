import xml.etree.ElementTree as ET
import pandas as pd

def parse_mxliff_to_df(mxliff_file): 
    # Register the namespace to properly handle prefixed attributes
    ET.register_namespace('m', 'urn:oasis:names:tc:xliff:document:1.2')

    # Parse the MXLIFF file
    tree = ET.parse(mxliff_file)
    root = tree.getroot()

    # Define the namespaces used in your MXLIFF file
    namespaces = {'m': 'urn:oasis:names:tc:xliff:document:1.2'}

    # Initialize lists to hold the extracted data
    sources = []
    targets = []
    match_qualities = []

    # Loop through each translation unit in the MXLIFF file
    for trans_unit in root.findall('.//m:trans-unit', namespaces):
        source_text = trans_unit.find('m:source', namespaces).text if trans_unit.find('m:source', namespaces) is not None else ''
        target_text = trans_unit.find('m:target', namespaces).text if trans_unit.find('m:target', namespaces) is not None else ''
        match_quality = '0' # Default value
        
        # Check for alt-trans elements with origin="memsource-tm" and extract match-quality
        for alt_trans in trans_unit.findall('.//m:alt-trans', namespaces):
            if alt_trans.attrib.get('origin') == 'memsource-tm':
                match_quality = alt_trans.attrib.get('match-quality', '0')
                match_quality = int(float(match_quality) * 100)
                break  # Assuming we only need the first matching alt-trans entry

        sources.append(source_text)
        targets.append(target_text)
        match_qualities.append(match_quality)

    # Create a DataFrame from the extracted data
    df = pd.DataFrame({
        'Source': sources,
        'Target': targets,
        'Match': match_qualities
    })

    df.index += 1
    df.reset_index(inplace=True)
    df.rename(columns={'index': 'Index'}, inplace=True)

    df['Comment'] = ""

    # View parsed data in CSV file
    csv_file = 'output_files/mxliff_to_csv_preview.csv'
    df.to_csv(csv_file, index=False)
    print(f"CSV file has been saved to {csv_file}.")

    return df

if __name__ == "__main__":
    filepath = "input_files/240226_良品計画様_統合報告2023対訳表_P37-40+-ja-en-T.mxliff"
    parse_mxliff_to_df(filepath)