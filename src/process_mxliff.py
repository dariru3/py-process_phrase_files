import xml.etree.ElementTree as ET
import pandas as pd
import re
from .config_loader import CONFIG

def remove_tags(text):
    patterns = CONFIG["ProcessingXliffSettings"]["TagPatterns"]
    cleansed_text = re.sub(patterns, '', text)

    return cleansed_text

def setup_root(mxliff_file, xliff_namespace):
    """
    Register the given namespace and parse the MXLIFF file,
    returning the root element.
    """
    ET.register_namespace('m', xliff_namespace)
    tree = ET.parse(mxliff_file)
    return tree.getroot()

def get_match_quality(alt_trans):
    """
    Extract and calculate the match quality from an alt-trans element.
    Returns the match quality as an integer percentage.
    """
    if alt_trans.attrib.get('origin') == 'memsource-tm':
        match_quality = alt_trans.attrib.get('match-quality', '0')
        return int(float(match_quality) * 100)
    else:
        return 0

def parse_mxliff_to_df(mxliff_file):
    print("Processing .MXLIFF file...")
    xliff_namespace = CONFIG["ProcessingXliffSettings"]["XliffNamespace"]
    root = setup_root(mxliff_file, xliff_namespace)

    # Define the namespaces used in your MXLIFF file
    namespaces = {'m': xliff_namespace}

    # Initialize lists to hold the extracted data
    ids = []
    sources = []
    targets = []
    match_qualities = []

    # Loop through each translation unit in the MXLIFF file
    find_text = lambda trans_unit, m : trans_unit.find(m, namespaces).text if trans_unit.find(m, namespaces) is not None else ''

    for trans_unit in root.findall('.//m:trans-unit', namespaces):
        unit_id = trans_unit.attrib.get('id', '')
        source_text = find_text(trans_unit, 'm:source')
        source_text = remove_tags(source_text)
        target_text = find_text(trans_unit, 'm:target')
        # target_text = remove_tags(target_text)

        # Check for alt-trans elements with origin="memsource-tm" and extract match-quality
        match_quality = 0 # Default value
        for alt_trans in trans_unit.findall('.//m:alt-trans', namespaces):
            match_quality = get_match_quality(alt_trans)
            if match_quality != 0:
                break  # Assuming we only need the first matching alt-trans entry

        ids.append(unit_id)
        sources.append(source_text)
        targets.append(target_text)
        match_qualities.append(match_quality)

    # Create a DataFrame from the extracted data
    df = pd.DataFrame({
        'ID': ids,
        'Source': sources,
        'Target': targets,
        'Match': match_qualities
    })

    df['Comment'] = ""

    return df
