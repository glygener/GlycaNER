import os
import xml.etree.ElementTree as ET
import pandas as pd
import re
import csv
import time

base_dir = os.path.dirname(os.path.abspath(__file__))
STAGE1_DIR = os.environ.get("GLYCANER_STAGE1_DIR", "/data/output/bert_detection")

input_dir = os.path.join(STAGE1_DIR, 'abstract_xml')
output_text_dir = os.path.join(STAGE1_DIR, 'abstract_text')

# Create the output directory if it doesn't exist
if not os.path.exists(output_text_dir):
    os.makedirs(output_text_dir)

for filename in os.listdir(input_dir):
    if filename.endswith('.xml'):
        file_path = os.path.join(input_dir, filename)
        
        # Parse the XML file
        tree = ET.parse(file_path)
        root = tree.getroot()
        
        # Initialize variables to store the title and abstract text
        title_text = ""
        abstract_text = ""
        
        # Extract the PMID
        pmid_elem = root.find('.//document/id')
        pmid = pmid_elem.text if pmid_elem is not None else "Unknown_PMID"
        
        # Extract title and abstract text
        for passage in root.findall('.//passage'):
            passage_type_elem = passage.find('infon[@key="type"]')
            passage_type = passage_type_elem.text if passage_type_elem is not None else ""
            if passage_type == 'title':
                title_text = passage.find('text').text if passage.find('text') is not None else ""
            elif passage_type == 'abstract':
                abstract_text = passage.find('text').text if passage.find('text') is not None else ""
        
        # Concatenate the title and abstract text
        concatenated_text = title_text + " " + abstract_text
        
        # Save the concatenated text to a file
        output_text_file = os.path.join(output_text_dir, f"{pmid}.txt")
        with open(output_text_file, 'w') as text_file:
            text_file.write(concatenated_text)
        

print("Process completed successfully.")
