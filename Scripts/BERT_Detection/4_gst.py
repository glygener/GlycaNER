import pandas as pd
import re
import os

# Load the data from the Excel file
base_dir = os.path.dirname(os.path.abspath(__file__))
STAGE1_DIR = os.environ.get("GLYCANER_STAGE1_DIR", "/data/output/bert_detection")
file_path = os.path.join(STAGE1_DIR, 'final_gst1.xlsx')
data = pd.read_excel(file_path)

# Prepare HTML output with CSS styling
html_content = '''<html><head>
<style>
    .highlight { color: blue; font-weight: bold; }
    body { font-family: Arial, sans-serif; }
    h1 { text-align: center; }
    h2 { color: darkred; }
    .sentence { margin-bottom: 10px; }
</style>
</head><body>
<h1>Highlighted GST ML Predictions for Abstracts</h1>'''

# Group data by pmid and iterate
for pmid, group in data.groupby('pmid'):
    html_content += f'<h2>PMID: {pmid}</h2>'
    
    # Iterate through each sentence in this pmid
    for _, row in group.iterrows():
        sent_index = row['sent_index']
        sent_text = row['sent_text']
        predictions = str(row['BERT_Output']).split()  # Assuming space-separated terms

        # Sort predictions by length in descending order to ensure longer terms are highlighted first
        predictions = sorted(predictions, key=len, reverse=True)

        # Highlight each prediction in the sent_text
        for prediction in predictions:
            pattern = r'(?<!\w){}(?!\w)'.format(re.escape(prediction))
            replacement = f'<span class="highlight">{prediction}</span>'
            sent_text = re.sub(pattern, replacement, sent_text, flags=re.IGNORECASE)

        # Add formatted text to HTML output
        html_content += f'<p class="sentence"><b>Sentence {sent_index}:</b> {sent_text}</p>'

html_content += '</body></html>'

# Save the HTML content to a file
html_file_path = os.path.join(STAGE1_DIR, 'Highlighted_ML_Predictions.html')
with open(html_file_path, 'w') as file:
    file.write(html_content)

print(f"HTML file created successfully: {html_file_path}")
