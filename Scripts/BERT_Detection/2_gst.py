import pandas as pd
import os

# Load the uploaded CSV file
base_dir = os.path.dirname(os.path.abspath(__file__))
STAGE1_DIR = os.environ.get("GLYCANER_STAGE1_DIR", "/data/output/bert_detection")
file_path = os.path.join(STAGE1_DIR, 'output_predictions.csv') 
df = pd.read_csv(file_path)

# Add the GST_BERT and I-GST_Start columns
gst_bert_list = []
igst_start_list = []

for index, row in df.iterrows():
    tokens_labels = row["predicted_labels"].split(", ")  # Split token(label) pairs
    gst_terms = []
    igst_start_terms = []
    
    current_gst_term = []

    for pair in tokens_labels:
        # Check if the pair has the expected format
        if "(" in pair and pair.endswith(")"):
            token, label = pair.rsplit("(", 1)
            label = label.rstrip(")")
        else:
            # Skip invalid or improperly formatted pairs
            continue
        
        if label == "B-GST":
            if current_gst_term:  # End the current GST term if it exists
                gst_terms.append(" ".join(current_gst_term))
                current_gst_term = []
            current_gst_term.append(token)
        
        elif label == "I-GST":
            if not current_gst_term:
                igst_start_terms.append(token)  # Collect stray I-GST tokens
            else:
                current_gst_term.append(token)
        
        elif label == "O-GST":
            if current_gst_term:
                gst_terms.append(" ".join(current_gst_term))
                current_gst_term = []

    # Add any remaining GST term
    if current_gst_term:
        gst_terms.append(" ".join(current_gst_term))
    
    gst_bert_list.append(" | ".join(gst_terms))
    igst_start_list.append(" ".join(igst_start_terms))

# Add the new columns to the DataFrame
df["GST_BERT"] = gst_bert_list
df["I-GST_Start"] = igst_start_list

# Save the updated DataFrame to a new CSV file
output_file_path = os.path.join(STAGE1_DIR, 'output_predictions1.csv') 
df.to_csv(output_file_path, index=False)

print(f"Updated DataFrame has been saved to {output_file_path}")
