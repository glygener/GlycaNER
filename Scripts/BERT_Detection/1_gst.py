import torch
from transformers import AutoTokenizer, AutoModelForTokenClassification
import pandas as pd
import os

# Load your tokenizer, model, id2label dictionary
# Ensure the tokenizer and model are properly loaded before running this script
base_dir = os.path.dirname(os.path.abspath(__file__))
STAGE1_DIR = os.environ.get("GLYCANER_STAGE1_DIR", "/data/output/bert_detection")

model_path = os.environ.get("GLYCANER_MODEL_DIR", "/app/model")
tokenizer = AutoTokenizer.from_pretrained(model_path)
model = AutoModelForTokenClassification.from_pretrained(model_path)
id2label = model.config.id2label

print("Model loaded successfully!")

# Define file paths
input_excel_file = os.path.join(STAGE1_DIR, 'sentences.xlsx')  # Input file with sentences
output_csv_file = os.path.join(STAGE1_DIR, 'output_predictions.csv')  # Output file with predictions

# Load the CSV file
df = pd.read_excel(input_excel_file)

# Check if a column named "sentence" exists
if "sent_text" not in df.columns:
    raise ValueError("The input CSV must contain a column named 'sentence'.")

# Initialize a list to store predictions
predictions_list = []

# Process each sentence in the CSV
for custom_sentence in df["sent_text"]:
    # Tokenize the sentence
    tokenized_inputs = tokenizer(
        custom_sentence.split(),
        is_split_into_words=True,
        return_tensors="pt",
        truncation=True,
        padding="max_length",
        max_length=128
    )

    # Retrieve word_ids separately
    word_ids = tokenizer(custom_sentence.split(), is_split_into_words=True).word_ids()

    # Move inputs to the same device as the model
    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    model = model.to(device)
    tokenized_inputs = {key: val.to(device) for key, val in tokenized_inputs.items()}

    # Get predictions
    with torch.no_grad():
        outputs = model(**tokenized_inputs)
        logits = outputs["logits"]

    # Convert logits to predicted label indices
    predicted_label_indices = torch.argmax(logits, dim=2).squeeze().tolist()

    # Decode predictions
    tokens = tokenizer.convert_ids_to_tokens(tokenized_inputs["input_ids"].squeeze().tolist())
    predicted_labels = [id2label[idx] for idx in predicted_label_indices]

    # Filter out special tokens ([CLS], [SEP], [PAD]) and merge subtokens based on word_ids
    final_tokens = []
    final_labels = []
    previous_word_id = -1
    merged_token = ""
    merged_label = None

    for token, label, word_id in zip(tokens, predicted_labels, word_ids):
        if word_id is None or token in ["[CLS]", "[SEP]", "[PAD]"]:
            continue  # Skip special tokens
        if word_id != previous_word_id:
            if merged_token:
                # Append the merged token and its label
                final_tokens.append(merged_token)
                final_labels.append(merged_label)
            # Start a new token
            merged_token = token
            merged_label = label
        else:
            # Merge subtokens
            merged_token += token[2:] if token.startswith("##") else token
        previous_word_id = word_id

    # Add the last merged token
    if merged_token:
        final_tokens.append(merged_token)
        final_labels.append(merged_label)

    # Prepare the prediction as a string in the format token1(label1), token2(label2)
    predictions = ", ".join(f"{token}({label})" for token, label in zip(final_tokens, final_labels))
    predictions_list.append(predictions)

# Add the predictions as a new column in the DataFrame
df["predicted_labels"] = predictions_list

# Save the updated DataFrame to a new CSV file
df.to_csv(output_csv_file, index=False)

print(f"Predictions have been saved to {output_csv_file}.")
