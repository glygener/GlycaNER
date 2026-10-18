import os
import pandas as pd
import regex as re

base_dir = os.path.dirname(os.path.abspath(__file__))
STAGE1_DIR = os.environ.get("GLYCANER_STAGE1_DIR", "/data/output/bert_detection")
output_text_dir = os.path.join(STAGE1_DIR, 'abstract_text')
excel_sentences = os.path.join(STAGE1_DIR, 'sentences.xlsx')
sentences_data = []

def get_sentences(text):
    # Enhanced regex to handle sentence boundaries and exceptions like 'e.g.'
    pattern = re.compile(
        r"""
        (?<!\b(?:Mr|Dr|Ms|Mrs|St|vs|fig|etc|i\.e|cf|Jr|Prof|Sr)\.)   # Avoid splitting after abbreviations
        (?<!\b[A-Z]\.)                                           # Avoid splitting after initials like 'P.'
        (?<!e\.g\.)                                              # Prevent splitting after 'e.g.'
        (?<!e\.g\.\s\(.*?\))                                     # Prevent splitting if 'e.g.' is followed by parentheses
        (?<!e\.g\.\s(?=[a-z]))                                   # Prevent splitting after 'e.g.' followed by lowercase
        \.\s(?=[A-Z])                                            # Match period-space-uppercase for sentence boundary
        """,
        re.VERBOSE  # Enables multi-line regex with comments
    )
    # Split the text using the pattern
    sentences = [sentence.strip() for sentence in pattern.split(text) if sentence.strip()]
    # Ensure each sentence ends with a period
    return [sentence + '.' if not sentence.endswith('.') else sentence for sentence in sentences]

def process_text(file_path, pmid):
    with open(file_path, 'r') as file:
        text = file.read()
        sentences = get_sentences(text)
        for i, sentence in enumerate(sentences):
            sentences_data.append({
                'pmid': pmid,
                'sent_index': i,
                'sent_text': sentence,
                'charStart': text.find(sentence),
                'charEnd': text.find(sentence) + len(sentence)
            })

def save_to_excel():
    sentences_df = pd.DataFrame(sentences_data)
    sentences_df.drop_duplicates(subset=['pmid', 'sent_index', 'sent_text', 'charStart', 'charEnd'], inplace=True)
    sentences_df.to_excel(excel_sentences, index=False)

# Process each text file to extract sentences
for pmid_file in os.listdir(output_text_dir):
    if pmid_file.endswith('.txt'):
        pmid = pmid_file.split('.')[0]
        file_path = os.path.join(output_text_dir, pmid_file)
        process_text(file_path, pmid)

save_to_excel()
print("Sentence extraction completed successfully.")
