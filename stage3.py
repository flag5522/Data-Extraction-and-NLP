import os
import pandas as pd

def load_words(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return set(word.strip().lower() for word in file.readlines())
    except UnicodeDecodeError:
        with open(file_path, 'r', encoding='latin-1') as file:
            return set(word.strip().lower() for word in file.readlines())

def calculate_scores(text, positive_words, negative_words):
    words = text.split()
    total_words = len(words)
    positive_score = sum(1 for word in words if word.lower() in positive_words)
    negative_score = sum(1 for word in words if word.lower() in negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (total_words + 0.000001)
    return positive_score, negative_score, polarity_score, subjectivity_score

# Load positive and negative words
master_dictionary_dir = 'MasterDictionary'
positive_words = load_words(os.path.join(master_dictionary_dir, 'positive-words.txt'))
negative_words = load_words(os.path.join(master_dictionary_dir, 'negative-words.txt'))

# Directory containing cleaned text files
swdata_directory = 'SWdata'

# Initialize a list to store results
results = []

# Process each file in the SWdata directory
for filename in os.listdir(swdata_directory):
    file_path = os.path.join(swdata_directory, filename)
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
        positive_score, negative_score, polarity_score, subjectivity_score = calculate_scores(content, positive_words, negative_words)
        results.append({
            'File Name': filename,
            'Positive Score': positive_score,
            'Negative Score': negative_score,
            'Polarity Score': polarity_score,
            'Subjectivity Score': subjectivity_score
        })

# Convert results to a DataFrame and save as CSV
results_df = pd.DataFrame(results)
results_df.to_csv('data.csv', index=False)

print("Scores calculated and saved to data.csv.")
