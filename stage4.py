import os
import re
import pandas as pd
import string
from nltk.tokenize import sent_tokenize, word_tokenize
from nltk.corpus import stopwords, cmudict
import nltk
import syllapy

# Download necessary NLTK resources
nltk.download('punkt')
nltk.download('stopwords')
nltk.download('cmudict')

# Stopwords and punctuation setup
stop_words = set(stopwords.words('english'))
punctuation_table = str.maketrans('', '', string.punctuation)

# CMU Pronouncing Dictionary for syllable counting
d = cmudict.dict()

def count_syllables(word):
    if word.lower() in d:
        return max([len(list(y for y in x if y[-1].isdigit())) for x in d[word.lower()]])
    else:
        return syllapy.count(word)

def is_complex_word(word):
    return count_syllables(word) > 2

def clean_and_count_words(text):
    words = word_tokenize(text)
    cleaned_words = [word.translate(punctuation_table).lower() for word in words if word.lower() not in stop_words and word.isalpha()]
    return len(cleaned_words)

def calculate_metrics(text):
    sentences = sent_tokenize(text)
    num_sentences = len(sentences)
    words = word_tokenize(text)  # Correctly define words here
    cleaned_word_count = clean_and_count_words(text)
    complex_words_count = sum(is_complex_word(word) for word in words)
    total_syllables = sum(count_syllables(word) for word in words)
    personal_pronouns_count = len(re.findall(r'\b(I|we|my|ours|us)\b', text, re.IGNORECASE)) - len(re.findall(r'\bUS\b', text))
    total_characters = sum(len(word) for word in words if word.isalpha())
    
    avg_sentence_length = cleaned_word_count / num_sentences if num_sentences else 0
    percentage_complex_words = (complex_words_count / cleaned_word_count * 100) if cleaned_word_count else 0
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    avg_word_length = total_characters / cleaned_word_count if cleaned_word_count else 0

    return {
        'File Name': '',
        'Word Count': cleaned_word_count,
        'Average Sentence Length': avg_sentence_length,
        'Average Number of Words Per Sentence': avg_sentence_length,  # Explicitly included
        'Percentage of Complex Words': percentage_complex_words,
        'Fog Index': fog_index,
        'Complex Word Count': complex_words_count,
        'Total Syllables': total_syllables,
        'Personal Pronouns Count': personal_pronouns_count,
        'Average Word Length': avg_word_length
    }

input_directory = 'SWdata'
results = []

# Process each file
for filename in os.listdir(input_directory):
    file_path = os.path.join(input_directory, filename)
    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()
        metrics = calculate_metrics(text)
        metrics['File Name'] = filename
        results.append(metrics)

# Save results to CSV
results_df = pd.DataFrame(results)
results_df.to_csv('output_analysis.csv', index=False)

print("Analysis completed and saved to EXDT_analysis.csv.")
