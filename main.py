import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import re
import string
import nltk
from nltk.tokenize import sent_tokenize, word_tokenize
from nltk.corpus import stopwords, cmudict
import syllapy



def extract_text_from_url(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            paragraphs = soup.find_all('p')
            text = '\n'.join(p.get_text() for p in paragraphs)
            return text
        else:
            return None
    except Exception as e:
        print(f"Error fetching URL {url}: {e}")
        return None

def save_text_to_file(text, file_name, directory):
    file_path = os.path.join(directory, f"{file_name}.txt")
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(text)

def load_stop_words(stop_words_dir):
    stop_words_files = [
        'StopWords_Names.txt',
        'StopWords_Geographic.txt',
        'StopWords_GenericLong.txt',
        'StopWords_Generic.txt',
        'StopWords_DatesandNumbers.txt',
        'StopWords_Currencies.txt',
        'StopWords_Auditor.txt'
    ]
    stop_words = set()
    for file_name in stop_words_files:
        file_path = os.path.join(stop_words_dir, file_name)
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                for line in file:
                    stop_words.add(line.strip().lower())
        except UnicodeDecodeError:
            with open(file_path, 'r', encoding='latin-1') as file:
                for line in file:
                    stop_words.add(line.strip().lower())
    return stop_words

def remove_stop_words(text, stop_words):
    words = text.split()
    filtered_words = [word for word in words if word.lower() not in stop_words]
    return ' '.join(filtered_words)

if __name__ == "__main__":
    # Load the Excel file
    input_file_path = 'Input.xlsx'  # Update this path if necessary
    data = pd.read_excel(input_file_path)
    
    # Directories for input, output, and stop words
    input_directory = 'EXDT'
    output_directory = 'SWdata'
    stop_words_dir = 'StopWords'
    
    # Ensure the directories exist
    if not os.path.exists(input_directory):
        os.makedirs(input_directory)
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    
    # Extract and save text from URLs
    for _, row in data.iterrows():
        url_id = row['URL_ID']
        url = row['URL']
        extracted_text = extract_text_from_url(url)
        if extracted_text:
            save_text_to_file(extracted_text, url_id, input_directory)
            print(f"Text extracted and saved for {url_id}")
        else:
            print(f"Failed to extract text for {url_id}")
    
    # Load stop words
    stop_words = load_stop_words(stop_words_dir)
    
    # Remove stop words from files and save cleaned text
    for filename in os.listdir(input_directory):
        input_file_path = os.path.join(input_directory, filename)
        with open(input_file_path, 'r', encoding='utf-8') as file:
            content = file.read()
            cleaned_content = remove_stop_words(content, stop_words)
        
        output_file_path = os.path.join(output_directory, filename)
        with open(output_file_path, 'w', encoding='utf-8') as file:
            file.write(cleaned_content)
    
    print("Processed files and removed stop words.")

# Load the positive and negative words for sentiment analysis
def load_words(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return set(word.strip().lower() for word in file.readlines())
    except UnicodeDecodeError:
        with open(file_path, 'r', encoding='latin-1') as file:
            return set(word.strip().lower() for word in file.readlines())

# Calculate sentiment scores
def calculate_scores(text, positive_words, negative_words):
    words = text.split()
    total_words = len(words)
    positive_score = sum(1 for word in words if word.lower() in positive_words)
    negative_score = sum(1 for word in words if word.lower() in negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (total_words + 0.000001)
    return positive_score, negative_score, polarity_score, subjectivity_score

# Download necessary NLTK resources
nltk.download('punkt')
nltk.download('stopwords')
nltk.download('cmudict')

# Set up stopwords, punctuation, and CMU Pronouncing Dictionary
stop_words = set(stopwords.words('english'))
punctuation_table = str.maketrans('', '', string.punctuation)
d = cmudict.dict()  # CMU Pronouncing Dictionary

# Function to count syllables in a word
def count_syllables(word):
    if word.lower() in d:
        return max([len(list(y for y in x if y[-1].isdigit())) for x in d[word.lower()]])
    else:
        return syllapy.count(word)

def calculate_avg_words_per_sentence(text):
    sentences = sent_tokenize(text)
    words = word_tokenize(text)
    num_sentences = len(sentences)
    num_words = len(words)
    avg_words_per_sentence = num_words / num_sentences if num_sentences else 0
    return avg_words_per_sentence

# Check if a word is complex (more than 2 syllables)
def is_complex_word(word):
    return count_syllables(word) > 2

# Clean text and count words excluding stopwords
def clean_and_count_words(text):
    words = word_tokenize(text)
    cleaned_words = [word.translate(punctuation_table).lower() for word in words if word.lower() not in stop_words and word.isalpha()]
    return len(cleaned_words)

# Calculate readability metrics
def calculate_metrics(text):
    sentences = sent_tokenize(text)
    num_sentences = len(sentences)
    words = word_tokenize(text)
    cleaned_word_count = clean_and_count_words(text)
    complex_words_count = sum(is_complex_word(word) for word in words)
    total_syllables = sum(count_syllables(word) for word in words)
    personal_pronouns_count = len(re.findall(r'\b(I|we|my|ours|us)\b', text, re.IGNORECASE)) - len(re.findall(r'\bUS\b', text))
    total_characters = sum(len(word) for word in words if word.isalpha())
    avg_words_per_sentence = calculate_avg_words_per_sentence(text)
    avg_sentence_length = cleaned_word_count / num_sentences if num_sentences else 0
    percentage_complex_words = (complex_words_count / cleaned_word_count * 100) if cleaned_word_count else 0
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    avg_word_length = total_characters / cleaned_word_count if cleaned_word_count else 0

    return {
        'WORD COUNT': cleaned_word_count,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sentence,
        'AVG SENTENCE LENGTH': avg_sentence_length,
        'PERCENTAGE OF COMPLEX WORDS': percentage_complex_words,
        'FOG INDEX': fog_index,
        'COMPLEX WORD COUNT': complex_words_count,
        'SYLLABLE PER WORD': total_syllables,
        'PERSONAL PRONOUNS': personal_pronouns_count,
        'AVG WORD LENGTH': avg_word_length
    }

# Load positive and negative words
master_dictionary_dir = 'MasterDictionary'
positive_words = load_words(os.path.join(master_dictionary_dir, 'positive-words.txt'))
negative_words = load_words(os.path.join(master_dictionary_dir, 'negative-words.txt'))

input_directory = 'SWdata'
results = []

# Process each file for both sentiment and readability analysis
for filename in os.listdir(input_directory):
    file_path = os.path.join(input_directory, filename)
    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()
        positive_score, negative_score, polarity_score, subjectivity_score = calculate_scores(text, positive_words, negative_words)
        metrics = calculate_metrics(text)
        metrics.update({
            'URL_ID': filename,
            'URL' : url,
            'POSITIVE SCORE': positive_score,
            'NEGATIVE SCORE': negative_score,
            'POLARITY SCORE': polarity_score,
            'SUBJECTIVITY SCORE': subjectivity_score
        })
        results.append(metrics)

# Convert results to a DataFrame
results_df = pd.DataFrame(results)

# Specify the column order explicitly and reorder the DataFrame
column_order = [
    'URL_ID', 'URL', 'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE',
    'SUBJECTIVITY SCORE', 'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS',
    'FOG INDEX', 'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT',
    'WORD COUNT', 'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH'
]


results_df = results_df[column_order]

# Save the reordered DataFrame to an Excel file
results_df.to_excel('Output.xlsx', index=False)

print("Analysis completed and saved to Output.xlsx.")
