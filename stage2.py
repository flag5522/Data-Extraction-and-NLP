import requests
from bs4 import BeautifulSoup
import pandas as pd
import os

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
