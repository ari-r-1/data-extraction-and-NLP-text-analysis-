# Import main requirements
import nltk, shutil
shutil.rmtree('/root/nltk_data', ignore_errors=True)
nltk.download('punkt')
nltk.download('stopwords')
import os
import re
import pandas as pd
import requests
from bs4 import BeautifulSoup
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
import time

# Force clean punkt download if corrupted
punkt_path = os.path.join(nltk.data.find("tokenizers"), "punkt")
if os.path.exists(punkt_path):
    shutil.rmtree(punkt_path)

nltk.download('punkt', force=True)
nltk.download('stopwords', force=True)
folder_path = os.path.expanduser('~/nltk_data/tokenizers/punkt')
if os.path.exists(folder_path):
    shutil.rmtree(folder_path)
    nltk.download('punkt')

# File upload
def load_wordlist(filepath):
    with open(filepath, 'r', encoding='latin-1') as f:
        return set(line.strip() for line in f if line.strip() and not line.startswith(';'))

# Ensure all NLTK resources are downloaded
def download_nltk_resources():
    resources = {
        'punkt': ['punkt', 'punkt_tab'],
        'stopwords': ['stopwords'],
        'wordnet': ['wordnet'],
        'averaged_perceptron_tagger': ['averaged_perceptron_tagger']
    }

    for resource, packages in resources.items():
        try:
            for package in packages:
                try:
                    nltk.data.find(f'tokenizers/{package}')
                except LookupError:
                    print(f"Downloading NLTK {package}...")
                    nltk.download(package)
        except Exception as e:
            print(f"Error downloading NLTK {resource}: {e}")

# Download required NLTK data
download_nltk_resources()

def load_wordlist(filepath):
    """Load word list from file, handling encoding issues."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            words = set(line.strip() for line in f if line.strip() and not line.startswith(';'))
        return words
    except UnicodeDecodeError:
        with open(filepath, 'r', encoding='latin-1') as f:
            words = set(line.strip() for line in f if line.strip() and not line.startswith(';'))
        return words

# Load input file
try:
    df = pd.read_excel("Input.xlsx")
except FileNotFoundError:
    print("Error: Input.xlsx file not found.")
    exit()
except Exception as e:
    print(f"Error loading Excel file: {e}")
    exit()

# Validate required columns
if not all(col in df.columns for col in ['URL_ID', 'URL']):
    print("Error: Input.xlsx must contain 'URL_ID' and 'URL' columns.")
    exit()

# Load positive and negative words
try:
    positive_words = load_wordlist("positive-words.txt")
    negative_words = load_wordlist("negative-words.txt")
except FileNotFoundError as e:
    print(f"Error loading word lists: {e}")
    exit()

# Load stopwords from StopWords folder
try:
    stop_words = set(stopwords.words("english"))
except LookupError:
    print("Downloading NLTK stopwords...")
    nltk.download('stopwords')
    stop_words = set(stopwords.words("english"))

# Helper functions
def count_syllables(word):
    """Count syllables in a word with improved accuracy."""
    word = word.lower().strip(".:;?!")
    if not word:
        return 0

    # Handle common exceptions
    if word.endswith('es') or word.endswith('ed'):
        word = word[:-2]
    if word.endswith('le') and len(word) > 2 and word[-3] not in 'aeiouy':
        pass

    vowels = "aeiouy"
    count = 0
    prev_char_was_vowel = False

    # Count vowel groups
    for char in word:
        if char in vowels:
            if not prev_char_was_vowel:
                count += 1
            prev_char_was_vowel = True
        else:
            prev_char_was_vowel = False

    # Adjust for words ending with 'e'
    if word.endswith('e') and count > 1:
        count -= 1

    return count if count > 0 else 1

def count_pronouns(text):
    """Count personal pronouns with improved regex."""
    pronouns = re.findall(r'\b(I|we|my|ours|us)\b', text, flags=re.I)
    # Exclude cases where "US" might refer to the country
    filtered = [p for p in pronouns if (p.lower() != 'us' or
               (p == 'us' and not any(word.lower() in ['united', 'states', 'u.s.']
                for word in text.split(maxsplit=10))))]
    return len(filtered)

def clean_text(text):
    """Clean and normalize text."""
    # Remove extra whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    # Remove unwanted characters
    text = re.sub(r'[^\w\s.,;:!?\'"-]', '', text)
    return text

# Configure request headers to mimic a real browser
# Function to scrape article content
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
    'Upgrade-Insecure-Requests': '1',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'none',
    'Sec-Fetch-User': '?1',
    'Cache-Control': 'max-age=0',
}

results = []
for _, row in df.iterrows():
    try:
        url_id, url = row['URL_ID'], row['URL']

        # Fetch URL content with error handling
        try:
            print(f"Fetching {url_id}...")
            response = requests.get(url, headers=headers, timeout=15)

            # Check for 406 error specifically
            if response.status_code == 406:
                # Try alternative headers
                alt_headers = headers.copy()
                alt_headers['User-Agent'] = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 Safari/605.1.15'
                response = requests.get(url, headers=alt_headers, timeout=15)

            response.raise_for_status()

        except requests.exceptions.RequestException as e:
            print(f"Failed to fetch {url_id}: {e}")
            results.append([url_id, url] + [None] * 13)
            continue

        # Parse content only if we got HTML
        if 'text/html' not in response.headers.get('Content-Type', ''):
            print(f"Non-HTML content for {url_id}")
            results.append([url_id, url] + [None] * 13)
            continue

        soup = BeautifulSoup(response.content, "html.parser")

        # Extract text from all paragraph tags
        paras = soup.find_all('p')
        text = " ".join(p.get_text().strip() for p in paras if p.get_text().strip())
        text = clean_text(text)

        if not text:
            print(f"No text extracted from {url_id}")
            results.append([url_id, url] + [None] * 13)
            continue

        # Tokenization and analysis
        try:
            sentences = sent_tokenize(text)
        except LookupError:
            print("Downloading NLTK punkt...")
            nltk.download('punkt')
            sentences = sent_tokenize(text)

        total_sentences = len(sentences)

        tokens = word_tokenize(text)
        words = [w for w in tokens if w.isalpha() and w.lower() not in stop_words]
        total_words = len(words)

        # Function to analyze text
        # Sentiment analysis
        pos_score = sum(1 for w in words if w.lower() in positive_words)
        neg_score = sum(1 for w in words if w.lower() in negative_words)
        polarity = (pos_score - neg_score) / ((pos_score + neg_score) + 1e-10)
        subjectivity = (pos_score + neg_score) / (total_words + 1e-10)

        # Readability metrics
        avg_sent_len = total_words / total_sentences if total_sentences else 0
        complex_words = [w for w in words if count_syllables(w) > 2]
        pct_complex = len(complex_words) / total_words if total_words else 0
        fog_index = 0.4 * (avg_sent_len + pct_complex)

        # Word metrics
        syll_per_word = sum(count_syllables(w) for w in words) / total_words if total_words else 0
        avg_word_len = sum(len(w) for w in words) / total_words if total_words else 0

        # Personal pronouns
        pronoun_count = count_pronouns(text)

        results.append([
            url_id, url, pos_score, neg_score, polarity, subjectivity,
            avg_sent_len, pct_complex * 100, fog_index, avg_sent_len, len(complex_words),
            total_words, syll_per_word, pronoun_count, avg_word_len
        ])

        # Add delay between requests to be polite
        time.sleep(2)

    except Exception as e:
        print(f"Error processing {row['URL_ID']}: {e}")
        results.append([url_id, url] + [None] * 13)

# Process all URLs and save results
cols = ['URL_ID', 'URL', 'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE',
        'SUBJECTIVITY SCORE', 'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS',
        'FOG INDEX', 'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT',
        'WORD COUNT', 'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH']

# Save to Excel
try:
    out_df = pd.DataFrame(results, columns=cols)
    out_df.to_excel("Output.xlsx", index=False)
    print("Done! File saved as Output.xlsx")
except Exception as e:
    print(f"Error saving results: {e}")

# Download output file
df = pd.DataFrame(cols)
output_path = "Output.xlsx"
df.to_excel(output_path, index=False)
print(f"File saved to {output_path}")