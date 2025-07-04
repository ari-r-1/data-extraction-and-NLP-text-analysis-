### 🧠 Data Extraction And NLP Analysis

## 📌 Objective

This project automates the extraction and analysis of article content from a list of URLs provided in Input.xlsx. It processes the text and calculates key sentiment and readability metrics. The results are exported to an Excel file, structured as per Output Data Structure.xlsx.

---

## 🧾 Input Files

Please ensure the following files are available in the same directory before running the script:

- `Input.xlsx` – List of articles with `URL_ID` and `URL`
- `positive-words.txt` – Positive words list (from MasterDictionary)
- `negative-words.txt` – Negative words list (from MasterDictionary)
- `StopWords/` – A folder containing all stopword `.txt` files such as:
  - `StopWords_Generic.txt`

---

## Dependencies

You must have Python 3.x and the following packages installed:

## Required core packages

pandas              # For reading/writing Excel and handling data
openpyxl            # For reading/writing .xlsx files with pandas
nltk                # For tokenization and stopword handling
beautifulsoup4      # For parsing HTML content from web pages
requests            # For sending HTTP requests to URLs


import nltk
nltk.download('punkt')


```bash
pip install pandas openpyxl nltk beautifulsoup4 requests

---

## 📊 Output

The output Excel file will follow the format defined in Output Data Structure.xlsx, containing:
- Cleaned text
- Sentiment scores (positive, negative, polarity, subjectivity)
- Readability scores (syllables, average sentence length, FOG index)
- Word and sentence counts
- Complex word analysis

---
