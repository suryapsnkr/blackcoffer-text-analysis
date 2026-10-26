# Blackcoffer Data Extraction and NLP Text Analysis

An automated end-to-end Python pipeline for web crawling and NLP text analysis, developed for the Blackcoffer text analysis assignment.

---

## 📌 Project Overview
This project performs two core tasks:
1. **Data Extraction:** Crawls article title and body text from URLs given in `Input.xlsx`, filtering out website headers, footers, navigation, sidebars, and ads, and saves each article into `extracted_articles/<URL_ID>.txt`.
2. **AI Text Analysis:** Computes 13 linguistic, sentimental, and readability metrics as defined in `Text Analysis.docx`, compiling the output into `Output Data Structure.xlsx` (and `Output Data Structure.csv`).

---

## 📊 Output Variables Computed
1. **POSITIVE SCORE**
2. **NEGATIVE SCORE**
3. **POLARITY SCORE**
4. **SUBJECTIVITY SCORE**
5. **AVG SENTENCE LENGTH**
6. **PERCENTAGE OF COMPLEX WORDS**
7. **FOG INDEX**
8. **AVG NUMBER OF WORDS PER SENTENCE**
9. **COMPLEX WORD COUNT**
10. **WORD COUNT**
11. **SYLLABLE PER WORD**
12. **PERSONAL PRONOUNS**
13. **AVG WORD LENGTH**

---

## 📁 Repository Structure
```
├── Input.xlsx                     # Input URLs and URL IDs (82 articles)
├── Output Data Structure.xlsx     # Final computed metrics output (Excel)
├── Output Data Structure.csv      # Final computed metrics output (CSV)
├── main.py                        # Primary Python execution pipeline
├── requirements.txt               # Dependencies list
├── Instructions.md                # Detailed methodology and formula guide
├── Instructions.docx              # Word document format of instructions
├── Objective.docx                 # Original assignment requirements
├── Text Analysis.docx             # Original metric definitions
├── extracted_articles/            # Directory of 82 extracted .txt files
├── MasterDictionary/              # Positive and negative word dictionaries
└── StopWords/                     # Stop words lists (Currencies, Names, etc.)
```

---

## 🚀 Setup & Execution

### 1. Prerequisites
- Python 3.8 or higher
- Git

### 2. Installation
Clone the repository and install required packages:
```bash
git clone <YOUR_REPO_URL>
cd "20260915 Test Assignment"
pip install -r requirements.txt
```

### 3. Run Pipeline
```bash
python main.py
```
The script will concurrently extract article text for all URLs, compute the 13 variables, save each article to `extracted_articles/<URL_ID>.txt`, and export the results to `Output Data Structure.xlsx`.

---

## 🛠️ Technology Stack
- **Python 3**
- **Requests & BeautifulSoup4:** Web crawling and clean DOM extraction
- **NLTK:** Tokenization (`word_tokenize`, `sent_tokenize`) and stop word filtering
- **Pandas & OpenPyXL:** Excel and CSV data manipulation and output generation
- **ThreadPoolExecutor:** Concurrent asynchronous HTTP fetching and processing
