# Instructions & Methodology Documentation

## Blackcoffer Data Extraction and NLP Text Analysis

---

### 1. Project Overview
This project implements an end-to-end automated Python solution for:
1. **Data Extraction:** Crawling article titles and textual content from a list of URLs provided in `Input.xlsx`, strictly filtering out headers, footers, navigation, sidebars, and comments, and saving each article as an individual text file named `<URL_ID>.txt`.
2. **AI Text Analysis & NLP:** Computing 13 linguistic, sentimental, and readability metrics as defined in `Text Analysis.docx`, and compiling the results into `Output Data Structure.xlsx` (and `Output Data Structure.csv`).

---

### 2. Solution Architecture & Approach

```
┌──────────────────┐       ┌────────────────────────┐
│    Input.xlsx    │ ───>  │  main.py               │
└──────────────────┘       │  (ThreadPoolExecutor)  │
                           └───────────┬────────────┘
                                       │
                ┌──────────────────────┴──────────────────────┐
                ▼                                             ▼
     [ Data Extraction ]                            [ Text Analysis ]
  • Fetch URL with HTTP retries                  • Filter MasterDictionary with StopWords
  • Parse via BeautifulSoup                      • Calculate Sentiment Scores (Pos/Neg)
  • Extract <h1> & .td-post-content              • Calculate Readability & Gunning Fog Index
  • Strip scripts, styles, & noise               • Calculate Syllables, Pronouns, & Word Length
                │                                             │
                ▼                                             ▼
┌──────────────────────────────┐              ┌────────────────────────────────┐
│  extracted_articles/<ID>.txt │              │  Output Data Structure.xlsx    │
└──────────────────────────────┘              └────────────────────────────────┘
```

#### Step 1: Data Extraction
* **Target Identification:** For each URL in `Input.xlsx`, an HTTP GET request is dispatched with modern browser user-agent headers and timeout handling.
* **HTML Parsing:** Using `BeautifulSoup`, the article title is retrieved from the `<h1>` tag (`h1.entry-title`).
* **Content Container:** The primary article content is located inside `.td-post-content` (with a fallback to `.tdb-block-inner.td-fix-index` for TagDiv themes).
* **Noise Elimination:** All `<script>`, `<style>`, `<noscript>`, `<header>`, and `<footer>` elements are decomposed.
* **Paragraph Structure:** Structured content blocks (paragraphs, lists, headings) are preserved, joined cleanly with double newlines (`\n\n`), and saved to `extracted_articles/<URL_ID>.txt` in UTF-8 encoding.

#### Step 2: Stop Words & Sentiment Dictionaries Setup
* **Custom Stop Words:** Loaded from all text files in the `StopWords/` directory (`Auditor`, `Currencies`, `DatesandNumbers`, `Generic`, `GenericLong`, `Geographic`, `Names`). Pipe delimiters (`|`) separating words from explanations or census annotations are parsed, and words are normalized to uppercase.
* **Master Dictionary:** Loaded from `MasterDictionary/positive-words.txt` and `MasterDictionary/negative-words.txt`. Words found in the custom Stop Words list are excluded to create cleaned Positive and Negative word sets.

#### Step 3: Text Analysis & Variable Definitions
For each article, the following 13 variables are computed:

1. **POSITIVE SCORE:** Number of words found in the filtered Positive Dictionary.
2. **NEGATIVE SCORE:** Number of words found in the filtered Negative Dictionary.
3. **POLARITY SCORE:** Determines whether the article tone is positive or negative:
   $$\text{Polarity Score} = \frac{\text{Positive Score} - \text{Negative Score}}{(\text{Positive Score} + \text{Negative Score}) + 0.000001}$$
   *Range:* -1.0 to +1.0.
4. **SUBJECTIVITY SCORE:** Determines whether the text is objective or subjective:
   $$\text{Subjectivity Score} = \frac{\text{Positive Score} + \text{Negative Score}}{(\text{Total Words after cleaning}) + 0.000001}$$
   *Range:* 0.0 to +1.0.
5. **AVG SENTENCE LENGTH:**
   $$\text{Average Sentence Length} = \frac{\text{Total Number of Words}}{\text{Total Number of Sentences}}$$
6. **PERCENTAGE OF COMPLEX WORDS:**
   $$\text{Percentage of Complex Words} = \left(\frac{\text{Number of Complex Words}}{\text{Total Number of Words}}\right) \times 100$$
   *(Complex words are defined as words containing more than 2 syllables).*
7. **FOG INDEX:** Gunning Fog readability index:
   $$\text{Fog Index} = 0.4 \times (\text{Average Sentence Length} + \text{Percentage of Complex Words})$$
8. **AVG NUMBER OF WORDS PER SENTENCE:**
   $$\text{Average Number of Words Per Sentence} = \frac{\text{Total Number of Words}}{\text{Total Number of Sentences}}$$
9. **COMPLEX WORD COUNT:** Total count of words with more than 2 syllables.
10. **WORD COUNT:** Total count of cleaned words after removing NLTK English stop words and punctuation (following Section 5 of `Text Analysis.docx`).
11. **SYLLABLE PER WORD:**
    $$\text{Syllable Per Word} = \frac{\text{Total Syllable Count}}{\text{Total Number of Words}}$$
    *(Syllables are determined by counting vowels `a, e, i, o, u`, with exceptions for words ending with `es` or `ed`).*
12. **PERSONAL PRONOUNS:** Count of words `I`, `we`, `my`, `ours`, and `us` using regex word boundaries, while strictly excluding the country acronym `US`.
13. **AVG WORD LENGTH:**
    $$\text{Average Word Length} = \frac{\text{Total Number of Characters in All Words}}{\text{Total Number of Words}}$$

---

### 3. Dependencies & Installation

#### Required Packages
* Python 3.8+
* `requests`
* `beautifulsoup4`
* `pandas`
* `openpyxl`
* `nltk`

To install all dependencies, run:
```bash
pip install -r requirements.txt
```

---

### 4. How to Run the Program

1. Ensure the workspace contains:
   * `main.py`
   * `Input.xlsx`
   * `MasterDictionary/` folder (`positive-words.txt`, `negative-words.txt`)
   * `StopWords/` folder (`StopWords_*.txt`)
   * `Output Data Structure.xlsx` (template)
2. Execute the Python script from terminal / command line:
   ```bash
   python main.py
   ```
3. The script will automatically:
   * Download necessary NLTK tokenizers (`punkt`, `punkt_tab`, `stopwords`).
   * Crawl all URLs in `Input.xlsx` concurrently using worker threads.
   * Save extracted articles to `extracted_articles/<URL_ID>.txt`.
   * Calculate all 13 metrics for every article in the exact order specified.
   * Output the completed results to `Output Data Structure.xlsx` and `Output Data Structure.csv`.

---

### 5. Deliverables Checklist
* [x] **Python script:** `main.py`
* [x] **Extracted text files:** `extracted_articles/<URL_ID>.txt` (all 82 articles)
* [x] **Output files:** `Output Data Structure.xlsx` and `Output Data Structure.csv`
* [x] **Dependencies file:** `requirements.txt`
* [x] **Documentation:** `Instructions.md` & `Instructions.docx`
