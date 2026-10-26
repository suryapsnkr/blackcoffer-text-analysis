"""
Blackcoffer Data Extraction and NLP Text Analysis Pipeline
---------------------------------------------------------
Author: Test Assignment Solution
Task:
  1. Data Extraction: Crawl article title and content from given URLs (Input.xlsx)
     and save each article into a text file named <URL_ID>.txt.
  2. AI Text Analysis: Compute 13 linguistic and sentimental variables as defined
     in 'Text Analysis.docx' and save the results in 'Output Data Structure.xlsx'.
"""

import os
import re
import sys
import string
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl

import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords as nltk_stopwords_corpus

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)

# Ensure required NLTK data models are available
for resource in ["punkt", "punkt_tab", "stopwords"]:
    try:
        nltk.download(resource, quiet=True)
    except Exception as e:
        logger.warning(f"Could not verify NLTK resource '{resource}': {e}")


def load_custom_stopwords(stopwords_dir="StopWords"):
    """
    Loads all stop words from the text files in the StopWords folder.
    Handles pipes '|' where metadata or category descriptions are delimited.
    """
    stopwords = set()
    if not os.path.exists(stopwords_dir):
        logger.error(f"StopWords directory not found at: {stopwords_dir}")
        return stopwords

    for filename in os.listdir(stopwords_dir):
        filepath = os.path.join(stopwords_dir, filename)
        if not os.path.isfile(filepath):
            continue
        try:
            with open(filepath, "r", encoding="latin-1") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    # Split on pipe '|' to isolate stop word from annotations
                    term = line.split("|")[0].strip().upper()
                    if term:
                        stopwords.add(term)
                        # Also add individual words if the term is a multi-word phrase
                        for part in term.split():
                            stopwords.add(part)
        except Exception as e:
            logger.error(f"Failed to read stopwords file {filename}: {e}")

    logger.info(f"Loaded {len(stopwords)} unique custom stop words.")
    return stopwords


def load_master_dictionary(dict_dir="MasterDictionary", custom_stopwords=None):
    """
    Loads positive and negative word lists from the MasterDictionary folder.
    Excludes any words found in the Stop Words lists as specified in the assignment.
    """
    if custom_stopwords is None:
        custom_stopwords = set()

    pos_words = set()
    neg_words = set()

    pos_path = os.path.join(dict_dir, "positive-words.txt")
    neg_path = os.path.join(dict_dir, "negative-words.txt")

    if os.path.exists(pos_path):
        with open(pos_path, "r", encoding="latin-1") as f:
            for line in f:
                w = line.strip()
                if w and not w.startswith(";"):
                    if w.upper() not in custom_stopwords:
                        pos_words.add(w.lower())
    else:
        logger.error(f"Positive words file not found at: {pos_path}")

    if os.path.exists(neg_path):
        with open(neg_path, "r", encoding="latin-1") as f:
            for line in f:
                w = line.strip()
                if w and not w.startswith(";"):
                    if w.upper() not in custom_stopwords:
                        neg_words.add(w.lower())
    else:
        logger.error(f"Negative words file not found at: {neg_path}")

    logger.info(
        f"Loaded {len(pos_words)} positive and {len(neg_words)} negative words "
        f"(after filtering against custom stop words)."
    )
    return pos_words, neg_words


def extract_article_content(url, timeout=20):
    """
    Extracts article title and clean body text from the given URL.
    Excludes website header, footer, navigation, and sidebar.
    """
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept-Language": "en-US,en;q=0.9",
    }

    try:
        response = requests.get(url, headers=headers, timeout=timeout)
        if response.status_code != 200:
            logger.warning(f"URL returned status {response.status_code}: {url}")
            return "", ""

        soup = BeautifulSoup(response.content, "html.parser")

        # 1. Title Extraction
        title = ""
        title_el = soup.find("h1")
        if title_el:
            title = title_el.get_text(strip=True)

        # 2. Article Text Extraction
        content_el = soup.find(class_="td-post-content")
        if not content_el:
            content_el = soup.find("div", class_="tdb-block-inner td-fix-index")

        if not content_el:
            logger.warning(f"No content container found for URL: {url}")
            return title, ""

        # Remove scripts, styles, and other non-content tags
        for unwanted in content_el.find_all(["script", "style", "noscript", "header", "footer"]):
            unwanted.decompose()

        # Extract structured paragraph blocks
        blocks = []
        for el in content_el.children:
            txt = el.get_text(separator=" ", strip=True) if hasattr(el, "get_text") else str(el).strip()
            if txt:
                blocks.append(txt)

        if not blocks:
            # Fallback if content container has text directly without block children
            raw_text = content_el.get_text(separator="\n")
            blocks = [line.strip() for line in raw_text.splitlines() if line.strip()]

        body_text = "\n\n".join(blocks)
        return title, body_text

    except Exception as e:
        logger.error(f"Error fetching URL {url}: {e}")
        return "", ""


def count_syllables(word):
    """
    Counts syllables in a word based on vowels (a, e, i, o, u).
    Handles exceptions like words ending in 'es' or 'ed'.
    """
    word = word.lower().strip()
    if not word:
        return 0
    # Strip any leading/trailing punctuation characters
    word = word.strip(string.punctuation)
    if not word:
        return 0

    # Rule: handle exceptions for words ending with "es" or "ed"
    if word.endswith("es") and not word.endswith(("tes", "ses", "ies")):
        word = word[:-2]
    elif word.endswith("ed") and not word.endswith(("ted", "ded", "ied")):
        word = word[:-2]

    vowels = "aeiou"
    count = sum(1 for char in word if char in vowels)
    return max(1, count)


def count_personal_pronouns(text):
    """
    Counts personal pronouns: 'I', 'we', 'my', 'ours', and 'us'.
    Excludes the country name 'US'.
    """
    pattern = r"\b(i|we|my|ours|us)\b"
    matches = [m.group() for m in re.finditer(pattern, text, re.IGNORECASE) if m.group() != "US"]
    return len(matches)


def compute_text_metrics(title, body, custom_stopwords, pos_words, neg_words, nltk_stopwords):
    """
    Computes all 13 textual and sentimental analysis metrics for an article.
    """
    full_text = f"{title}\n\n{body}".strip() if title else body.strip()

    if not full_text:
        return {
            "POSITIVE SCORE": 0,
            "NEGATIVE SCORE": 0,
            "POLARITY SCORE": 0.0,
            "SUBJECTIVITY SCORE": 0.0,
            "AVG SENTENCE LENGTH": 0.0,
            "PERCENTAGE OF COMPLEX WORDS": 0.0,
            "FOG INDEX": 0.0,
            "AVG NUMBER OF WORDS PER SENTENCE": 0.0,
            "COMPLEX WORD COUNT": 0,
            "WORD COUNT": 0,
            "SYLLABLE PER WORD": 0.0,
            "PERSONAL PRONOUNS": 0,
            "AVG WORD LENGTH": 0.0,
        }

    # Sentences
    sentences = sent_tokenize(full_text)
    num_sentences = len(sentences) if len(sentences) > 0 else 1

    # Tokenize words
    raw_tokens = word_tokenize(full_text)
    # Filter words: must contain at least one alphanumeric character and not be purely punctuation
    words = [
        w for w in raw_tokens
        if any(c.isalnum() for c in w) and not all(c in string.punctuation for c in w)
    ]
    total_words = len(words)

    # Clean words for Sentiment Analysis using custom StopWords lists
    cleaned_sentiment_words = [w.lower() for w in words if w.upper() not in custom_stopwords]
    total_cleaned_sentiment_words = len(cleaned_sentiment_words)

    # 1. Positive Score
    pos_score = sum(1 for w in cleaned_sentiment_words if w in pos_words)

    # 2. Negative Score
    neg_score = sum(1 for w in cleaned_sentiment_words if w in neg_words)

    # 3. Polarity Score: (Pos - Neg) / ((Pos + Neg) + 0.000001)
    polarity_score = (pos_score - neg_score) / ((pos_score + neg_score) + 0.000001)

    # 4. Subjectivity Score: (Pos + Neg) / (Total Words after cleaning + 0.000001)
    subjectivity_score = (pos_score + neg_score) / (total_cleaned_sentiment_words + 0.000001)

    # 5. Average Sentence Length = number of words / number of sentences
    avg_sentence_len = total_words / num_sentences if num_sentences > 0 else 0.0

    # 6. Syllable count per word & Complex Word Count (words with syllables > 2)
    syllables_list = [count_syllables(w) for w in words]
    total_syllables = sum(syllables_list)
    complex_word_count = sum(1 for s in syllables_list if s > 2)

    # 7. Percentage of Complex Words = (complex words / total words) * 100
    percentage_complex_words = (complex_word_count / total_words) * 100 if total_words > 0 else 0.0

    # 8. Fog Index = 0.4 * (Average Sentence Length + Percentage of Complex Words)
    fog_index = 0.4 * (avg_sentence_len + percentage_complex_words)

    # 9. Average Number of Words Per Sentence
    avg_words_per_sentence = total_words / num_sentences if num_sentences > 0 else 0.0

    # 10. Syllable Count Per Word = total syllables / total words
    syllables_per_word = total_syllables / total_words if total_words > 0 else 0.0

    # 11. Word Count (cleaned using NLTK stopwords and punctuation stripped)
    cleaned_nltk_words = []
    for w in words:
        w_clean = re.sub(r"^[^\w]+|[^\w]+$", "", w)
        if w_clean and any(c.isalpha() for c in w_clean) and w_clean.lower() not in nltk_stopwords:
            cleaned_nltk_words.append(w_clean)
    word_count_nltk = len(cleaned_nltk_words)

    # 12. Personal Pronouns
    personal_pronouns = count_personal_pronouns(full_text)

    # 13. Average Word Length = sum of word lengths / total words
    total_characters = sum(len(w) for w in words)
    avg_word_length = total_characters / total_words if total_words > 0 else 0.0

    return {
        "POSITIVE SCORE": pos_score,
        "NEGATIVE SCORE": neg_score,
        "POLARITY SCORE": polarity_score,
        "SUBJECTIVITY SCORE": subjectivity_score,
        "AVG SENTENCE LENGTH": avg_sentence_len,
        "PERCENTAGE OF COMPLEX WORDS": percentage_complex_words,
        "FOG INDEX": fog_index,
        "AVG NUMBER OF WORDS PER SENTENCE": avg_words_per_sentence,
        "COMPLEX WORD COUNT": complex_word_count,
        "WORD COUNT": word_count_nltk,
        "SYLLABLE PER WORD": syllables_per_word,
        "PERSONAL PRONOUNS": personal_pronouns,
        "AVG WORD LENGTH": avg_word_length,
    }


def process_single_row(row, custom_stopwords, pos_words, neg_words, nltk_stopwords, output_articles_dir):
    """
    Worker function to process a single URL:
    Extract text, save text file, and compute analysis metrics.
    """
    url_id = str(row["URL_ID"]).strip()
    url = str(row["URL"]).strip()

    title, body = extract_article_content(url)

    # Save extracted article text to text file with URL_ID as filename
    article_filename = f"{url_id}.txt"
    article_path = os.path.join(output_articles_dir, article_filename)
    try:
        with open(article_path, "w", encoding="utf-8") as f:
            if title:
                f.write(title + "\n\n")
            f.write(body + "\n")
    except Exception as e:
        logger.error(f"Failed to save article text for {url_id}: {e}")

    # Compute metrics
    metrics = compute_text_metrics(title, body, custom_stopwords, pos_words, neg_words, nltk_stopwords)

    result = {
        "URL_ID": url_id,
        "URL": url,
        **metrics
    }
    return result


def main():
    input_file = "Input.xlsx"
    output_excel_file = "Output Data Structure.xlsx"
    output_csv_file = "Output Data Structure.csv"
    articles_dir = "extracted_articles"

    if not os.path.exists(input_file):
        logger.error(f"Input file '{input_file}' not found.")
        sys.exit(1)

    os.makedirs(articles_dir, exist_ok=True)

    # Load stop words and dictionaries
    logger.info("Loading stop words and master dictionaries...")
    custom_stopwords = load_custom_stopwords("StopWords")
    pos_words, neg_words = load_master_dictionary("MasterDictionary", custom_stopwords)
    nltk_stopwords = set(nltk_stopwords_corpus.words("english"))

    # Read input file
    logger.info(f"Reading {input_file}...")
    df_input = pd.read_excel(input_file)
    total_articles = len(df_input)
    logger.info(f"Total articles to process: {total_articles}")

    results = []
    # Use ThreadPoolExecutor for concurrent web scraping and processing
    with ThreadPoolExecutor(max_workers=10) as executor:
        future_to_idx = {
            executor.submit(
                process_single_row,
                row,
                custom_stopwords,
                pos_words,
                neg_words,
                nltk_stopwords,
                articles_dir
            ): idx
            for idx, row in df_input.iterrows()
        }

        for future in as_completed(future_to_idx):
            idx = future_to_idx[future]
            try:
                res = future.result()
                results.append((idx, res))
                logger.info(f"Processed [{len(results)}/{total_articles}]: {res['URL_ID']}")
            except Exception as e:
                logger.error(f"Row {idx} failed with error: {e}")

    # Sort results back into the original order
    results.sort(key=lambda x: x[0])
    ordered_data = [item[1] for item in results]

    # Create output DataFrame
    df_output = pd.DataFrame(ordered_data)

    # Ensure column order matches Output Data Structure.xlsx exactly
    expected_columns = [
        "URL_ID",
        "URL",
        "POSITIVE SCORE",
        "NEGATIVE SCORE",
        "POLARITY SCORE",
        "SUBJECTIVITY SCORE",
        "AVG SENTENCE LENGTH",
        "PERCENTAGE OF COMPLEX WORDS",
        "FOG INDEX",
        "AVG NUMBER OF WORDS PER SENTENCE",
        "COMPLEX WORD COUNT",
        "WORD COUNT",
        "SYLLABLE PER WORD",
        "PERSONAL PRONOUNS",
        "AVG WORD LENGTH",
    ]
    df_output = df_output[expected_columns]

    # Save to Excel and CSV
    df_output.to_excel(output_excel_file, index=False)
    df_output.to_csv(output_csv_file, index=False)
    logger.info(f"Results successfully saved to '{output_excel_file}' and '{output_csv_file}'")
    logger.info(f"All {len(ordered_data)} text files saved in directory: '{articles_dir}/'")


if __name__ == "__main__":
    main()
