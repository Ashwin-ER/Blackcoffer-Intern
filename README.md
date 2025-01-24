# Blackcoffer-Intern

# Article Extraction and Text Analysis

This repository contains a Python script to extract articles from URLs and analyze their text. It uses libraries like `pandas`, `requests`, `BeautifulSoup`, and `nltk` to scrape, process, and analyze text data.

## Overview

The script does the following:
1. Takes a list of URLs from an Excel file.
2. Extracts the article title and text from each URL.
3. Saves the extracted articles as text files.
4. Analyzes the text for metrics like readability, word count, and more.
5. Saves the analysis results in an Excel file.

---

## Requirements

Make sure you have these Python libraries installed:
- `pandas`
- `requests`
- `BeautifulSoup`
- `nltk`
- `textstat`

Install them using:

```bash
pip install pandas requests beautifulsoup4 nltk textstat
