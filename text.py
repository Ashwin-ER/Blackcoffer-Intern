# Import required libraries
import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import nltk
import textstat
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords

# Ensure necessary NLTK data is downloaded
nltk.download('punkt')
nltk.download('stopwords')

# Load the input Excel file
input_file = "Input.xlsx"  # Update with the actual path to your input file
data = pd.read_excel(input_file)

# Create a folder to save the articles
output_folder = "Articles"
os.makedirs(output_folder, exist_ok=True)

# Function to extract article title and text
def extract_article(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, "html.parser")

        # Extract the title (modify selectors as needed based on website structure)
        title = soup.find("h1").get_text(strip=True)

        # Extract the main content (modify selectors as needed based on website structure)
        content = soup.find("article")  # Replace "article" with the appropriate tag or class
        if content:
            text = content.get_text(separator="\n", strip=True)
        else:
            text = "Content not found."

        return title, text
    except Exception as e:
        return "Error", f"Failed to extract: {e}"

# Process each URL in the Excel file
for index, row in data.iterrows():
    url_id = row['URL_ID']
    url = row['URL']

    print(f"Processing URL_ID: {url_id}")

    # Extract article content
    title, text = extract_article(url)

    # Save the extracted content into a text file
    file_name = f"{output_folder}/{url_id}.txt"
    with open(file_name, "w", encoding="utf-8") as file:
        file.write(f"Title: {title}\n\n")
        file.write(f"Text: {text}")

print("Article extraction completed. Files are saved in the 'Articles' folder.")

# Function to analyze text
def analyze_text(text):
    # Tokenization
    words = word_tokenize(text)
    sentences = sent_tokenize(text)

    # Remove stop words
    stop_words = set(stopwords.words('english'))
    filtered_words = [w for w in words if not w.lower() in stop_words]

    # Calculate variables
    positive_score = textstat.coleman_liau_index(text)  # Proxy for positive score
    negative_score = textstat.gunning_fog(text)  # Proxy for negative score
    polarity_score = positive_score - negative_score
    subjectivity_score = textstat.flesch_reading_ease(text)
    avg_sentence_length = len(words) / len(sentences)
    percentage_of_complex_words = textstat.difficult_words(text) / len(words)
    fog_index = textstat.gunning_fog(text)
    avg_number_of_words_per_sentence = avg_sentence_length
    complex_word_count = textstat.difficult_words(text)
    word_count = len(words)
    syllable_per_word = textstat.syllable_count(text) / len(words)
    personal_pronouns = len([word for word in words if word.lower() in ['i', 'me', 'my', 'mine', 'we', 'us', 'our', 'ours']])
    avg_word_length = sum(len(word) for word in words) / len(words)

    return {
        'POSITIVE SCORE': positive_score,
        'NEGATIVE SCORE': negative_score,
        'POLARITY SCORE': polarity_score,
        'SUBJECTIVITY SCORE': subjectivity_score,
        'AVG SENTENCE LENGTH': avg_sentence_length,
        'PERCENTAGE OF COMPLEX WORDS': percentage_of_complex_words,
        'FOG INDEX': fog_index,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_number_of_words_per_sentence,
        'COMPLEX WORD COUNT': complex_word_count,
        'WORD COUNT': word_count,
        'SYLLABLE PER WORD': syllable_per_word,
        'PERSONAL PRONOUNS': personal_pronouns,
        'AVG WORD LENGTH': avg_word_length,
    }

# Load the output structure file
output_structure = pd.read_excel("Output Data Structure.xlsx")

# Create a list to store the results
results = []

# Process each extracted text file
for filename in os.listdir(output_folder):
    if filename.endswith(".txt"):
        with open(os.path.join(output_folder, filename), 'r', encoding="utf-8") as f:
            text = f.read().split('Text: ')[1]  # Extract text content after "Text:"
            url_id = filename[:-4]  # Extract URL_ID from filename

            # Perform analysis
            analysis_result = analyze_text(text)

            # Add URL_ID to the result dictionary
            analysis_result['URL_ID'] = url_id

            # Append the result to the list
            results.append(analysis_result)

# Create a DataFrame from the results
df = pd.DataFrame(results)

# Reorder columns to match output structure using common columns
common_cols = df.columns.intersection(output_structure.columns)
df = df[common_cols]

# Save the output to an Excel file
df.to_excel("Output.xlsx", index=False)

print("Text analysis completed. Results are saved in 'Output.xlsx'.")
