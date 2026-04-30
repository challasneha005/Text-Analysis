#importing necessary libraries
import pandas as pd
import requests
from bs4 import BeautifulSoup
import  os
import nltk
import string
import re
from nltk.tokenize import word_tokenize,sent_tokenize 
nltk.download('punkt')

#taking input and declaring output
input_df = pd.read_excel("Input.xlsx")
output_folder = "Extracted_Articles"

os.makedirs(output_folder,exist_ok=True)

#declaring function of extraction of the url
def extract_article(url):
    try:
        response = requests.get(url,timeout = 10)
        if response.status_code != 200:
            return None,None
        
        soup = BeautifulSoup(response.text,'html.parser')
        title = soup.find('h1')
        if title:
            title = title.get_text(strip = True)

        article_body = soup.find('article')
        if not article_body:
            article_body = soup.find('div',class_='article') or soup.find('div',class_='td-post-content')

        if article_body:
            paragraph = article_body.find_all(['p','h2','h3'])
            content = "\n".join(p.get_text(strip = True) for p in paragraph if p.get_text(strip = True))

        else:
            content = ""
        return title,content
    
    except Exception as e:
        print(f"Error occurred while extracting  {url}: {str(e)}")
        return None,None

#iteration through the excel sheet
for index,row in input_df.iterrows():
    url_id=str(row['URL_ID'])
    url=row['URL']
    print(f"processing{url_id}-{url}")
    title,content=extract_article(url)

    if title and content:
        with open(f"{output_folder}/{url_id}.txt","w",encoding = "utf-8") as file:
            file.write(title+"\n\n"+content)
    else:
        print(f"Failed to extract article for {url_id}-{url}")

print("Extraction completed")
 
#path assignment i.e folders and files
stopwords_folder = "StopWords"
master_dict_folder = "MasterDictionary"
output_structure_file = "Output Data Structure.xlsx"
input_file = "Input.xlsx"
articles_folder = "Extracted_Articles"
final_output_file = "final_output.xlsx"

stop_words = set()

#loading all stopwords files from stopwords folder
for file_name in os.listdir(stopwords_folder):
    file_path = os.path.join(stopwords_folder,file_name)
    if file_path.endswith('.txt'):
        with open(file_path,'r',encoding='ISO-8859-1') as f:
            for line in f:
                word = line.strip().lower()
                if word:
                    stop_words.add(word)

#loading Master Dictionary Words (positive & negative)
def load_word_list(filename):
    path = os.path.join(master_dict_folder,filename)
    with open(path, 'r', encoding='ISO-8859-1') as f:
        return set(word.strip().lower() for word in f if word.strip() and word.strip().isalpha())

positive_words = load_word_list("positive-words.txt") - stop_words
negative_words = load_word_list("negative-words.txt") - stop_words

#evaluating necessary columns
vowels = "aeiou"

def clean_text(text):
    tokens = word_tokenize(text.lower())
    words = [word.strip(string.punctuation) for word in tokens if word.strip(string.punctuation)]
    return [word for word in words if word and word not in stop_words]

def count_syllables(word):
    word = word.lower()
    if word.endswith(("es", "ed")):
        word = word[:-2]
    return sum(1 for char in word if char in vowels)

def is_complex(word):
    return count_syllables(word) > 2

def count_personal_pronouns(text):
    return len(re.findall(r'\b(I|we|my|ours|us)\b',text,re.I)) - len(re.findall(r'\bUS\b',text))

def analyze_article(text):
    words = clean_text(text)
    num_words = len(words)
    num_sentences = max(1, len(sent_tokenize(text)))

    pos_score = sum(1 for word in words if word in positive_words)
    neg_score = sum(1 for word in words if word in negative_words)
    polarity_score = (pos_score - neg_score) / ((pos_score + neg_score) + 1e-6)
    subjectivity_score = (pos_score + neg_score) / (num_words + 1e-6)

    avg_sentence_length = num_words / num_sentences
    complex_words = [word for word in words if is_complex(word)]
    percent_complex_words = len(complex_words) / num_words
    fog_index = 0.4 * (avg_sentence_length + percent_complex_words)

    avg_words_per_sentence = num_words / num_sentences
    complex_word_count = len(complex_words)
    syllable_per_word = sum(count_syllables(word) for word in words) / num_words
    personal_pronouns = count_personal_pronouns(text)
    avg_word_length = sum(len(word) for word in words) / num_words

    return {
        "POSITIVE SCORE": pos_score,
        "NEGATIVE SCORE": neg_score,
        "POLARITY SCORE": polarity_score,
        "SUBJECTIVITY SCORE": subjectivity_score,
        "AVG SENTENCE LENGTH": avg_sentence_length,
        "PERCENTAGE OF COMPLEX WORDS": percent_complex_words,
        "FOG INDEX": fog_index,
        "AVG NUMBER OF WORDS PER SENTENCE": avg_words_per_sentence,
        "COMPLEX WORD COUNT": complex_word_count,
        "WORD COUNT": num_words,
        "SYLLABLE PER WORD": syllable_per_word,
        "PERSONAL PRONOUNS": personal_pronouns,
        "AVG WORD LENGTH": avg_word_length
    }

#loading input and ouput files
input_df = pd.read_excel(input_file)
output_df = pd.read_excel(output_structure_file)

results = []

#analysis for each text file
for _, row in input_df.iterrows():
    url_id = str(row['URL_ID'])
    file_path = os.path.join(articles_folder,f"{url_id}.txt")

    if os.path.exists(file_path):
        with open(file_path, 'r', encoding = 'utf-8') as f:
            text = f.read()

        metrics = analyze_article(text)
        metrics['URL_ID'] = url_id
        metrics['URL'] = row['URL']
        results.append(metrics)
    else:
        print(f" Article not found for URL_ID: {url_id}")
 
output_df = pd.read_excel(output_structure_file)

#update evaluation columns
result_df = pd.DataFrame(results)
 
output_df['URL_ID'] = output_df['URL_ID'].astype(str)
result_df['URL_ID'] = result_df['URL_ID'].astype(str)

output_df.set_index('URL_ID', inplace = True)
result_df.set_index('URL_ID', inplace = True)

#list of evaluation columns
eval_cols = [
    "POSITIVE SCORE", "NEGATIVE SCORE", "POLARITY SCORE", "SUBJECTIVITY SCORE",
    "AVG SENTENCE LENGTH", "PERCENTAGE OF COMPLEX WORDS", "FOG INDEX",
    "AVG NUMBER OF WORDS PER SENTENCE", "COMPLEX WORD COUNT", "WORD COUNT",
    "SYLLABLE PER WORD", "PERSONAL PRONOUNS", "AVG WORD LENGTH"
]
#filling each column from result_df to output_df
for col in eval_cols:
    if col in result_df.columns:
        output_df[col] = result_df[col]

 #save final output to excel
output_df.reset_index(inplace=True)
output_df.to_excel(final_output_file,index = False)
print(f"\nFinally output saved to '{final_output_file}'")
 
