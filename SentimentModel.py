"""
SENTIMENT ANALYSIS ON GAME REVIEW

Steps:
1. install TextBlob pythonic text processing and sentiment analysis library
        Run following command!
        $ pip install -U textblob
        $ python -m textblob.download_corpora

2. install pandas : python package for data loading and analysis
        Run below command!
        $ pip install pandas 

3. install opnpyxl: python library to read/write excel
        Run below command!
        $ pip install openpyxl

4. install matplotlib
        Run below command!
        $ pip install matplotlib
"""

from textblob import TextBlob
import pandas as pd
import matplotlib.pyplot as plt
import re

SYMBOLS = re.compile("[.;:!\'?,\"()\[\]]")
TAGS = re.compile("(<br\s*/><br\s*/>)|(\-)|(\/)")
POSITIVE = "Positive"
NEGATIVE = "Negative"
NEUTRAL = "Neutral"

# Load file and extract review column
def load_data(file):
    dataframe = pd.read_excel(file)
    data = dataframe['review_text']
    dataframe.head()
    return data

# Process the review, removes the symbols and tags
def preprocess_reviews(reviews):
    reviews = [SYMBOLS.sub("", line.lower()) for line in reviews]
    reviews = [TAGS.sub(" ", line) for line in reviews]
    return reviews

# returns the polarity score of a review
def get_polarity_score(sentence):
    return TextBlob(sentence).sentiment.polarity

# get the sentiment of reviews
def get_sentiment(reviews):
    sentiments = []
    for review in reviews:
        # decide sentiment as positive, negative and neutral 
        # positive sentiment : (polarity_score >= 0.05)
        # neutral sentiment : (polarity_score > -0.05) and (polarity_score < 0.05)
        # negative sentiment : (polarity_score <= -0.05)
        polarity_score = get_polarity_score(review)

        if polarity_score >= 0.05 : 
            sentiments.append(POSITIVE) 
    
        elif polarity_score <= - 0.05 : 
            sentiments.append(NEGATIVE) 
    
        else : 
            sentiments.append(NEUTRAL)
    return sentiments

# write sentiment of review in a new excel file
def write_sentiment(sentiments, original_file):
    df = pd.DataFrame(pd.read_excel(original_file))
    df['Sentiment'] = sentiments
    df.to_excel("output.xlsx")

def count(sentiments):
    count_pos = sentiments.count(POSITIVE)
    count_neg = sentiments.count(NEGATIVE)
    count_neu = sentiments.count(NEUTRAL)
    return [count_pos, count_neg, count_neu]

# Plot pie graph of sentiment and saves as staistics.png
def plot_pie(sentiments):
    arr = count(sentiments)
    df = pd.DataFrame({'sentiment': arr},
                  index=[POSITIVE, NEGATIVE, NEUTRAL])
    df.plot.pie(y='sentiment', figsize=(5, 5), autopct='%1.1f%%', startangle=90)
    plt.savefig("statistics.png")
    plt.show()

def main():
    reviews = load_data('test.xlsx')
    processed_reviews = preprocess_reviews(reviews)
    sentiments = get_sentiment(processed_reviews)
    write_sentiment(sentiments, 'test.xlsx')
    plot_pie(sentiments)

main()