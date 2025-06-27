import nltk
from nltk.corpus import stopwords
from nltk.stem.snowball import SnowballStemmer
from nltk.stem import WordNetLemmatizer


class TextPrepare():
    def __init__(self):
        self.prep_text_dict = {}

    def prepare(self, text_dict):
        stop_words = set(stopwords.words('russian'))
        stop_words.add(',')
        stop_words.add('.')
        stemmer = SnowballStemmer("russian")
        texts = text_dict
        for keys, values in texts.items():
            text_tokenize = []
            sentences = nltk.sent_tokenize(values)
            for sent in sentences:
                words = nltk.word_tokenize(sent)
                clean_words = []
                for word in words:
                    if word not in stop_words:  # Проверка на стоп слово
                        clean_words.append((stemmer.stem(word.lower())))  # Стемминг
                text_tokenize.append(clean_words)
            texts[keys] = text_tokenize
        self.prep_text_dict = texts
        return texts

    def prepare_en(self, text_dict):
        stop_words = set(stopwords.words('english'))
        stop_words.add(',')
        stop_words.add('.')
        lemmatizer = WordNetLemmatizer()
        texts = text_dict
        for keys, values in texts.items():
            text_tokenize = []
            sentences = nltk.sent_tokenize(values)
            for sent in sentences:
                words = nltk.word_tokenize(sent)
                clean_words = []
                for word in words:
                    if word not in stop_words:  # Проверка на стоп слово
                        clean_words.append((lemmatizer.lemmatize(word.lower())))  # Стемминг
                text_tokenize.append(clean_words)
            texts[keys] = text_tokenize
        self.prep_text_dict = texts
        return texts