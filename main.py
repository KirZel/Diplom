import os
from spire.doc import *
from spire.doc.common import *
from pypdf import PdfReader
import nltk
from nltk.corpus import stopwords
from nltk.stem.snowball import SnowballStemmer
import pickle
import matplotlib.pyplot as plt


class FileReader():

    def __init__(self):
        self.files_text = {}

    def read_files(self, filepaths):
        for file in filepaths:
            file_type = file.split('.')[-1]
            file_name = file.split('/')[-1]
            if file_type == 'txt':
                file = open(file, "r", encoding="utf-8")
                text = file.read()
                self.files_text[file_name] = text
                file.close()
            if file_type == 'docx':
                doc = Document()
                doc.LoadFromFile(file)
                text = doc.GetText()
                self.files_text[file_name] = text
            if file_type == 'pdf':
                reader = PdfReader(file)
                text = ''
                for page in reader.pages:
                    text += page.extract_text()
                self.files_text[file_name] = text
        return self.files_text


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


class Analyzer():
    def __init__(self, dict_name):
        self.analyzed_data = {}
        with open(dict_name, 'rb') as f:
            self.emot_dict = pickle.load(f)

    def analyze(self, prep_text_dict, key_words, search_range=2):
        stemmer = SnowballStemmer("russian")
        stem_key_words = []
        for word in key_words:
            stem_key_words.append(stemmer.stem(word.lower()))
        for keys, values in prep_text_dict.items():
            all_word_count = 0
            count = [0, 0, 0, 0, 0, 0]  # общий счётчик, счётчик 1, процент всего слов/упоминаний, счётчик 2, счётчик -2, счётчик -1
            for sent in values:
                all_word_count += len(sent)
                for i in range(0, len(sent)):
                    if sent[i] in stem_key_words:
                        j = 0
                        i2 = i
                        many_values_flag = False
                        emot_value = 0
                        while i > 0 and j < search_range:
                            i2 -= 1
                            j += 1
                            if sent[i2] in self.emot_dict:
                                if emot_value == 0 or emot_value == self.emot_dict[sent[i2]]:
                                    emot_value = self.emot_dict[sent[i2]]
                                else:
                                    many_values_flag = True
                        j = 0
                        i2 = i
                        while i < len(sent) and j < search_range:
                            i2 += 1
                            j += 1
                            if sent[i2] in self.emot_dict:
                                if emot_value == 0 or emot_value == self.emot_dict[sent[i2]]:
                                    emot_value = self.emot_dict[sent[i2]]
                                else:
                                    many_values_flag = True
                        count[0] += 1
                        if not many_values_flag and emot_value != 0:
                            count[emot_value] += 1
            count[3] = round(count[0]/all_word_count, 3) * 100
            prep_text_dict[keys] = count
        self.analyzed_data = prep_text_dict
        return prep_text_dict

class DictWork():
    def __init__(self,dict_name):
        self.dict_name = dict_name
        self.analyzed_data = {}
        with open(dict_name, 'rb') as f:
            self.emot_dict = pickle.load(f)

    def add_word(self, word, value):
        self.emot_dict[word] = value

    def delete_word(self, word):
        del self.emot_dict[word]

    def save_dict(self):
        with open(self.dict_name, 'wb') as f:
            pickle.dump(self.emot_dict, f)

    def save_dict_as(self, new_name):
        with open(new_name, 'wb') as f:
            pickle.dump(self.emot_dict, f)


if __name__ == "__main__":

    reader = FileReader()
    reader.read_files(['C:/Dev/DIPLOM/docs/1.txt', 'C:/Dev/DIPLOM/docs/2.docx', 'C:/Dev/DIPLOM/docs/3.pdf'])

    preparer = TextPrepare()
    preparer.prepare(reader.files_text)
    # print(reader.files_text)

    analyzer = Analyzer('dictionary.pkl')
    analyzer.analyze(preparer.prep_text_dict,['текст', 'метод'], 3)
    print(analyzer.analyzed_data)
    print(analyzer.emot_dict)

    docs = analyzer.analyzed_data.keys()
    count = []
    count_procents = []
    count_pos1 = []
    count_pos2 = []
    count_neg1 = []
    count_neg2 = []
    for values in analyzer.analyzed_data.values():
        count.append(values[0])
        count_procents.append(values[3])
        count_pos1.append(values[1])
        count_pos2.append(values[2])
        count_neg1.append(values[-1])
        count_neg2.append(values[-2])

    plt.subplot(1, 2, 1)
    plt.plot(docs, count, ':', label='Вхождений всего', color='blue', marker='o', markersize=3)
    plt.plot(docs, count_pos1, '-', label='Вхождений с оценкой 1', color='green', marker='o', markersize=5)
    plt.plot(docs, count_pos2, '--', label='Вхождений с оценкой 2', color='lime', marker='o', markersize=6)
    plt.plot(docs, count_neg1, '-.', label='Вхождений с оценкой -1', color='darkred', marker='o', markersize=4)
    plt.plot(docs, count_neg2, ':', label='Вхождений с оценкой -2', color='red', marker='o', markersize=3)
    plt.xlabel('Документы')
    plt.ylabel('Кол-во вхождений')
    plt.title('График вхождений')
    plt.legend(fontsize=9)


    plt.subplot(1, 2, 2)
    plt.plot(docs, count_procents, ':', label='Процентное соотношение', color='blue', marker='o', markersize=3)
    plt.xlabel('Документы')
    plt.ylabel('% вхождений')
    plt.title('График % вхождений')
    plt.show()
