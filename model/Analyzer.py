from nltk.stem.snowball import SnowballStemmer
from nltk.stem import WordNetLemmatizer
import pickle
from TextPrepare import TextPrepare
from FileReader import FileReader
import matplotlib.pyplot as plt

class Analyzer():
    def __init__(self, dict_name):
        self.analyzed_data = {}
        with open(dict_name, 'rb') as f:
            self.emot_dict = pickle.load(f)

    def analyze(self, prep_text_dict, key_words, range_l=2, range_r=2):
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
                        while i2 > 0 and j < range_l:
                            i2 -= 1
                            j += 1
                            if sent[i2] in self.emot_dict:
                                if emot_value == 0 or emot_value == self.emot_dict[sent[i2]]:
                                    emot_value = self.emot_dict[sent[i2]]
                                else:
                                    many_values_flag = True
                        j = 0
                        i2 = i
                        while i2 < len(sent)-1 and j < range_r:
                            i2 += 1
                            j += 1
                            if sent[i2] in self.emot_dict:
                                if emot_value == 0 or emot_value == self.emot_dict[sent[i2]]:
                                    emot_value = self.emot_dict[sent[i2]]
                                else:
                                    many_values_flag = True
                        count[0] += 1
                        if not many_values_flag and emot_value != 0:
                            count[int(emot_value)] += 1
            count[3] = round(count[0]/all_word_count, 3) * 100
            prep_text_dict[keys] = count
        self.analyzed_data = prep_text_dict
        return prep_text_dict

    def analyze_en(self, prep_text_dict, key_words, range_l=2, range_r=2):
        lemmatizer = WordNetLemmatizer()
        lemma_key_words = []
        for word in key_words:
            lemma_key_words.append(lemmatizer.lemmatize(word.lower()))
        for keys, values in prep_text_dict.items():
            all_word_count = 0
            count = [0, 0, 0, 0, 0, 0]  # общий счётчик, счётчик 1, процент всего слов/упоминаний, счётчик 2, счётчик -2, счётчик -1
            for sent in values:
                all_word_count += len(sent)
                for i in range(0, len(sent)):
                    if sent[i] in lemma_key_words:
                        j = 0
                        i2 = i
                        many_values_flag = False
                        emot_value = 0
                        while i2 > 0 and j < range_l:
                            i2 -= 1
                            j += 1
                            if sent[i2] in self.emot_dict:
                                if emot_value == 0 or emot_value == self.emot_dict[sent[i2]]:
                                    emot_value = self.emot_dict[sent[i2]]
                                else:
                                    many_values_flag = True
                        j = 0
                        i2 = i
                        while i2 < len(sent)-1 and j < range_r:
                            i2 += 1
                            j += 1
                            if sent[i2] in self.emot_dict:
                                if emot_value == 0 or emot_value == self.emot_dict[sent[i2]]:
                                    emot_value = self.emot_dict[sent[i2]]
                                else:
                                    many_values_flag = True
                        count[0] += 1
                        if not many_values_flag and emot_value != 0:
                            count[int(emot_value)] += 1
            count[3] = round(count[0]/all_word_count, 3) * 100
            prep_text_dict[keys] = count
        self.analyzed_data = prep_text_dict
        return prep_text_dict


if __name__ == "__main__":

    reader = FileReader()
    reader.read_files(['C:/Dev/DIPLOM/docs/NSS-2010.pdf', 'C:/Dev/DIPLOM/docs/NSS-2015.pdf', 'C:/Dev/DIPLOM/docs/NSS-2017.pdf'])

    preparer = TextPrepare()
    preparer.prepare_en(reader.files_text)
    # print(reader.files_text)

    analyzer = Analyzer('dicts/dictionary.pkl')
    analyzer.analyze_en(preparer.prep_text_dict,['China', 'PRC'], 4)
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
