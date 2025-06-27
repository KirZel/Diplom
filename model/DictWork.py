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