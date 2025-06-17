from main import FileReader, TextPrepare, Analyzer, DictWork
import customtkinter as ctk
from tkinter import filedialog
from tkinter import messagebox as mb
import matplotlib.pyplot as plt
import pickle

class App(ctk.CTk):
    filepaths = []
    dictpath = "C:\Dev\DIPLOM\dicts\dictionary.pkl" # ВСТАВИТЬ ДЕФОЛТ ВАРИАНТ
    keywords = []
    search_range = "1"
    emot_dict = {}

    def __init__(self):
        super().__init__()

        self.title("Main")
        self.geometry("800x650")
        self.grid_columnconfigure((0, 1), weight=1)
        self.resizable(False, False)

        self.create_dict_chosing()

        self.create_file_table()

        self.create_keywords_list()

        self.create_settings()

        self.file_button = ctk.CTkButton(self, text="Редактор словарей", command=self.dict_redactor)
        self.file_button.grid(row=4, column=0, padx=20, pady=20, sticky="w", columnspan=2)

    def dict_redactor(self):
        self.new_window = ctk.CTkToplevel(self)
        self.new_window.title("Редактор")
        self.new_window.geometry("800x600")
        self.new_window.resizable(False, False)

        label = ctk.CTkLabel(self.new_window, text="Выберите словарь", font=("Arial", 12, "bold"))
        label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.dict_button = ctk.CTkButton(self.new_window, text="Выбрать словарь", command=self.load_dict)
        self.dict_button.grid(row=1, column=0, padx=5, pady=5, sticky="w", columnspan=2)

        self.dict_table = ctk.CTkScrollableFrame(self.new_window, width=160, height=200)
        self.dict_table.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="w")

        self.header_dict_table = ctk.CTkFrame(self.dict_table, fg_color="transparent")
        self.header_dict_table.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        label_dict_table_name = ctk.CTkLabel(self.header_dict_table, text="Слово", font=("Arial", 12, "bold"))
        label_dict_table_name.pack(side="left", padx=(0, 4))

        self.header_dict_table = ctk.CTkFrame(self.dict_table, fg_color="transparent")
        self.header_dict_table.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        label_dict_table_name = ctk.CTkLabel(self.header_dict_table, text="Оценка", font=("Arial", 12, "bold"))
        label_dict_table_name.pack(side="left", padx=(0, 4))

        self.entry_dict_word = ctk.CTkEntry(self.new_window, placeholder_text="Введите слово")
        self.entry_dict_word.grid(row=3, column=0, padx=5, pady=5, sticky="nsew")

        ranges = ["2", "1", "-1", "-2"]
        self.dict_range = ctk.CTkOptionMenu(self.new_window, values=ranges)
        self.dict_range.grid(row=3, column=1, padx=5, pady=5, sticky="nsew")

        self.dict_words_add_button = ctk.CTkButton(self.new_window, width=40, height=20, text="+", font=("Arial", 20, "bold"), command=self.button_add_dict_word)
        self.dict_words_add_button.grid(row=3, column=2, padx=5, pady=5, sticky="nsew", columnspan=2)

        self.entry_delete_dict_word = ctk.CTkEntry(self.new_window, placeholder_text="Введите слово для удаления")
        self.entry_delete_dict_word.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")

        self.dict_words_delete_button = ctk.CTkButton(self.new_window, width=40, height=20, text="-", fg_color="red", text_color="black", font=("Arial", 20, "bold"), command=self.button_delete_dict_word)
        self.dict_words_delete_button.grid(row=4, column=1, padx=5, pady=5, sticky="w", columnspan=2)

        self.dict_button1 = ctk.CTkButton(self.new_window, text="Сохранить",)
        self.dict_button1.grid(row=5, column=0, padx=20, pady=20, sticky="s")

        self.dict_button2 = ctk.CTkButton(self.new_window, text="Сохранить как", command=self.save_dict_as)
        self.dict_button2.grid(row=5, column=1, padx=20, pady=20, sticky="s")

        self.dict_table_draw(self.emot_dict)

    def dict_table_draw(self, edict):

        self.dict_table = ctk.CTkScrollableFrame(self.new_window, width=160, height=200)
        self.dict_table.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="w")

        id = 1
        for keys, values in edict.items():
            celld1 = ctk.CTkLabel(self.dict_table, text=keys, anchor="w", fg_color="#2fa3de", text_color="black", corner_radius=5)
            celld1.grid(row=id, column=0, padx=10, pady=5, sticky="w")
            celld2 = ctk.CTkLabel(self.dict_table, text=values, anchor="w", fg_color="#2fa3de", text_color="black", corner_radius=5)
            celld2.grid(row=id, column=1, padx=10, pady=5, sticky="w")
            id += 1

    def load_dict(self):
        filepath = filedialog.askopenfilename()
        if filepath.split('.')[-1] == "pkl":
            with open(filepath, 'rb') as f:
                emot_dict = pickle.load(f)
        else:
            mb.showerror("Ошибка", "Не верный формат словаря, словарь должен иметь формат .pkl")
        if type(emot_dict) == dict:
            self.emot_dict = emot_dict
            self.dict_table_draw(self.emot_dict)
        else:
            mb.showerror("Ошибка", "Струтура файлв не является словарём")

    def button_add_dict_word(self):
        word = self.entry_dict_word.get()
        value = self.dict_range.get()
        self.emot_dict[word] = value
        self.dict_table_draw(self.emot_dict)
        self.entry_dict_word.delete(0, "end")

    def button_delete_dict_word(self):
        word = self.entry_delete_dict_word.get()
        del self.emot_dict[word]
        self.dict_table_draw(self.emot_dict)
        self.entry_delete_dict_word.delete(0, "end")

    def save_dict_as(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".pkl")
        with open(file_path, 'wb') as f:
            pickle.dump(self.emot_dict, f)

    def create_file_table(self):

        self.files_frame = ctk.CTkFrame(self)
        self.files_frame.grid(row=0, column=0, padx=20, pady=20, sticky="w")

        label_table_frame = ctk.CTkLabel(self.files_frame, text="Выберите файлы", font=("Arial", 12, "bold"))
        label_table_frame.grid(row=0, column=0, sticky="w", padx=5, pady=5)

        label_table_frame = ctk.CTkLabel(self.files_frame, text="Файлы должны быть форматов .docx, .pdf, .txt", font=("Arial", 12))
        label_table_frame.grid(row=1, column=0, sticky="w", padx=20, pady=5)

        self.file_button = ctk.CTkButton(self.files_frame, text="Обзор файлов", command=self.button_ask_open_file)
        self.file_button.grid(row=2, column=0, padx=20, pady=20, sticky="w", columnspan=2)

        self.file_table = ctk.CTkScrollableFrame(self.files_frame, width=200, height=90)
        self.file_table.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="w")

        self.header_file_table = ctk.CTkFrame(self.file_table, fg_color="transparent")
        self.header_file_table.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        label_file_table_name = ctk.CTkLabel(self.header_file_table, text="Файл", font=("Arial", 12, "bold"))
        label_file_table_name.pack(side="left", padx=(0, 4))

    def create_dict_chosing(self, dict_name="По умолчанию"):
        self.dict_frame = ctk.CTkFrame(self)
        self.dict_frame.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")
        label = ctk.CTkLabel(self.dict_frame, text="Выберите словарь", font=("Arial", 12, "bold"))
        label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        label = ctk.CTkLabel(self.dict_frame, text=f"Выбранный словарь: {dict_name}", font=("Arial", 12))
        label.grid(row=1, column=0, sticky="w", padx=20, pady=5)
        self.dict_button = ctk.CTkButton(self.dict_frame, text="Выбрать словарь", command=self.button_ask_open_dict)
        self.dict_button.grid(row=2, column=0, padx=20, pady=5, sticky="w", columnspan=2)

    def create_keywords_list(self):
        self.keywords_frame = ctk.CTkFrame(self)
        self.keywords_frame.grid(row=0, column=1, padx=20, pady=20, sticky="ne")

        self.label_keywords_frame = ctk.CTkLabel(self.keywords_frame, text="Ключевые слова", font=("Arial", 12, "bold"))
        self.label_keywords_frame.grid(row=0, column=0, sticky="w", padx=5, pady=5)

        self.entry_keywords = ctk.CTkEntry(self.keywords_frame, placeholder_text="Введите ключевое слово")
        self.entry_keywords.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")

        self.keywords_add_button = ctk.CTkButton(self.keywords_frame, width=40, height=20, text="+", font=("Arial", 20, "bold"), command=self.button_add_keyword)
        self.keywords_add_button.grid(row=1, column=1, padx=5, pady=5, sticky="w", columnspan=2)

        self.keywords_table = ctk.CTkScrollableFrame(self.keywords_frame, width=300, height=40)
        self.keywords_table.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="w")

        self.header_keywords_table = ctk.CTkFrame(self.keywords_table, fg_color="transparent")
        self.header_keywords_table.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        label_keywords_table = ctk.CTkLabel(self.header_keywords_table, text="Ключевое слово", font=("Arial", 12, "bold"))
        label_keywords_table.pack(side="left", padx=(0, 4))

        i = 1
        for kw in self.keywords:
            cell = ctk.CTkLabel(self.keywords_table, text=kw, anchor="w", fg_color="#2fa3de", text_color="black", corner_radius=5)
            cell.grid(row=i, column=0, padx=10, pady=5, sticky="w")
            i += 1

        self.keywords_delete_button = ctk.CTkButton(self.keywords_frame, text="Очистить", fg_color="pink", text_color="black", font=("Arial", 12, "bold"), command=self.button_delete_keyword)
        self.keywords_delete_button.grid(row=3, column=0, padx=5, pady=5, sticky="w", columnspan=2)

    def create_settings(self):
        self.settings_frame = ctk.CTkFrame(self)
        self.settings_frame.grid(row=1, column=1, padx=20, pady=20, sticky="se")
        label = ctk.CTkLabel(self.settings_frame, text="Настройки и запуск", font=("Arial", 12, "bold"))
        label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        label = ctk.CTkLabel(self.settings_frame, text=f"Выберите размер контекста", font=("Arial", 12))
        label.grid(row=1, column=0, sticky="w", padx=20, pady=5)
        ranges=["1","2","3","4","5","6"]
        self.settings_range = ctk.CTkOptionMenu(self.settings_frame, values=ranges, command=self.set_range)
        self.settings_range.grid(row=2, column=0, padx=5, pady=5, sticky="n")

        self.keywords_add_button = ctk.CTkButton(self.settings_frame, text="Начать анализ", font=("Arial", 20, "bold"), command=self.start_analyze)
        self.keywords_add_button.grid(row=2, column=1, padx=5, pady=5, sticky="n", columnspan=2)

    def set_range(self, value):
        self.search_range = value
        print(self.search_range)

    def start_analyze(self):
        reader = FileReader()
        reader.read_files(self.filepaths)

        preparer = TextPrepare()
        preparer.prepare(reader.files_text)

        analyzer = Analyzer(self.dictpath)
        analyzer.analyze(preparer.prep_text_dict,self.keywords, int(self.search_range))

        docs_full = analyzer.analyzed_data.keys()
        docs = []
        for doc in docs_full:
            docs.append(doc.split('/')[-1])
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

    def button_add_keyword(self):
        if self.entry_keywords.get() != "":
            self.keywords.append(self.entry_keywords.get())
            self.header_keywords_table = ctk.CTkFrame(self.keywords_table, fg_color="transparent")
            self.header_keywords_table.grid(row=0, column=0, padx=10, pady=5, sticky="w")
            label_keywords_table = ctk.CTkLabel(self.header_keywords_table, text="Ключевое слово", font=("Arial", 12, "bold"))
            label_keywords_table.pack(side="left", padx=(0, 4))

            i = 1
            for kw in self.keywords:
                cell = ctk.CTkLabel(self.keywords_table, text=kw, anchor="w", fg_color="#2fa3de", text_color="black", corner_radius=5)
                cell.grid(row=i, column=0, padx=10, pady=5, sticky="w")
                i += 1
            self.entry_keywords.delete(0, "end")


    def button_delete_keyword(self):
        if self.keywords != []:
            self.keywords.clear()
            self.create_keywords_list()

    def button_ask_open_file(self):

        self.create_file_table()

        self.filepaths = filedialog.askopenfilenames()
        id = 1
        for file in self.filepaths:
            file_name = file.split('/')[-1]
            cell2 = ctk.CTkLabel(self.file_table, text=file_name, anchor="w", fg_color="#2fa3de", text_color="black", corner_radius=5)
            cell2.grid(row=id, column=0, padx=10, pady=5, sticky="w")
            id += 1
    
    def button_ask_open_dict(self):

        dictpath = filedialog.askopenfilename()

        if dictpath.split('.')[-1] != "pkl":
            mb.showerror("Ошибка", "Не верный формат словаря, словарь должен иметь формат .pkl")
        else:
            self.dictpath = dictpath
            name = dictpath.split('/')[-1]
            self.create_dict_chosing(name)


app = App()
app.mainloop()