from main import FileReader, TextPrepare, Analyzer, DictWork
import customtkinter as ctk
from tkinter import filedialog
from tkinter import messagebox as mb
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tksheet
from CTkTable import CTkTable
import pickle

class App(ctk.CTk):
    filepaths = []
    dictpath = "C:\Dev\DIPLOM\dicts\dictionary.pkl" # ВСТАВИТЬ ДЕФОЛТ ВАРИАНТ
    keywords = []
    search_range = "1"
    emot_dict = {}
    emot_dict_path = ""
    lang = "ru"
    dict_redactor_closed_flag = True
    results_open_flag = False

    def __init__(self):
        super().__init__()

        self.title("Main")
        self.geometry("800x700")
        self.grid_columnconfigure((0, 1), weight=1)
        self.resizable(False, False)

        self.create_dict_chosing()

        self.create_file_table()

        self.create_keywords_list()

        self.create_settings()

        self.file_button = ctk.CTkButton(self, text="Редактор словарей", command=self.dict_redactor)
        self.file_button.grid(row=4, column=0, padx=20, pady=20, sticky="w", columnspan=2)

    def dict_redactor(self):
        if self.dict_redactor_closed_flag:
            self.dict_redactor_closed_flag = False
            self.new_window = ctk.CTkToplevel(self)
            self.new_window.title("Редактор")
            self.new_window.geometry("800x600")
            self.new_window.resizable(False, False)

            label_dict_red = ctk.CTkLabel(self.new_window, text="Выберите словарь", font=("Arial", 12, "bold"))
            label_dict_red.grid(row=0, column=0, sticky="w", padx=5, pady=5)
            self.dict_button = ctk.CTkButton(self.new_window, text="Выбрать словарь", command=self.load_dict)
            self.dict_button.grid(row=1, column=0, padx=5, pady=5, sticky="w", columnspan=2)

            self.dict_table = ctk.CTkFrame(self.new_window)
            self.dict_table.grid(row=2, column=0, padx=5, pady=5, sticky="nsew")

            self.dict_button1 = ctk.CTkButton(self.new_window, text="Сохранить", command=self.save_dict)
            self.dict_button1.grid(row=5, column=0, padx=20, pady=20, sticky="s")

            self.dict_button2 = ctk.CTkButton(self.new_window, text="Сохранить как", command=self.save_dict_as)
            self.dict_button2.grid(row=5, column=1, padx=20, pady=20, sticky="s")

            self.dict_table_draw(self.emot_dict)

            #self.test_entry = ctk.CTkEntry(self.new_window)
            #self.test_entry.grid(row=6, column=0, padx=20, pady=20, sticky="nsew")
            #self.test_entry.bind("<Return>", self.change)
            #table.append(self.test_entry)
            #table[0].insert(0,"Text")
            #self.new_window
            # Create the table with write=1 to enable editing
            #self.table = CTkTable(self.new_window, values=self.emot_dict, write=1)
            #self.table.grid(row=6, column=0, padx=20, pady=20, sticky="nsew")

            self.new_window.protocol("WM_DELETE_WINDOW", self.closing_dict)

    def change(self, event):
        print(event)
        print(event.widget.widgetName)
        print(event.widget.grid_info())
        print(self.test_entry.get())

    def closing_dict(self):
        self.dict_redactor_closed_flag = True
        self.new_window.destroy()

    def dict_table_draw(self, flag):

        self.sheet = tksheet.Sheet(self.dict_table)
        self.sheet.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.sheet.headers(["Слово", "Оценка"])
        self.sheet.enable_bindings(("single_select","row_select","column_width_resize","arrowkeys","right_click_popup_menu","rc_select","rc_insert_row","rc_delete_row","copy","cut","paste","delete","undo","edit_cell"))
        self.sheet.dropdown("B", values=[2, 1, -1, -2])
        table = [["Пример", "1"]]
        for key, value in self.emot_dict.items():
            table.append([key, value])
        self.sheet.set_sheet_data(table)

    def add_data_table(self, data):
        self.sheet.set_sheet_data([])
        table = []
        for key, value in self.emot_dict.items():
            table.append([key, value])
        self.sheet.set_sheet_data(table)

    def from_table_to_dict(self):
        edict = {}
        data = self.sheet.get_sheet_data()
        for elem in data:
            if elem[0] != "":
                edict[elem[0]] = elem[1]
        self.emot_dict = edict


    def load_dict(self):
        filepath = filedialog.askopenfilename()
        if filepath.split('.')[-1] == "pkl":
            with open(filepath, 'rb') as f:
                emot_dict = pickle.load(f)
        else:
            mb.showerror("Ошибка", "Не верный формат словаря, словарь должен иметь формат .pkl")
        if type(emot_dict) == dict:
            self.emot_dict = emot_dict
            self.emot_dict_path = filepath
            self.add_data_table(self.emot_dict)
            name = filepath.split('/')[-1]
            label_dict_red = ctk.CTkLabel(self.new_window, text=f"Выбранный словарь:{name}", font=("Arial", 12, "bold"))
            label_dict_red.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        else:
            mb.showerror("Ошибка", "Струтура файлв не является словарём")

    def save_dict(self):
        self.from_table_to_dict()
        if self.emot_dict_path != "":
            with open(self.emot_dict_path, 'wb') as f:
                pickle.dump(self.emot_dict, f)
        self.sheet.set_sheet_data([])

    def save_dict_as(self):

        self.from_table_to_dict()
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

        label = ctk.CTkLabel(self.settings_frame, text=f"Выберите язык документов", font=("Arial", 12))
        label.grid(row=1, column=1, sticky="w", padx=20, pady=5)
        ranges=["ru","en"]
        self.settings_lang = ctk.CTkOptionMenu(self.settings_frame, values=ranges, command=self.set_lang)
        self.settings_lang.grid(row=2, column=1, padx=5, pady=5, sticky="n")


        self.keywords_add_button = ctk.CTkButton(self.settings_frame, text="Начать анализ", font=("Arial", 20, "bold"), command=self.start_analyze)
        self.keywords_add_button.grid(row=3, column=0, padx=5, pady=5, sticky="s", columnspan=2)

    def set_range(self, value):
        self.search_range = value
        print(self.search_range)

    def set_lang(self, value):
        self.lang = value
        print(self.search_range)

    def start_analyze(self):

        if not self.results_open_flag:
            print(self.filepaths)
            if self.filepaths == [] or self.filepaths == '':
                mb.showerror("Ошибка", "Список файлов слов пуст")
            elif self.keywords == []:
                mb.showerror("Ошибка", "Список ключевых слов пуст")
            else:
                if self.lang == "ru":
                    reader = FileReader()
                    reader.read_files(self.filepaths)

                    preparer = TextPrepare()
                    preparer.prepare(reader.files_text)

                    analyzer = Analyzer(self.dictpath)
                    analyzer.analyze(preparer.prep_text_dict, self.keywords, int(self.search_range))

                elif self.lang == "en":
                    reader = FileReader()
                    reader.read_files(self.filepaths)

                    preparer = TextPrepare()
                    preparer.prepare_en(reader.files_text)

                    analyzer = Analyzer(self.dictpath)
                    analyzer.analyze_en(preparer.prep_text_dict,self.keywords, int(self.search_range))

                self.open_results_window(analyzer.analyzed_data)

                self.results_open_flag = True

    def open_results_window(self, data):
        self.dict_redactor_closed_flag = False
        self.results_window = ctk.CTkToplevel(self)
        self.results_window.title("Результаты")
        self.results_window.geometry("1000x800")
        self.results_window.resizable(False, False)

        self.results_table = ctk.CTkScrollableFrame(self.results_window, width=750, height=300)
        self.results_table.grid(row=0, column=0, padx=20, pady=20, sticky="w")

        names = ['Документ', 'Вхождений всего', 'Вхождений 1', 'Вхождений 2', 'Процент Вхождений', 'Вхождений -1', 'Вхождений -2',]
        for i in range(0,7):
            self.header_res_table = ctk.CTkFrame(self.results_table, fg_color="transparent")
            self.header_res_table.grid(row=0, column=i, padx=10, pady=5, sticky="w")
            label_dict_table_name = ctk.CTkLabel(self.header_res_table, text=names[i], font=("Arial", 12, "bold"))
            label_dict_table_name.grid(sticky="nsew",)

        i = 1
        for key, value in data.items():
            i2 = 1
            cell = ctk.CTkLabel(self.results_table, text=key, anchor="w", fg_color="white", text_color="black", corner_radius=5)
            cell.grid(row=i, column=0, padx=10, pady=5, sticky="w")
            for val in value:
                cell = ctk.CTkLabel(self.results_table, text=val, anchor="w", fg_color="white", text_color="black", corner_radius=5)
                cell.grid(row=i, column=i2, padx=10, pady=5, sticky="w")
                i2 += 1
            i += 1

        self.results_window.protocol("WM_DELETE_WINDOW", self.closing_res)

        docs_full = data.keys()
        docs = []
        for doc in docs_full:
            docs.append(doc.split('/')[-1])
        count = []
        count_procents = []
        count_pos1 = []
        count_pos2 = []
        count_neg1 = []
        count_neg2 = []
        for values in data.values():
            count.append(values[0])
            count_procents.append(values[3])
            count_pos1.append(values[1])
            count_pos2.append(values[2])
            count_neg1.append(values[-1])
            count_neg2.append(values[-2])

        fig1, ax1 = plt.subplots()
        ax1.plot(docs, count, ':', label='Вхождений всего', color='blue', marker='o', markersize=3)
        ax1.plot(docs, count_pos1, '-', label='Вхождений с оценкой 1', color='green', marker='o', markersize=5)
        ax1.plot(docs, count_pos2, '--', label='Вхождений с оценкой 2', color='lime', marker='o', markersize=6)
        ax1.plot(docs, count_neg1, '-.', label='Вхождений с оценкой -1', color='darkred', marker='o', markersize=4)
        ax1.plot(docs, count_neg2, ':', label='Вхождений с оценкой -2', color='red', marker='o', markersize=3)
        ax1.set_xlabel('Документы')
        ax1.set_ylabel('Кол-во вхождений')
        ax1.legend(fontsize=9)

        fig2, ax2 = plt.subplots()

        ax2.plot(docs, count_procents, ':', label='Процентное соотношение', color='blue', marker='o', markersize=3)
        ax2.set_xlabel('Документы')
        ax2.set_ylabel('% вхождений')

        fig1.set_size_inches(4, 4)

        fig2.set_size_inches(4, 4)

        self.results_graph = ctk.CTkFrame(self.results_window)
        self.results_graph.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")


        self.results_graph1 = ctk.CTkFrame(self.results_graph)
        self.results_graph1.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.results_graph2 = ctk.CTkFrame(self.results_graph)
        self.results_graph2.grid(row=0, column=1, padx=10, pady=5, sticky="w")

        canvas1 = FigureCanvasTkAgg(fig1, self.results_graph1)
        canvas1.get_tk_widget().grid(row=0, column=0, padx=5, pady=5, sticky="w")

        canvas2 = FigureCanvasTkAgg(fig2, self.results_graph2)
        canvas2.get_tk_widget().grid(row=0, column=0, padx=5, pady=5, sticky="w")


    def closing_res(self):
        self.results_open_flag = False
        self.results_window.destroy()

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