from models.TextPrepare import TextPrepare
from models.Analyzer import Analyzer
from models.FileReader import FileReader
import customtkinter as ctk
from tkinter import filedialog
from tkinter import messagebox as mb
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tksheet
import pickle

class App(ctk.CTk):
    filepaths = []
    dictpath = "dicts/dictionary.pkl" # ВСТАВИТЬ ДЕФОЛТ ВАРИАНТ
    keywords = []
    search_range_l = "1"
    search_range_r = "1"
    emot_dict = {}
    emot_dict_path = ""
    lang = "Русский"
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

    def dict_redactor(self):
        if self.dict_redactor_closed_flag:
            self.dict_redactor_closed_flag = False
            self.new_window = ctk.CTkToplevel(self)
            self.new_window.title("Редактор")
            self.new_window.geometry("800x600")
            self.new_window.resizable(False, False)

            self.label_dict_red = ctk.CTkLabel(self.new_window, text="Выберите словарь", font=("Arial", 12, "bold"))
            self.label_dict_red.grid(row=0, column=0, sticky="w", padx=5, pady=5)
            self.dict_button = ctk.CTkButton(self.new_window, text="Выбрать словарь", command=self.load_dict)
            self.dict_button.grid(row=1, column=0, padx=5, pady=5, sticky="nsew", columnspan=2)

            self.dict_table = ctk.CTkFrame(self.new_window)
            self.dict_table.grid(row=2, column=0, padx=5, pady=5, sticky="nsew")

            self.dict_frame = ctk.CTkFrame(self.new_window)
            self.dict_frame.grid(row=2, column=1, padx=5, pady=5, sticky="nsew")

            self.dict_button1 = ctk.CTkButton(self.new_window, text="Сохранить", command=self.save_dict)
            self.dict_button1.grid(row=5, column=0, padx=20, pady=20, sticky="nsew")

            self.dict_button2 = ctk.CTkButton(self.new_window, text="Сохранить как", command=self.save_dict_as)
            self.dict_button2.grid(row=5, column=1, padx=20, pady=20, sticky="nsew")

            self.dict_sort1 = ctk.CTkButton(self.dict_frame, text="Отсортировать\n по словам", command=self.sort_word)
            self.dict_sort1.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

            self.dict_sort2 = ctk.CTkButton(self.dict_frame, text="Отсортировать\n по оценкам", command=self.sort_value)
            self.dict_sort2.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")

            self.dict_table_draw()
            self.new_window.protocol("WM_DELETE_WINDOW", self.closing_dict)
            self.new_window.after(80, self.new_window.lift)

    def closing_dict(self):
        self.dict_redactor_closed_flag = True
        self.new_window.destroy()

    def dict_table_draw(self):
        self.sheet = tksheet.Sheet(self.dict_table, sort_key=tksheet.natural_sort_key)
        self.sheet.enable_bindings(("single_select",
                                    "row_select",
                                    "column_width_resize",
                                    "arrowkeys",
                                    "right_click_popup_menu",
                                    "rc_select",
                                    "rc_insert_row",
                                    "rc_delete_row",
                                    "copy",
                                    "cut",
                                    "paste",
                                    "delete",
                                    "undo",
                                    "edit_cell"))
        self.sheet.set_options(edit_cell_label= "Редактировать",
                                cut_label= "Вырезать",
                                cut_contents_label= "Скопировать",
                                copy_label= "Скопировать",
                                paste_label= "Вставить",
                                delete_label= "Удалить",
                                clear_contents_label= "Очистить содержимое",
                                delete_rows_label= "Удалить строки",
                                insert_rows_above_label= "Вставить строку сверху",
                                insert_rows_below_label= "Вставить строку снизу",
                                insert_row_label= "Вставить строку",
                                )
        self.sheet.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.sheet.headers(["Слово", "Оценка"])
        self.sheet.dropdown("B", values=[2, 1, -1, -2])
        table = [["Пример", "1"]]
        for key, value in self.emot_dict.items():
            table.append([key, value])
        self.sheet.set_sheet_data(table)

    def sort_word(self):
        self.sheet.sort_rows_by_column(0)

    def sort_value(self):
        self.sheet.sort_rows_by_column(1, True)

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
            if type(emot_dict) == dict:
                self.emot_dict = emot_dict
                self.emot_dict_path = filepath
                self.add_data_table(self.emot_dict)
                name = filepath.split('/')[-1]
                self.label_dict_red.configure(text=f"Выбранный словарь: {name}")
            else:
                mb.showerror("Ошибка", "Струтура файлв не является словарём")
        else:
            mb.showerror("Ошибка", "Не верный формат словаря, словарь должен иметь формат .pkl")

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

        self.docs_frame = ctk.CTkFrame(self)
        self.docs_frame.grid(row=0, column=0, padx=20, pady=20, sticky="w")

        label_docs_frame = ctk.CTkLabel(self.docs_frame, text="Выберите файлы", font=("Arial", 12, "bold"))
        label_docs_frame.grid(row=0, column=0, sticky="w", padx=5, pady=5)

        label_table_docs = ctk.CTkLabel(self.docs_frame, text="Файлы должны быть форматов .docx, .pdf, .txt", font=("Arial", 12))
        label_table_docs.grid(row=1, column=0, sticky="w", padx=20, pady=5)

        self.docs_button = ctk.CTkButton(self.docs_frame, text="Обзор файлов", command=self.button_ask_open_file)
        self.docs_button.grid(row=2, column=0, padx=20, pady=20, sticky="w", columnspan=2)

        self.docs_table = ctk.CTkScrollableFrame(self.docs_frame, width=200, height=90)
        self.docs_table.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="w")

        self.header_docs_table = ctk.CTkFrame(self.docs_table, fg_color="transparent")
        self.header_docs_table.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.label_docs_table_name = ctk.CTkLabel(self.header_docs_table, text="№", font=("Arial", 12, "bold"))
        self.label_docs_table_name.pack(side="left", padx=(0, 4))

        self.header_docs_table = ctk.CTkFrame(self.docs_table, fg_color="transparent")
        self.header_docs_table.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        self.label_docs_table_name = ctk.CTkLabel(self.header_docs_table, text="Файл", font=("Arial", 12, "bold"))
        self.label_docs_table_name.pack(side="left", padx=(0, 4))

    def create_dict_chosing(self, dict_name="По умолчанию"):
        self.dict_frame = ctk.CTkFrame(self)
        self.dict_frame.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")
        label = ctk.CTkLabel(self.dict_frame, text="Выберите словарь", font=("Arial", 12, "bold"))
        label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        label = ctk.CTkLabel(self.dict_frame, text=f"Выбранный словарь: {dict_name}", font=("Arial", 12))
        label.grid(row=2, column=0, sticky="w", padx=20, pady=5)
        self.dict_button = ctk.CTkButton(self.dict_frame, text="Выбрать словарь", command=self.button_ask_open_dict)
        self.dict_button.grid(row=1, column=0, padx=20, pady=5, sticky="w", columnspan=2)
        self.dict_redac_button = ctk.CTkButton(self.dict_frame, text="Редактор словарей", command=self.dict_redactor)
        self.dict_redac_button.grid(row=3, column=0, padx=20, pady=(60,5), sticky="sw", columnspan=2)

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

        self.keywords_delete_button = ctk.CTkButton(self.keywords_frame, text="Очистить", fg_color="#d36f6f", text_color="black", font=("Arial", 12, "bold"), command=self.button_delete_keyword)
        self.keywords_delete_button.grid(row=3, column=0, padx=5, pady=5, sticky="w", columnspan=2)

    def create_settings(self):
        self.settings_frame = ctk.CTkFrame(self)
        self.settings_frame.grid(row=1, column=1, padx=20, pady=20, sticky="se")
        label = ctk.CTkLabel(self.settings_frame, text="Настройки и запуск", font=("Arial", 12, "bold"))
        label.grid(row=0, column=0, sticky="w", padx=5, pady=5)

        ranges=["1","2","3","4","5","6"]
        label = ctk.CTkLabel(self.settings_frame, text=f"Выберите размер левого контекста", font=("Arial", 12))
        label.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.settings_range_l = ctk.CTkOptionMenu(self.settings_frame, values=ranges, command=self.set_range_l)
        self.settings_range_l.grid(row=2, column=0, padx=5, pady=5, sticky="n")

        label = ctk.CTkLabel(self.settings_frame, text=f"Выберите размер правого контекста", font=("Arial", 12))
        label.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        self.settings_range_r = ctk.CTkOptionMenu(self.settings_frame, values=ranges, command=self.set_range_r)
        self.settings_range_r.grid(row=2, column=1, padx=5, pady=5, sticky="n")

        label = ctk.CTkLabel(self.settings_frame, text=f"Выберите язык документов", font=("Arial", 12))
        label.grid(row=3, column=0, sticky="w", padx=5, pady=5)
        ranges=["Русский", "Английский"]
        self.settings_lang = ctk.CTkOptionMenu(self.settings_frame, values=ranges, command=self.set_lang)
        self.settings_lang.grid(row=4, column=0, padx=5, pady=5, sticky="n")


        self.keywords_add_button = ctk.CTkButton(self.settings_frame, text="Начать анализ", font=("Arial", 20, "bold"), command=self.start_analyze)
        self.keywords_add_button.grid(row=5, column=0, padx=5, pady=5, sticky="s", columnspan=2)

    def set_range_l(self, value):
        self.search_range_l = value

    def set_range_r(self, value):
        self.search_range_r = value

    def set_lang(self, value):
        self.lang = value

    def start_analyze(self):

        if not self.results_open_flag:
            print(self.filepaths)
            if self.filepaths == [] or self.filepaths == '':
                mb.showerror("Ошибка", "Список файлов слов пуст")
            elif self.keywords == []:
                mb.showerror("Ошибка", "Список ключевых слов пуст")
            else:
                if self.lang == "Русский":
                    reader = FileReader()
                    reader.read_files(self.filepaths)

                    preparer = TextPrepare()
                    preparer.prepare(reader.files_text)

                    analyzer = Analyzer(self.dictpath)
                    analyzer.analyze(preparer.prep_text_dict, self.keywords, int(self.search_range_l), int(self.search_range_r))

                elif self.lang == "Английский":
                    reader = FileReader()
                    reader.read_files(self.filepaths)

                    preparer = TextPrepare()
                    preparer.prepare_en(reader.files_text)

                    analyzer = Analyzer(self.dictpath)
                    analyzer.analyze_en(preparer.prep_text_dict,self.keywords, int(self.search_range_l), int(self.search_range_r))

                self.open_results_window(analyzer.analyzed_data)

                self.results_open_flag = True

    def open_results_window(self, data):
        self.results_open_flag = False
        self.results_window = ctk.CTkToplevel(self)
        self.results_window.title("Результаты")
        self.results_window.geometry("1200x900")
        self.results_window.resizable(False, False)

        self.result_frame = ctk.CTkFrame(self.results_window)
        self.result_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        self.sheet_result = tksheet.Sheet(self.results_window)
        self.sheet_result.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.sheet_result.headers(['Документ', 'Вхождений всего', 'Процент Вхождений', 'Вхождений 2', 'Вхождений 1', 'Вхождений -1', 'Вхождений -2',])
        self.sheet_result.enable_bindings(("column_width_resize",))

        table_data = []
        for key, value in data.items():
            elem = [key, value[0], value[3], value[2], value[1], value[-1], value[-2]]
            table_data.append(elem)
        self.sheet_result.set_sheet_data(table_data)
        self.sheet_result.highlight_columns(3, fg="#03a8a0")
        self.sheet_result.highlight_columns(4, fg="#3FA703")
        self.sheet_result.highlight_columns(5, fg="#fa6323")
        self.sheet_result.highlight_columns(6, fg="#ad0000")

        self.results_window.protocol("WM_DELETE_WINDOW", self.closing_res)

        docs_full = data.keys()
        docs = []
        for doc in docs_full:
            name = doc.split('/')[-1]
            if len(name) > 20:
                name = name[0:7] + "..." + name[-7:-1] + name[-1]
                print(name)
            docs.append(name)
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
        ax1.plot(docs, count, '-', label='Вхождений всего', color='blue', marker='o', markersize=3)
        ax1.set_xlabel('Документы')
        ax1.set_title('Количество вхождений')
        ax1.set_xticklabels(docs, minor=False, rotation=90)

        fig2, ax2 = plt.subplots()

        ax2.plot(docs, count_procents, '-', label='Процентное соотношение', color='blue', marker='o', markersize=3)
        ax2.set_xlabel('Документы')
        ax2.set_title('Процент вхождений')
        ax2.set_xticklabels(docs, minor=False, rotation=90)

        fig3, ax3 = plt.subplots()
        ax3.plot(docs, count_pos1, '-', label='Вхождений с оценкой 1', color="#3FA703", marker='o', markersize=5)
        ax3.plot(docs, count_pos2, '--', label='Вхождений с оценкой 2', color="#03a8a0", marker='o', markersize=6)
        ax3.plot(docs, count_neg1, '-.', label='Вхождений с оценкой -1', color="#fa6323", marker='o', markersize=4)
        ax3.plot(docs, count_neg2, ':', label='Вхождений с оценкой -2', color="#ad0000", marker='o', markersize=3)
        ax3.set_xlabel('Документы')
        ax3.set_title('Оценки вхождений')
        ax3.legend(fontsize=8)
        ax3.set_xticklabels(docs, minor=False, rotation=90)

        fig1.set_size_inches(3.8, 5)
        fig2.set_size_inches(3.8, 5)
        fig3.set_size_inches(3.8, 5)
        fig1.tight_layout()
        fig2.tight_layout()
        fig3.tight_layout()

        self.results_graph = ctk.CTkFrame(self.results_window)
        self.results_graph.grid(row=1, column=0, padx=0, pady=0, sticky="nsew")

        self.results_graph1 = ctk.CTkFrame(self.results_graph)
        self.results_graph1.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        self.results_graph2 = ctk.CTkFrame(self.results_graph)
        self.results_graph2.grid(row=0, column=1, padx=0, pady=5, sticky="nsew")

        self.results_graph3 = ctk.CTkFrame(self.results_graph)
        self.results_graph3.grid(row=0, column=2, padx=5, pady=5, sticky="nsew")

        canvas1 = FigureCanvasTkAgg(fig1, self.results_graph1)
        canvas1.get_tk_widget().grid(row=0, column=0, padx=0, pady=0, sticky="nsew")

        canvas2 = FigureCanvasTkAgg(fig2, self.results_graph3)
        canvas2.get_tk_widget().grid(row=0, column=0, padx=0, pady=0, sticky="nsew")

        canvas3 = FigureCanvasTkAgg(fig3, self.results_graph2)
        canvas3.get_tk_widget().grid(row=0, column=0, padx=0, pady=0, sticky="nsew")
        self.results_window.after(80, self.results_window.lift)


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
            cell1 = ctk.CTkLabel(self.docs_table, text=id, anchor="w", fg_color="#2fa3de", text_color="black", corner_radius=5)
            cell1.grid(row=id, column=0, padx=10, pady=5, sticky="nsew")
            cell2 = ctk.CTkLabel(self.docs_table, text=file_name, anchor="w", fg_color="#2fa3de", text_color="black", corner_radius=5)
            cell2.grid(row=id, column=1, padx=10, pady=5, sticky="nsew")
            id += 1

    def button_ask_open_dict(self):

        dictpath = filedialog.askopenfilename()
        if dictpath == "":
            pass
        elif dictpath.split('.')[-1] != "pkl":
            mb.showerror("Ошибка", "Не верный формат словаря, словарь должен иметь формат .pkl")
        else:
            self.dictpath = dictpath
            name = dictpath.split('/')[-1]
            self.create_dict_chosing(name)


app = App()
app.mainloop()