from spire.doc import *
from spire.doc.common import *
from pypdf import PdfReader

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
