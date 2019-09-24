import docx
from docx.shared import Pt

class WordFile:

    def __init__(self, filename):
        self.__filename = filename
        self.__content = []
        self.__word_doc = docx.Document(filename)

    def get_filename(self):
        return self.__filename

    def get_file_content(self):
        for paragraph in self.__word_doc.paragraphs:
            self.__content.append(''+paragraph.text)
        return self.__content

    def modify_content(self, placeholder, replacement):
        for paragraph in self.__word_doc.paragraphs:
            #paragraph.text.replace(placeholder, replacement)
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, replacement)

    def replace_string(self, place_holder, replacement):
        for index in range(0, len(self.__content)):
            print(self.__content[index])
            self.__content[index] = self.__content[index].replace(place_holder, replacement)

    def save(self, filepath):
        self.__word_doc.save(filepath)


    def set_font(self, name, size):
        style = self.__word_doc.styles['Normal']
        font = style.font
        font.name = name
        font.size = Pt(size)

