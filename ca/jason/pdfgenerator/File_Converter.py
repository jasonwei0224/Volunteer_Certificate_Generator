import comtypes.client
import os


class File_Converter:

    def __init__(self):
        pass

    def word_to_pdf(self, in_filename):

        wdFormatPDF = 17

        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(os.path.abspath(in_filename))
        doc.SaveAs(os.path.abspath(in_filename + ".pdf"), FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
        return 0