import comtypes.client
import os


class File_Converter:
    """
    File Converter Class
    This class helps convert file to different file types
    """
    def __init__(self):
        pass

    def word_to_pdf(self, filename) -> int:
        """
        Converts given word file to pdf

        """
        PDFformat = 17

        word = comtypes.client.CreateObject('Word.Application')
        word_doc = word.Documents.Open(os.path.abspath(filename))
        word_doc.SaveAs(os.path.abspath(filename + ".pdf"), FileFormat = PDFformat)
        word_doc.Close()
        word.Quit()
        return 0
