import PySimpleGUI

class Input:

    def __init__(self):
        pass

    def get_file_path(self, file_type_name, file_type_extension):
        file_path = PySimpleGUI.PopupGetFile("Select " + file_type_name,
                                             file_types=((file_type_name, file_type_extension),))
        return file_path

    def get_txt_input(self, input_hint):
        text = PySimpleGUI.PopupGetText(input_hint)

        return text

    def get_folder(self, input_hint):
        folder = PySimpleGUI.PopupGetFolder(input_hint)
        return folder

    def show_popup(self, message):
        PySimpleGUI.Popup(message)