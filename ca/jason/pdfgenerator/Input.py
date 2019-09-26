import PySimpleGUI

class Input:

    def __init__(self):
        pass

    def get_file_path(self, file_type_name: str, file_type_extension: str) -> str:
        """
        Display a dialog box for user to browse and get the file name and path 

        file_type_name: the string name of the file type that's being accepted
        file_type_extension: the string of the extension a file has (e.g. .exe)
        """
        file_path = PySimpleGUI.PopupGetFile("Select " + file_type_name,
                                             file_types=((file_type_name, file_type_extension),))
        return file_path

    def get_txt_input(self, input_hint: str) -> str:
        """
        Display a dialog box to get text input from user

        intput_hint: the string hint that is display about what the user needs to input
        """
        text = PySimpleGUI.PopupGetText(input_hint)

        return text

    def get_folder(self, input_hint: str) -> str:
        """
        Display a dialog box for user to borwse a folder

        input_hint: the string hint that is display 
        """
        folder = PySimpleGUI.PopupGetFolder(input_hint)
        return folder

    def show_popup(self, message: str) -> None:
        """
        shows a pop up with message

        message: message that is display on the dialog box 
        """
        PySimpleGUI.Popup(message)
