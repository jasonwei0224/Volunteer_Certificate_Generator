from ca.jason.pdfgenerator.VolunteerExcelFile import VolunteerExcelFile
from ca.jason.pdfgenerator.WordFile import WordFile
from ca.jason.pdfgenerator.Input import Input
from ca.jason.pdfgenerator.File_Converter import File_Converter
import os
def main(word_template_filename, volunteer_hour_excel_filename, save_file_path, event_name, date):

    volunteer_excel = VolunteerExcelFile(volunteer_hour_excel_filename)
    volunteer_excel.get_file_content()
    volunteer_data = volunteer_excel.get_data()


    for key in volunteer_data:
        vol_cert_template = WordFile(word_template_filename)
        vol_cert_template.modify_content('{{Name}}', key)
        vol_cert_template.modify_content('{{Hours}}', str(volunteer_data[key][0]))
        vol_cert_template.modify_content('{{Event Name}}', event_name)
        vol_cert_template.modify_content('{{Date}}', date)
        vol_cert_template.modify_content('{{Position}}', str(volunteer_data[key][1]))
        vol_cert_template.modify_content('{{Comment}}', str(volunteer_data[key][2]))
        vol_cert_template.set_font('Helvetica Neue', 13)
        new_file_name = save_file_path + "\/"+ event_name + ' ' + key + ' Volunteer Certificate.docx'
        print(new_file_name)
        vol_cert_template.save(new_file_name)


if __name__ == "__main__":

    word_template_filename = Input().get_file_path("Word File", ".doc, .docx")
    save_file_path = Input().get_folder("Select which folder to save to")
    volunteer_hour_excel_filename = Input().get_file_path("Excel File", ".xlsx")
    event_name = Input().get_txt_input("Enter the Event Name ")
    date = Input().get_txt_input("Enter the Event Dates")
    main(word_template_filename, volunteer_hour_excel_filename, save_file_path, event_name, date)

    dir = os.listdir(save_file_path)
    for file in dir:
        File_Converter().word_to_pdf(save_file_path +'/' + file)
    Input().show_popup("Completed!")


