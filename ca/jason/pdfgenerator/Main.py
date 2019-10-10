from VolunteerExcelFile import VolunteerExcelFile
from WordFile import WordFile
from Input import Input
from File_Converter import File_Converter
from Email import Email 
import os

VOLUNTEER_NUMBER = 0
VOLUNTEER_FIRST_NAME= 1
VOLUNTEER_LAST_NAME= 2
VOLUNTEER_EMAIL = 3
VOLUNTEER_HOUR= 4
VOLUNTEER_POSITION = 5
VOLUNTEER_COMMENT = 6
VOLUNTEER_CERT_LOC = 7
def main(word_template_filename, volunteer_hour_excel_filename, save_file_path, event_name, date):

    volunteer_excel = VolunteerExcelFile(volunteer_hour_excel_filename)
    volunteer_excel.get_file_content()
    volunteer_data = volunteer_excel.get_data()

    
    for key in volunteer_data:
        vol_cert_template = WordFile(word_template_filename)
        vol_cert_template.modify_content('{{Name}}', key)
        vol_cert_template.modify_content('{{Hours}}', str(volunteer_data[key][VOLUNTEER_HOUR]))
        vol_cert_template.modify_content('{{Event Name}}', event_name)
        vol_cert_template.modify_content('{{Date}}', date)
        vol_cert_template.modify_content('{{Position}}', str(volunteer_data[key][VOLUNTEER_POSITION]))
        vol_cert_template.modify_content('{{Comment}}', str(volunteer_data[key][VOLUNTEER_COMMENT]))
        vol_cert_template.set_font('Helvetica Neue', 13)
        new_file_name = save_file_path + "\/"+ event_name + ' ' + key + ' Volunteer Certificate.docx'
        print(new_file_name)
        vol_cert_template.save(new_file_name)
        File_Converter().word_to_pdf(new_file_name)
        volunteer_data[key].append(new_file_name)
        e = Email('jasonwei0224@gmail.com', volunteer_data[key][VOLUNTEER_EMAIL],"TEST","", new_file_name)
        e.send_mail()

    print(volunteer_data['Jason Wei'][VOLUNTEER_CERT_LOC])

if __name__ == "__main__":

    word_template_filename = Input().get_file_path("Word File", ".doc, .docx")
    save_file_path = Input().get_folder("Select which folder to save to")
    volunteer_hour_excel_filename = Input().get_file_path("Excel File", ".xlsx")
    event_name = Input().get_txt_input("Enter the Event Name ")
    date = Input().get_txt_input("Enter the Event Dates")
    main(word_template_filename, volunteer_hour_excel_filename, save_file_path, event_name, date)
    
    Input().show_popup("Completed!")
    
