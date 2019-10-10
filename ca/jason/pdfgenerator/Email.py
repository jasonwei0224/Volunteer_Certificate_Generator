
import smtplib
import email
import email.mime.application
from email.mime.multipart import MIMEMultipart

class Email:

    def __init__(self, from_address, to_address, subject, cc, filename):
        self.__from_address = from_address
        self.__to_address = to_address
        self.__cc = cc

        self.__subject = subject
        self.__message = MIMEMultipart()
        self.set_header()
        self.__attachments = None
        self.set_attachment(filename, "pdf")
        self.__sever = smtplib.SMTP('smtp.gmail.com', 587)


    def set_to_address(self, to_address):
        self.__to_address = to_address

    def get_message(self):
        return self.__message

    def set_header(self):
        self.__message['Subject'] = self.__subject
        self.__message['From'] = self.__from_address
        self.__message['To'] = self.__to_address
        self.__message['Cc'] = self.__cc

    def set_body(self, body):
        self.__message.attach(body)

    def set_attachment(self, filename, type):
        file = open(filename, 'rb')
        self.__attachments = email.mime.application.MIMEApplication(file.read(), _subtype= type)
        file.close()
        new_name = filename.split("\/")

        self.__attachments.add_header('Content-Disposition', 'attachment', filename = new_name[1])
        self.__message.attach(self.__attachments)

    def send_mail(self):
        self.__sever.starttls()
        self.__sever.login('jasonwei0224@gmail.com', 'Ja1241$$')
        self.__sever.sendmail(self.__from_address, self.__to_address, self.__message.as_string())
