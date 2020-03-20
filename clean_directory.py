import os
import shutil


class Mover:

    def __init__(self, saving_name_docx, saving_name_pdf, firma, email):
        self.pdf = saving_name_pdf
        self.docx = saving_name_docx
        self.firma = firma
        self.email = email

    def clean(self):
        if self.email == "":
            print("Not send. Find the files in a root directory")
        elif self.firma != "d":
            saving_directory = 'sent_emails/' + self.firma + " " + self.email + '/'
            shutil.move(self.pdf, saving_directory + self.pdf)
            shutil.move(self.docx, saving_directory + self.docx)
            print("Saved documents in " + saving_directory)

        else:
            saving_directory = 'sent_emails/' + "NO firma " + self.email + '/'
            shutil.move(self.pdf, saving_directory + self.pdf)
            shutil.move(self.docx, saving_directory + self.docx)
            print("Saved documents in " + saving_directory)

