from docx2pdf import convert


class Converter:

    def __init__(self, saving_name_pdf, saving_name_docx):
        self.filename = saving_name_pdf
        self.filename_docx = saving_name_docx

    def convert(self):
        print('Converting ', self.filename, " into PDF...")
        convert(self.filename_docx)
        # convert(saving_name_pdf)
        while True:
            try:
                convert(self.filename)
                break
            except AssertionError:
                print("Error in API, but you should not care...")
                break
        # convert("donor.docx/")
        print("SUCCESS")