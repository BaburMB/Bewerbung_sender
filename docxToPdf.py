from docx2pdf import convert


class Converter:

    def __init__(self, saving_name_pdf, saving_name_docx):
        self.filename = saving_name_pdf
        self.filename_docx = saving_name_docx

    def convert(self):
        print('Converting ', self.filename, " into PDF...")
        convert(self.filename_docx)
        while True:
            try:
                convert(self.filename)
                break
            except AssertionError:
                print("Error in 3rd-party API, but you should not care...")
                break

        print("SUCCESS")