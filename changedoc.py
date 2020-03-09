from docx import Document
from docx.shared import Pt
from docx.shared import Inches
from docx2pdf import convert


print("Initializing donor document...")
doc = Document('donor.docx')                    #read default document
print("Initializing donor document... OK")

run = doc.add_paragraph().add_run()
font = run.font
font.name = "Arial"
font.size = Pt(12)

d = "d"
gruss_receiver = input(">> WHOM are you sending? ")
if gruss_receiver == d: gruss_receiver = "Damen und Herren"
print("Okay, sending to ", gruss_receiver)

job = input(">> What kind of JOB is that? ")
if job == d: job = "Mitarbeiter"
print("Okay, your possible job is ", job)

firma_receiver = input(">> What FIRMA is that? ")
if firma_receiver == d: firma_receiver = "Ihrer Firma"
print("Okay, let's write ", firma_receiver)

print("Parsing paragraph...")
paragraph = doc.paragraphs[2].text
print("Parsing paragraph... OK")

print(gruss_receiver[0])
gruss_receiver_gender = d

if gruss_receiver[0] == "F" or gruss_receiver[0] == "f":
    gruss_receiver_gender = "Female"
elif gruss_receiver[0] == "H" or gruss_receiver[0] == "h":
    gruss_receiver_gender = "Male"
else:
    print("Couldn't detect gender for ", gruss_receiver[0], ", will use default gender for Grüß")
print("Receiver's gender is ", gruss_receiver_gender)


if gruss_receiver_gender == "Male":
    paragraph = paragraph.replace("%%receiver%%", "r " + gruss_receiver)
    print("Putting male name... OK")

elif gruss_receiver_gender == "Female":
    paragraph = paragraph.replace("%%receiver%%", " " + gruss_receiver)
    print("Putting female name... OK")
else:
    paragraph = paragraph.replace("%%receiver%%", " " + gruss_receiver)
    print("Putting default name... OK")

paragraph = paragraph.replace("%%firma%%", firma_receiver)
print("Putting firma name... OK")
paragraph = paragraph.replace("%%job%%", job)
print("Putting job name... OK")
doc.paragraphs[2].text = paragraph
print("Pasting text... OK")
doc.paragraphs[2].style = "bewerbung_default_paragraph"
print("Applying style... OK")

# saving_name = job + ".docx"
saving_name = "Anschreiben als " + job + "in " + firma_receiver
if gruss_receiver != d: saving_name = saving_name + " für " + gruss_receiver
saving_name_docx = saving_name + ".docx"
saving_name_pdf = saving_name + ".pdf"
doc.save(saving_name_docx)
# print("Saved as ", saving_name)
# print(doc.paragraphs[2].text)

#####CONVERTING DOCX INTO PDF########

# wdFormatPDF = 17
#
# in_file = saving_name_docx
# out_file = saving_name_pdf
#
#
# # word = comtypes.client.CreateObject('Word.Application')
# word = win32com.client.Dispatch('Word.Application')
# document = word.Documents.Open(in_file)
# document.SaveAs(out_file, FileFormat=wdFormatPDF)
# document.Close()
# word.Quit()

print('Converting ', saving_name_pdf, " into PDF...")
convert(saving_name_docx)
# convert(saving_name_pdf)
while True:
    try:
        convert(saving_name_pdf)
        break
    except AssertionError:
        print("Error in API, but you shoudn't care...")
        break
# convert("my_docx_folder/")
print("SUCCESS")