import ezgmail


class EmailSender:

    def __init__(self, receiver_name, email, job, gender, pdf):
        self.receiver_name = receiver_name
        self.email = email
        self.job = job
        self.gender = gender
        self.pdf = pdf

    def send_email(self):
        body = self.write_body()
        ezgmail.send(
            self.email,
            'Bewerbung als ' + self.job,
            body,
            ['lebenslauf_CV_mamadzhanov.pdf', self.pdf, 'Diploma_Bachelor.pdf']
        )
        print("Sent to the " + self.email + ". Success!")

    def write_body(self):
        main_text = """
im Anhang zu dieser E-Mail finden Sie meine Bewerbungsunterlagen für die Stelle. Rückfragen stehe ich Ihnen gerne zur Verfügung.
"""
        signature = """
Mit freundlichen Grüßen, 
Bobur Mamadzhanov, BSc

📱 tel : +43 676 508 7535
"""

        if self.gender == "Male":
            heading = "Sehr geehrter "
        else:
            heading = "Sehr geehrte "

        return heading + self.receiver_name + ",\r\n" + main_text + "\r\n" + signature


# class Main:
#     bo = "hello"
#
#     print(bo)
#     bo = "world, " + bo + " world!"
#     print(bo)
#     a = EmailSender("Herr Bob", "95babur@gmail.com", "Geber", "Male", "delete.pdf")
#     print(a.write_body())
#     a.send_email()


