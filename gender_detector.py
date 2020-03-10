class GenderDetector:
    gender = "d"

    def __init__(self, g_receiver):
        self.gruss_receiver = g_receiver

    def detect_gender(self):
        if self.gruss_receiver[0] == "F" or self.gruss_receiver[0] == "f":
            self.gender = "Female"
        elif self.gruss_receiver[0] == "H" or self.gruss_receiver[0] == "h":
            self.gender = "Male"
        else:
            print("Couldn't detect gender for ", self.gruss_receiver[0], ", will use default gender for Sehr...")
        print("Receiver's gender is ", self.gender)
        # self.condition = 'Used'

# class Main:
#     gruss_receiver = "Herr Muller"
#     Receiver = GenderDetector(gruss_receiver)
#     Receiver.detect_gender()


