class GenderDetector:
    gender = "d"

    def __init__(self, g_receiver):
        self.gruss_receiver = g_receiver

    def detect_gender(self):
        global gender
        if self.gruss_receiver[0] == "F" or self.gruss_receiver[0] == "f":
            return "Female"
        elif self.gruss_receiver[0] == "H" or self.gruss_receiver[0] == "h":
            return "Male"
        else:
            print("Couldn't detect gender for ", self.gruss_receiver[0], ", will use default gender for Sehr...")
        print("Receiver's gender is ", self.gender)
        

