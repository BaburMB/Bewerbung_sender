class SpaceRemover:

    def __init__(self, text_input):
        self.text = text_input

    def remove(self):
        self.text = self.text.lstrip()
        self.text = self.text.rstrip()
        return self.text

