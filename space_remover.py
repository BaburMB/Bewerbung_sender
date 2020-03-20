class SpaceRemover:

    def __init__(self, text_input):
        self.text = text_input

    def remove(self):
        self.text = self.text.lstrip()
        self.text = self.text.rstrip()
        return self.text

        # if self.text[0] == " ":
        #     self.text = self.text.lstrip()
        #     return self.text
        #     #print(self.text)
        #     #return SpaceRemover(self.text).remove()
        # elif self.text[len(self.text)-1] == " ":
        #     self.text = self.text[len(self.text)-1].lstrip()
        #     return self.text
            #print(self.text)
            #return SpaceRemover(self.text).remove()
        # else:
        #     print(self.text)
        #     pass

#
# class Main:
#
#     a = "            any text "
#     a = SpaceRemover(a).remove()
#     print(a)
#     print("s" + a + "s")
    # print("")
