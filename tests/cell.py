
#simulation of excel cell in workbook sheet
class Cell():

    def __init__(self, value):
        self.value = value

    def __repr__(self):
        return '{}'.format(self.value)