class PartName():

    jpg = 0
    png = 0
    stp = 0
    dxf = 0

    def __init__(self, sheet, row_number):
        self.name = ''
        self.order = sheet['G1'].value
        self.revision = sheet['A{}'.format(row_number)].value
        self.pos = sheet['B{}'.format(row_number)].value
        self.nor = sheet['D{}'.format(row_number)].value
        self.mirror = sheet['E{}'.format(row_number)].value
        self.material = sheet['H{}'.format(row_number)].value
        if self.material:
            self.material = self.material.replace('.', '_')
        self.thickness = sheet['L{}'.format(row_number)].value
        if self.thickness:
                if ',' in self.thickness:
                    self.thickness = self.thickness.replace(',', '_')
                elif '.' in self.thickness:
                    self.thickness = self.thickness.replace('.', '_')
        self.material_type = sheet['C{}'.format(row_number)].value

    def make_a_word_order(self):
        """
        This method joining words in correct order and making a name of the file
        """
        material_types = ('A','KN', 'G')
        if self.material_type in material_types:
            self.name = 'A' + str(self.order) + '_' + str(self.revision) + '_Pos' + str(self.pos) + '_' + str(self.nor) + 'xNorm'
            if self.mirror:
                self.name = self.name + '_' + str(self.mirror) + 'xGes'
            if self.material:
                self.name = self.name + '_Mat_' + str(self.material)
            if self.thickness:
                self.name = self.name + '_t=' + str(self.thickness[2:]) + 'mm'
            return self.name
    
    def add_extension_from(self,file_name):
        """
        This method adds extension to the new partname
        """
        extensions_list = [".jpg", ".png", ".stp", ".dxf"]
        file_name = file_name.lower()
        for item in extensions_list:
            if file_name.endswith(item):
                result = self.name + item
                return result
        
    @classmethod
    def extension_counter(cls, file_name):
        """
        This method counts the extensions of files which names were changed
        """
        if file_name.endswith(".jpg"):
            cls.jpg += 1
        elif file_name.endswith(".png"):
            cls.png += 1
        elif file_name.endswith(".stp"):
            cls.stp += 1
        elif file_name.endswith(".dxf"):
            cls.dxf += 1