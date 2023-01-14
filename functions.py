import os


class PartName():

    def __init__(self, sheet, row_number):
        self.sheet = sheet
        self.row_number = row_number
        self.order = self.sheet['G1'].value
        self.revision = self.sheet[f'A{self.row_number}'].value
        self.pos = self.sheet[f'B{self.row_number}'].value
        self.nor = self.sheet[f'D{self.row_number}'].value
        self.mirror = self.sheet[f'E{self.row_number}'].value
        self.material = self.sheet[f'H{self.row_number}'].value
        self.thickness = self.sheet[f'L{self.row_number}'].value
        self.material_type = self.sheet[f'C{self.row_number}'].value
        if self.material:
            self.material = self.material.replace('.', '_')
        if self.thickness:
            if ',' in self.thickness:
                self.thickness = self.thickness.replace(',', '_')
            elif '.' in self.thickness:
                self.thickness = self.thickness.replace('.', '_')

    def make_a_word_order(self):
        """
        This method joining words in correct order and making a name of the file
        """
        if self.material_type == 'A' or self.material_type == 'KN' or self.material_type == 'G':
            self.name = 'A' + str(self.order) + '_' + str(self.revision) + '_Pos' + str(self.pos) + '_' + str(self.nor) + 'xNorm'
            if self.mirror:
                self.name = self.name + '_' + str(self.mirror) + 'xGes'
            if self.material:
                self.name = self.name + '_Mat_' + str(self.material)
            if self.thickness:
                self.name = self.name + '_t=' + str(self.thickness[2:]) + 'mm'
            return self.name


def make_part_names_list(sheet):
    """ 
        This function create a namelist of parts in according to XYZ standard.
        List includes correct name order which you can use to AGD files.
    """
    k = 9
    name_order_list = []
    
    while sheet[f'B{k}'].value is not None:
        part_name = PartName(sheet, k)
        part_name = part_name.make_a_word_order()
        if part_name:
            name_order_list.append(part_name)
        k = k+1
    return name_order_list


def main(wb,path):
    """
        This is a main function which renames files according to the parts list and returns how many file names were changed (extension type included).
    """
    names = make_part_names_list(wb.active)
    files = os.listdir(path)
    amount_of_files_type = {'.jpg':0,'.png':0,'.stp':0,'.dxf':0}

    for file_name in files:
        for name in names:
            try:
                file_name = file_name.lower()
                position = "Pos" + file_name[:4]
                if position in name:
                    if file_name.endswith((".jpg",".png",".stp",".dxf")):
                        amount_of_files_type[file_name[-4:]] += 1
                        extension = file_name[-4:]
                        name = name + extension
                    os.rename(path + file_name, path + name)
            except FileNotFoundError:
                pass
            except IndexError:
                pass
    return amount_of_files_type