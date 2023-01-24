import os

from .partname import PartName


def make_part_names_list(sheet):
    """ 
        This function creates a namelist of parts in according to XYZ standard.
        List includes correct name order which you can use to AGD files.
    """
    k = 9
    name_order_list = []
    
    while sheet['B{}'.format(k)].value is not None:
        part_name = PartName(sheet, k)
        part_name = part_name.make_a_word_order()
        if part_name:
            name_order_list.append(part_name)
        k = k+1
    return name_order_list

def name_changer(wb,path):
    """
        This is a main function which renames files according to the parts list.
    """
    names = make_part_names_list(wb.active)
    files = os.listdir(path)

    for file_name in files:
        for new_name in names:
            try:
                position = "Pos" + file_name[:4]
                if position in new_name:
                    new_name.add_extension(file_name)
                    PartName.extension_counter(file_name)
                    os.rename(path + file_name, path + new_name)
            except FileNotFoundError:
                pass
            except IndexError:
                pass