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
        part_name.make_a_word_order()
        if part_name:
            name_order_list.append(part_name)
        k = k+1
    return name_order_list

def name_changer(wb,path):
    """
        This is a main function which renames files according to the parts list.
    """
    new_files = make_part_names_list(wb.active)
    old_files = os.listdir(path)

    for old_file in old_files:
        for new_file in new_files:
            try:
                position = "Pos" + old_file[:4]
                if position in new_file.name:
                    new_file = new_file.add_extension_from(old_file)
                    PartName.extension_counter(new_file)
                    os.rename(path + old_file, path + new_file)
            except FileNotFoundError:
                print(f"{old_file} - File Not Found")
            except IndexError:
                print("Index Error")