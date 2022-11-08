import os


def words_linking(order, revision, pos, nor, mirror, material, material_type, thickness):

    """
        This function joining words in correct order and making a name of the file
    """

    if material_type == 'A' or material_type == 'KN' or material_type == 'G':
        name = 'A' + str(order) + '_' + str(revision) + '_Pos' + str(pos) + '_' + str(nor) + 'xNorm'

        if mirror:
            name = name + '_' + str(mirror) + 'xGes'

        if material:
            name = name + '_Mat_' + str(material)

        if thickness:
            name = name + '_t=' + str(thickness[2:]) + 'mm'

        return name


def agd_namelist_func(sheet):

    """ 
        This function create a namelist of parts in according to XYZ standard.
        List includes correct name order which you can use to AGD files.
    """
    
    k = 9
    my_list = []

    while sheet[f'B{k}'].value is not None:

        order = sheet['G1'].value
        revision = sheet[f'A{k}'].value
        pos = sheet[f'B{k}'].value
        nor = sheet[f'D{k}'].value
        mirror = sheet[f'E{k}'].value
        material = sheet[f'H{k}'].value
        thickness = sheet[f'L{k}'].value
        material_type = sheet[f'C{k}'].value

        if material:
            material = material.replace('.', '_')

        if thickness:

            if ',' in thickness:
                thickness = thickness.replace(',', '_')
            elif '.' in thickness:
                thickness = thickness.replace('.', '_')

    # words linking

        part_name = words_linking(order, revision, pos, nor, mirror, material, material_type, thickness)
        
        if part_name:
            my_list.append(part_name)

        k = k+1

    return my_list

def main(wb,path):

    """
        This is a main function which changing files name according to the part list.
    """

    names = agd_namelist_func(wb.active)
    files = os.listdir(path)

    image = 0
    stp = 0
    dxf = 0

    for file in files:
        for name in names:
            try:
                position = 'Pos' + file[:4]
                if position in name:
                    if file.endswith('.jpg') or file.endswith('.JPG'):
                        new_name = name + '.jpg'
                        os.rename(path + file, path + new_name)
                        image+=1
                    elif file.endswith('.stp') or file.endswith('.STP'):
                        new_name = name + '.stp'
                        os.rename(path + file, path + new_name)
                        stp+=1
                    elif file.endswith('.dxf') or file.endswith('.DXF'):
                        new_name = name + '.dxf'
                        os.rename(path + file, path + new_name)
                        dxf+=1
                    elif file.endswith('.png') or file.endswith('.PNG'):
                        new_name = name + '.png'
                        os.rename(path + file, path + new_name)
                        image+=1
            except FileNotFoundError:
                pass
            except IndexError:
                pass
            
    return image,stp,dxf
