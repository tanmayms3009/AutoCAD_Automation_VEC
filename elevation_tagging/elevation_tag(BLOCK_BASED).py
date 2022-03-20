#BLOCK BASED ELEVATION TAGGING
from pyautocad import *
import win32com.client

acad = Autocad(create_if_not_exists=True)

def get_blocks():
    blocks = {}
    for i in acad.iter_objects_fast(object_name_or_list='Block'):
        blocks[i.Name] = i
    print(blocks)
    return blocks

def get_tag_attributes():
    element_attrs = {}
    for i in acad.iter_objects_fast(object_name_or_list='Block'):
        if i.Name not in element_attrs.keys():  
            element_attrs[i.Name] = []
            dwg_reference= input(str("Enter the shop/arcitectural drawing reference ID for CW/Window " + str(i.Name) +" : "))
            if dwg_reference != "None" or dwg_reference != "none" or dwg_reference != "":
                element_attrs[i.Name].append(dwg_reference)
            else:
                continue
    print(element_attrs)
    return element_attrs

def get_attributes():
    acad_com = win32com.client.Dispatch("AutoCAD.Application")
    # iterate through all objects (entities) in the currently opened drawing
    # and if its a BlockReference, display its attributes.
    for entity in acad_com.ActiveDocument.ModelSpace:
        name = entity.EntityName
        if name == 'AcDbBlockReference':
            HasAttributes = entity.HasAttributes
            if HasAttributes:
                for attrib in entity.GetAttributes():
                    print("  {}: {}".format(attrib.TagString, attrib.TextString))

def set_attributes(attr_dict: dict, ip_dict : dict):
    acad_com = win32com.client.Dispatch("AutoCAD.Application")
    # iterate through all objects (entities) in the currently opened drawing
    # and if its a BlockReference, display its attributes.
    for entity in acad_com.ActiveDocument.ModelSpace:
        name = entity.EntityName
        if name == 'AcDbBlockReference':
            HasAttributes = entity.HasAttributes
            if HasAttributes and entity.InsertionPoint[0] in ip_dict.keys():
                for attrib in entity.GetAttributes():
                    if attrib.TagString == 'DWG_NO':
                        attrib.TextString = attr_dict[ip_dict[entity.InsertionPoint[0]]][0]
                    elif attrib.TagString == 'ELEMENT_NAME':
                        attrib.TextString = ip_dict[entity.InsertionPoint[0]]                    
                    print("  {}: {}".format(attrib.TagString, attrib.TextString))

def insert_tags(size_dict: dict, blocks_dict: dict, attr_dict: dict):
    ip_to_block = {}
    for i in acad.iter_objects_fast(object_name_or_list='BlockReference'):
        try:
            print(i.Name)
            x = i.InsertionPoint[0] + size_dict[i.Name][0]/2
            y = i.InsertionPoint[1] + size_dict[i.Name][1]/2
            ip = [x, y, 0]
            if ip[0] not in ip_to_block.keys():
                ip_to_block[ip[0]] = i.Name
            else:
                k1 = ip[0] + 1
                ip_to_block[k1] = i.Name

            print(acad.model.InsertBlock(APoint(ip), blocks_dict['tags'].Name, 1, 1, 1, 0))
        except Exception as e:
            print('Error: ' + str(e))
    print(ip_to_block)
    return ip_to_block

window_blocks = get_blocks()
element_size = {'A1':[100, 200],'B1':[150, 200],'C2':[200, 200]}
windows_attr = get_tag_attributes()
tag_positions = insert_tags(element_size, window_blocks, windows_attr)
set_attributes(windows_attr, tag_positions)