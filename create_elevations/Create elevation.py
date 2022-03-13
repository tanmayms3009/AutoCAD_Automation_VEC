from pyautocad import Autocad, APoint, aDouble

acad = Autocad(create_if_not_exists=True)


def list_of_layers_available():
    all_layers=[]
    for i in acad.doc.Layers:
        if i.Name in all_layers:
            continue
        else:
            all_layers.append(i.Name)
    
    print("List of all layers present in the document: ", end="")
    print(all_layers)
    return all_layers        

def list_used_layers():
    layers = []
    for i in acad.iter_objects_fast(object_name_or_list=None, limit=None):
        if i.Layer in layers:
            continue
        else:
            layers.append(i.Layer)
            print(i.Layer)
    print("Layers used in drawings: ", end="")
    print(layers)        

#Create a layer
# acad.doc.Layers.Add('tanmay')


print('Active layer: ' + acad.doc.ActiveLayer.Name)

#General properties of layer>TrueColor object:
def gen_color_prop():
    for i in acad.iter_objects_fast(object_name_or_list=None, limit=None):
        if i.Layer == 'A-GLAZ':
            print(i.TrueColor.ColorName)
            print(i.TrueColor.ColorIndex)
            print(i.TrueColor.ColorMethod)        
            print(i.TrueColor.Red)
            print(i.TrueColor.Green)
            print(i.TrueColor.Blue)
            print(i.TrueColor.BookName)
            print(i.TrueColor.EntityColor)

lay1 = acad.doc.ActiveLayer
lay1.TrueColor.ColorMethod
print(lay1.TrueColor.ColorMethod)


#Filter out other materials
# for i in acad.iter_objects_fast(object_name_or_list=None, limit=None):
#         if i.Layer not in ['A-WALL', 'A-WALL-PATT', 'A-DETL', 'A-GLAZ', 'A-DETL-GENF']:
#             i.Move(APoint(0,0,0), APoint(0,-2000,0))

# y = 1000
# for i in acad.iter_objects_fast(object_name_or_list='BlockReference', limit=None):
#     i.Move(APoint(0, 0, 0), APoint(0, y, 0))
#     y+=1000

for i in acad.doc.SelectionSets:
    print(i.Name)
mode= 'acSelectionSetWindow'
po1= aDouble(0, 0, 0)
po2= aDouble(4000,2000,0)
po3= APoint(0, 0, 0)
po4= APoint(4000,2000,0)
# for i in acad.iter_objects_fast(object_name_or_list=None, limit=None):
    # s1.Select(mode, po1, po2)
    



# for i in acad.iter_objects_fast():
#     if i.Layer=="PDF_TEXT":
#         i.Move(APoint(0,0,0), APoint(0,100,0))


#1. List out all layers
#2. Ask user input for filtering out the unwanted layers
#3. Move the objects with specified layers to a specific location
#4. Select the objects to be put in screening layer
#5. Create a block of screened object
