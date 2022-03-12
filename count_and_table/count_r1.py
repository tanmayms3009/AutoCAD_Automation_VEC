#Importing pyautocad library:
from pyautocad import Autocad, APoint, aDouble
from openpyxl import Workbook, load_workbook

#---------

def merge_sort(unsorted_list:list):
   if len(unsorted_list) <= 1:
      return unsorted_list
# Find the middle point and devide it
   middle = len(unsorted_list) // 2
   left_list = unsorted_list[:middle]
   right_list = unsorted_list[middle:]

   left_list = merge_sort(left_list)
   right_list = merge_sort(right_list)
   return list(merge(left_list, right_list))
# Merge the sorted halves
def merge(left_half,right_half):
   res = []
   while len(left_half) != 0 and len(right_half) != 0:
      if left_half[0] < right_half[0]:
         res.append(left_half[0])
         left_half.remove(left_half[0])
      else:
         res.append(right_half[0])
         right_half.remove(right_half[0])
   if len(left_half) == 0:
      res = res + right_half
   else:
      res = res + left_half
   return res

#---------

wb = Workbook()
ws = wb.active
ws.title = "Details"

#---------

headings = ['Elevation', 'Window Type', 'Levels', 'Quantity']
ws.append(headings)

#---------

#Setting "create_if_not_exists"  to "True" will open & create AutoCAD template if not already:
acad = Autocad(create_if_not_exists=True)

#---------

#Basic object dictionary
no_objects = int(input("Enter number of objects to be found: "))
objects = []
for i in range(no_objects):
    obj = input("Enter name of object number " + str(i+1) + " :")
    objects.append(obj)
print(objects)

#---------

#ElevationDictionary = {0: e, 1: w, 2: n1, 3: s1}
no_elevs = int(input("Enter number of elevations: "))
elevs = []
for j in range(no_elevs):
    elev = input("Enter name of elevation number " + str(j+1) + " :")
    elevs.append(elev)
print(elevs)
spacing = int(input("Enter the width of plot w.r.t X-axis: "))
elev_spac = {}
sp = 0
ep = spacing
for jj in range(no_elevs):
    elev_spac[elevs[jj]] = (sp, ep)
    sp+=spacing
    ep+=spacing
print(elev_spac)

#---------

#Layer in which object to be searched is present
layer_of_obj = str(input("Enter layer of object: "))


#ObjectDictionry = {A1:{1: (x, y, z)}, B1:{3: (x, y, z)}}
det_objects = {}
for k in range(len(objects)):
    count_a = 0
    key =  0
    det_objects[objects[k]] = []
    for l in acad.iter_objects_fast(object_name_or_list="Text"):
        if l.TextString == objects[k] and l.Layer == layer_of_obj:
            print(objects[k] + " " + l.Layer, end=": ")
            print(l.InsertionPoint)
            det_objects[objects[k]].append(l.InsertionPoint)
            count_a+=1
    print("Number of " + objects[k] + ": " + str(count_a))
print(det_objects)

#---------

#Get key of specific value from a dictionary
def get_key(val):
     for key, value in elev_spac.items():
        if val == value:
            return key
     return "Key doesnt exist"

#---------

#FIND LEVELS
lev_lay = str(input("Enter name of layer used to specify level line: "))
def extract_levels(lev_lay_input: str, start_point: int, end_point: int):
    for i in acad.iter_objects_fast(object_name_or_list="Text"):
        #" ' " CONDITION MIGHT CREATE PROBLEM
        #CREATE LIST OF LEVEL NOMENCLATURE
        if i.Layer == lev_lay and "'" in i.TextString and i.InsertionPoint[0] > sp and i.InsertionPoint[0] <= ep:
            print(i.TextString)
    levels_y = []
    for i in acad.iter_objects_fast(object_name_or_list=["AcDbLine", "AcDbPolyLine"]):
        if i.Layer == lev_lay and i.StartPoint[0] > sp and i.EndPoint[0] <= ep:
            levels_y.append(i.StartPoint[1])
    # print(merge_sort(levels_y))
    return levels_y


#---------

#GET LEVEL OF SELECTED OBJECT
def find_level(y_coord_tag: int, level_points: list):
    levl = 1
    ii= 0
    jj = ii+1
    while jj < len(level_points):
        if y_coord_tag > level_points[ii] and y_coord_tag <= level_points[jj]:
            # print(level_points[i], level_points[j])
            return "Level " + str(ii)
        levl+=1
        ii+=1
        jj+=1
# print(find_level(2482.03, levels_present))

#---------

#Seggregation on the basis of elevation name
#Iterate through every object of window detail list and check if X is within a certain range and then fetch the value from elevations
#Iterate through every type of window
levels_coord = {}
for x in det_objects:
    sp = 0
    ep = spacing
    #Iterate through every elevation
    for y in range(len(elevs)):
        print(levels_coord)
        no_on_elev = 0
        level_key = get_key((sp, ep))
        if level_key in levels_coord:
            print("List of levels present in elevation " + str(get_key((sp, ep))) + ": ", end="")
            print(levels_coord[level_key])
        else:
            levels_present = merge_sort(extract_levels(lev_lay, sp, ep))
            levels_coord[level_key] = levels_present
            print("List of levels present in elevation " + str(get_key((sp, ep))) + ": ", end="")
            print(levels_present)

        #Iterate through Insertion points of window name
        for i in range(len(det_objects[x])):
            if det_objects[x][i][0] > sp and det_objects[x][i][0] <= ep:
                det_w = [get_key((sp, ep)), x, find_level(int(det_objects[x][i][1]), levels_present), 1]
                # To print coordinates: print(det_objects[x][i][0], det_objects[x][i][1], det_objects[x][i][2]) 
                ws.append(det_w)
                no_on_elev+=1
        print("Number of " + x + " type window on " + get_key((sp, ep)) + " elevation: " + str(no_on_elev))
        sp+=spacing
        ep+=spacing
print(levels_coord)
#---------

wb.save("Windows.xlsx")

#============================GENETRATE PIVOT IN DIFFERENT SHEET===============================

# USE PIVOT.PY 

#============================CREATE SORTED LIST FOR LEVEL NOMENCLATURE======================== 