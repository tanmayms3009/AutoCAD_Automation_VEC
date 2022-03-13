from pyautocad import Autocad, APoint, aDouble

acad = Autocad(create_if_not_exists=True)


for j in acad.iter_objects_fast(object_name_or_list=['BlockReference']):
        if j.Name=="*U7":
                print(j.InsertionPoint) 
                h1 = j


for i in acad.iter_objects_fast(object_name_or_list=['BlockReference']):
        if i.Name != "*U7":
                print(i.InsertionPoint)
                print(i.Name)
                # print(i.XEffectiveScaleFactor)
                # print(i.YEffectiveScaleFactor)
                # print(i.ZEffectiveScaleFactor)
                # print(i.XScaleFactor)
                # print(i.YScaleFactor)
                # print(i.ZScaleFactor)
                print(i.PlotStyleName)
                print(i.Layer)
                h1_1 = h1.Copy()
                h1_1.Move(APoint(33.41334083425318, -45.187925824544436, 0.0), APoint(i.InsertionPoint[0]+10, i.InsertionPoint[1]-10, i.InsertionPoint[2],))
                # 'minExt', 'maxExt'


#1.Detect block-Done
#2.Check if block is getting considered as a whole consisting of all the sub elements-Done
#3.Check if block properties consisits of coordinates-No
#4.Use similar filtering technique as tag filtering
#5.Provide both elevation and level based solutions, separate & integrated