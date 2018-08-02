from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.styles.borders import Border, Side


# def find_all_classes(ws):
#     find_column = False
#     while(!find_column):
        
def find_all_classes(ws):
    class_row = None
    for row in ws.iter_rows(min_row=1, max_col=1, max_row=10):
        for cell in row:
            if(cell.value != None and cell.value.lower() == "day"):
                class_row = cell.row
    if class_row is None:
        return None
    classes = []
    for column in ws.iter_cols(min_row=class_row ,max_col=100,max_row=class_row):
        for cell in column:
            if(cell.value!=None and (cell.value.lower()!="day" or cell.value.lower() != "hour")):
                classes.append(cell.value)
    return classes

def find_class(ws,class_code):
    class_row = None
    for row in ws.iter_rows(min_row=1, max_col=1, max_row=10):
        for cell in row:
            if(cell.value != None and cell.value.lower() == "day"):
                class_row = cell.row
    if class_row is None:
        return None
    for column in ws.iter_cols(min_row=class_row ,max_col=100,max_row=class_row):
        for cell in column:
            if(cell.value!=None and cell.value.lower() == class_code.lower()):
                return cell
    return None

def get_timetable(ws,name_cell):
    to_skip = 0
    start_cell = name_cell.offset(1,0)
    for x in range(5):
        for row in ws.iter_rows(min_row=start_cell.row,max_row=(int(start_cell.row)+2*10),max_col=column_index_from_string(start_cell.column),min_col=column_index_from_string(start_cell.column)):
            for cell in row:
                if(to_skip<=0):
                    class_code,class_room,teacher_code,to_skip = get_period(cell)
                    print(class_code)
                else:
                    to_skip-=1
                end_cell = cell
        print("END OF DAY")
        start_cell = end_cell.offset(1,0)

def get_period(cell):
    #to find the code
    class_code = cell.value
    cell_under=cell.offset(1,0)
    class_cell = cell
    if(class_code == None and cell_under.border.left.style == None):
        while(class_code==None and class_cell.border.left.style != "medium"):
            class_cell = class_cell.offset(0,-1)
            class_code = class_cell.value

    #find cells to skip
    to_skip = 0
    current_cell = class_cell
    while(current_cell.border.bottom.style==None):
        current_cell = current_cell.offset(1,0)
        to_skip+=1
    if to_skip>3:
        to_skip=3

    #find class room
    class_room = class_cell.offset(1,0).value

    #find teacher code 
    if(to_skip==3):
        teacher_code = class_cell.offset(3,0).value
    elif(class_code!=None):
        teacher_cell = class_cell.offset(1,1)
        while(teacher_cell.value==None):
            teacher_cell = teacher_cell.offset(0,1)
        teacher_code = teacher_cell.value
    else:
        teacher_code = None

    return class_code,class_room,teacher_code,to_skip
