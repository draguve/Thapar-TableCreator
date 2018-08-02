from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import column_index_from_string
from openpyxl.styles.borders import Border, Side
from openpyxl.compat import range as pyxlrange


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
    
    finalworkbook,finalsheet = create_empty_table()
    finalsheet.title = name_cell.value

    to_skip = 0
    start_cell = name_cell.offset(1,0)

    daycell = finalsheet["B2"]
    current_cell = daycell

    for x in range(5):
        for row in ws.iter_rows(min_row=start_cell.row,max_row=(int(start_cell.row)+2*11-1),max_col=column_index_from_string(start_cell.column),min_col=column_index_from_string(start_cell.column)):
            for cell in row:
                if(to_skip<=0):
                    #get the period data at the right location
                    class_code,class_room,teacher_code,to_skip = get_period(cell)
                    if(class_code!= None and class_code[-1] != "P" and to_skip == 3):
                        print(cell)
                    if(class_code!=None and class_room!=None and teacher_code!=None):
                        to_write = current_cell
                        to_write.value = class_code
                        if to_skip == 3:
                            to_write.offset(1,0).value = class_room
                            to_write.offset(2,0).value = "LAB"
                            to_write.offset(3,0).value = teacher_code
                        else:
                            to_write.offset(1,0).value = "%s | %s"%(class_room,teacher_code)
                    current_cell = current_cell.offset(to_skip+1,0) 
                else:
                    to_skip-=1
                end_cell = cell
        #End of day
        to_skip = 0
        start_cell = end_cell.offset(1,0)
        daycell = daycell.offset(0,1)
        current_cell = daycell

    return finalworkbook

def get_period(cell):
    #to find the code
    class_code = cell.value
    cell_under=cell.offset(1,0)
    class_cell = cell
    if(class_code == None and cell_under.border.left.style == None):
        counter = 2
        while(class_code==None and class_cell.border.left.style != "medium" and counter>0):
            class_cell = class_cell.offset(0,-1)
            class_code = class_cell.value
            counter -= 1
            if(class_code!= None and class_code[-1] == "T"):
                class_code = None
                break


    #find cells to skip
    if(class_code!=None and class_code[-1] == "P"):
        to_skip = 3
    else:
        to_skip = 1

    #find class room
    if(class_code!=None):
        class_room = class_cell.offset(1,0).value
    else:
        class_room = None

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

def create_empty_table():
    #creating a new workbook to store the new timetable
    wb = Workbook()
    finalsheet = wb.active

    #formatting the table
    finalsheet["A2"].value = "Time/Day"

    current_cell = finalsheet["A2"]
    time = 8
    for x in range(1,11):
        current_cell.value = str(time%12) + " To"
        current_cell.offset(1,0).value = str((time+1)%12)
        current_cell = current_cell.offset(2,0)
        time+=1

    
    #removing borders
    empty_border = Border(left=Side(border_style=None,
                           color='FF000000'),
                 right=Side(border_style=None,
                            color='FF000000'),
                 top=Side(border_style=None,
                          color='FF000000'),
                 bottom=Side(border_style=None,
                             color='FF000000'),
                 diagonal=Side(border_style=None,
                               color='FF000000'),
                 diagonal_direction=0,
                 outline=Side(border_style=None,
                              color='FF000000'),
                 vertical=Side(border_style=None,
                               color='FF000000'),
                 horizontal=Side(border_style=None,
                                color='FF000000')
                )

    for x in range(1,12):
        for y in range(1,12):
            finalsheet.cell(row=x,column=y).border = empty_border

    return wb,finalsheet   