from openpyxl import load_workbook

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