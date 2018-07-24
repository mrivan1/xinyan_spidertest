import xlrd

#获取表的某一列所在的脚标
def getColumnIndex(table,columnName):
    columnIndex = None
    for i in range(table.ncols):
        if (table.cell_value(0, i) == columnName):
            columnIndex = i
            break

    return columnIndex

#根据sheet名获取sheet内容
def readExcelDataByName(fileName, sheetName):
    table = None
    errorMsg = ""
    try:
        data = xlrd.open_workbook(fileName)
        table = data.sheet_by_name(sheetName)
    except Exception as msg:
        errorMsg = msg

    return table, errorMsg

#根据sheet顺序获取sheet内容
def readExcelDataByIndex(fileName, sheetIndex):
    table = None
    errorMsg = ""
    try:
        data = xlrd.open_workbook(fileName)
        table = data.sheet_by_index(sheetIndex)
    except Exception as msg:
        errorMsg = msg
    return table, errorMsg




#获取

