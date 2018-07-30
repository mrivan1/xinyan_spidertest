import readconfig
import xlrd
from common.excel.Obtain_excel import getColumnIndex,readExcelDataByIndex


localRead = readconfig.ReadConfig()
path = localRead.get_address('CASE')

# 获取案例所在的table,以及各项所在的列
table_n = readExcelDataByIndex(path, 1)[0]
ISRUN = getColumnIndex(table_n,'ISRUN')
TYPE = getColumnIndex(table_n,'TYPE')

data = xlrd.open_workbook(path)
data_sheet = data.sheets()[1]

for rown in range(data_sheet.nrows-1):
    rows = data_sheet.row_values(rown + 1)
    ISRUN_v = int(rows[ISRUN])
    TYPE_v = rows[TYPE]
    if ISRUN_v == 1:
        if TYPE_v == '公积金':
            from Auto.house_fund import house_fund
            house_fund()
        elif TYPE_v == '社保':
            from Auto.Social import social
            social()
    elif ISRUN_v == 0:
        continue


