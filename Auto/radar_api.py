import json

import requests
import xlrd

from common import rand
from readconfig import ReadConfig

localread = ReadConfig()
path = localread.get_address("auto")
try:
    file = xlrd.open_workbook(path)
    sheet = file.sheet_by_name('InterfaceEnvironment')

    for i in range(1,sheet.nrows):
        DoEx = sheet.cell_value(i,0)
        if DoEx == 'Y':
            enve = sheet.cell_value(i,2)
            merchant_id = sheet.cell_value(i,3)
            terminal_id = sheet.cell_value(i,4)
            secret_key = sheet.cell_value(i,5)
            key_pfx = sheet.cell_value(i,6)
            key_cer = sheet.cell_value(i,7)
            key_password = sheet.cell_value(i,8)
        else:
            continue
except:
    print("请确认文件是否解密")

file1 = xlrd.open_workbook(path)
sheet1 = file.sheet_by_name("Test")

for j in range(1, sheet1.nrows):
    DoExt = sheet1.cell_value(j,1)
    if DoExt == 'Y':
        Sheetname = sheet1.cell_value(j,2)
        file2 = xlrd.open_workbook(path)
        sheet2 = file.sheet_by_name(Sheetname)
        datalist = json.loads(sheet2.cell_value(1,0),encoding="utf-8")
        datalist1 = json.loads(sheet2.cell_value(1,1),encoding='utf-8')
        for k in range(3,4):
            DoExt1 = sheet2.cell_value(k,0)
            if DoExt1 =='Y':
                datalist['member_id'] = merchant_id
                datalist['terminal_id'] = terminal_id
                datalist['trans_id'] = rand.ranreq()
                datalist['trade_date'] = rand.tradedate()
                datalist['id_info'] = rand.data_fc(sheet2.cell_value(k, 11))
                datalist['id_card'] = rand.data_fc(sheet2.cell_value(k, 9))
                datalist['id_no'] = rand.data_fc(sheet2.cell_value(k, 10))
                datalist['name'] = rand.data_fc(sheet2.cell_value(k, 12))
                datalist['id_name'] = rand.data_fc(sheet2.cell_value(k, 13))
                datalist['versions'] = rand.data_fc(sheet2.cell_value(k, 14))
                datalist['number_type'] = rand.data_fc(sheet2.cell_value(k, 15))
                datalist['phoneNo'] = rand.data_fc(sheet2.cell_value(k, 16))
                datalist['phone_no'] = rand.data_fc(sheet2.cell_value(k, 17))
                datalist['product_codes'] = rand.data_fc(sheet2.cell_value(k, 19))
                datalist['bankcard_no'] = rand.data_fc(sheet2.cell_value(k, 18))
                datalist1["member_id"] = merchant_id
                datalist1["terminal_id"] = terminal_id
                datalist1["data_type"] = rand.data_fc(sheet2.cell_value(k, 24))
                datalist1["data_content"] = str(datalist)
                data_final = datalist1
                url = enve + rand.data_fc(sheet2.cell_value(k, 4))
                r = requests.post(url,data_final)
                print(r.text)
            else:
                continue


