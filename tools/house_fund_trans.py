import json
import xlsxwriter

workbook = xlsxwriter.Workbook('D:/DevCode/data/radar/gongjj.xlsx')
work_sheet = workbook.add_worksheet('公积金城市信息')


with open('D:/DevCode/data/radar/gjj.txt',encoding= 'utf-8') as gjj:
    txt = gjj.read()
    txt = json.loads(txt)
    for i in range(len(txt)):
        city = txt[i]["city"]
        web_url = txt[i]["web_url"]
        work_sheet.write(i,0,city)
        work_sheet.write(i,1,web_url)

workbook.close()