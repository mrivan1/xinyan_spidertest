import requests
from common.rand import ranreq,tradedate
from common.Rsa.EncryptData import getEncryptData
import json
import xlrd
import urllib3
import xlsxwriter
from common.rand import tradedate

# 新建结果集
now = tradedate()

path = 'D:\DevCode\data\spider\\radar.xlsx'
data = xlrd.open_workbook(path)
data_sheet1 = data.sheets()[1]
data_sheet = data.sheets()[0]

filename = '测试报告'+now+'.xlsx'
workbook = xlsxwriter.Workbook('D:\DevCode\data\spider\\report\\radar\\'+filename)
work_sheet = workbook.add_worksheet('测试结果')


uri = data_sheet1.row_values(0)[1]
ENV = data_sheet1.row_values(0)[2]

if ENV == 'TEST':
    url = data_sheet.row_values(1)[1]
    member_id = data_sheet.row_values(1)[2]
    terminal_id = data_sheet.row_values(1)[3]
    pfx = data_sheet.row_values(1)[3]
    pass_word = data_sheet.row_values(1)[3]
else:
    url = data_sheet.row_values(2)[1]
    member_id = data_sheet.row_values(2)[2]
    terminal_id = data_sheet.row_values(2)[3]
    pfx = data_sheet.row_values(2)[3]
    pass_word = data_sheet.row_values(2)[3]

for m in range(data_sheet1.nrows - 2):

    rows = data_sheet1.row_values(m + 2)
    ID = rows[1]
    name = rows[0]
    bk_no = rows[3]
    mobile = str(rows[2]).replace('.0','')

    content = data_sheet1.row_values(0)[0]
    content = json.dumps(content)
    content = json.loads(content)
    print(type(content))
    #创建订单
    print("开始请求...")

    content_rsa = getEncryptData(pfx, pass_word, content).replace('"', '')
    paramas = {"member_id": member_id, "terminal_id": terminal_id,"data_type":"json", "data_content": content_rsa}
    urllib3.disable_warnings()
    print("请求参数为：" + str(content))
    response = requests.post(url+uri,data = paramas ,verify = False)
    print(response.json())
    work_sheet.write(m, 0, name)
    work_sheet.write(m, 1, ID)
    work_sheet.write(m, 2, mobile)
    work_sheet.write(m,3, str(response.text))

    print("开始下一条请求，进度为"+str(m+2))

workbook.close()
print(u'结果文件生成成功')


