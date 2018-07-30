import requests
from common.rand import ranreq,tradedate
from common.Rsa.EncryptData import getEncryptData
import json
import xlrd
import urllib3
import xlsxwriter

path = 'D:\DevCode\data\spider\HJ.xlsx'
data = xlrd.open_workbook(path)
data_sheet = data.sheets()[0]

filename = '黑镜测试报告.xlsx'
workbook = xlsxwriter.Workbook('D:\DevCode\data\spider\\report\\radar\\'+filename)
work_sheet = workbook.add_worksheet('拉黑结果')

for m in range(data_sheet.nrows - 1):
    try:
        rows = data_sheet.row_values(m + 1)
        ID = rows[1]
        name = rows[0]
        mobile = str(rows[2]).replace('.0','')

        #创建订单
        print("开始请求...")
        content = {
            "member_id": "1022743",
            "terminal_id": "32635",
            "trade_date": tradedate(),
            "trans_id": ranreq(),
            "industry_type": "A1",
            "phone_no": mobile,
            "versions": "1.3.0",
            "id_no": ID,
            "id_name": name
        }
        content_rsa = getEncryptData('bfkey_1022743@@32635.pfx', '123456', content).replace('"', '')
        paramas = {"member_id": "1022743", "terminal_id": "32635","data_type":"json", "data_content": content_rsa}
        urllib3.disable_warnings()
        print("请求参数为：" + str(content))
        response = requests.post('https://api.xinyan.com/product/wash/simple/black',data = paramas ,verify = False)
        print(response.json())
        work_sheet.write(m, 0, name)
        work_sheet.write(m, 1, ID)
        work_sheet.write(m, 2, mobile)
        work_sheet.write(m,3, str(response.text))
    except:
        continue
    print("开始下一条请求，进度为"+str(m+1))

workbook.close()
print(u'结果文件生成成功')


