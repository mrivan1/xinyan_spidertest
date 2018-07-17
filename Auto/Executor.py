import requests
from common.rand import membertrans,tradedate
from common.Rsa.EncryptData import getEncryptData
import json
import xlrd
import urllib3
import time
import xlsxwriter

path = 'D:\DevCode\data\spider\Excutor1.xlsx'
data = xlrd.open_workbook(path)
data_sheet = data.sheets()[0]

filename = '苏宁测试报告.xlsx'
workbook = xlsxwriter.Workbook('D:\DevCode\data\spider\\'+filename)
work_sheet = workbook.add_worksheet('法院被执行人')

result = []
for m in range(data_sheet.nrows - 1):
    try:
        rows = data_sheet.row_values(m + 1)
        ID = rows[0]
        mobile = rows[1]
        name = rows[2]

        #创建订单
        print("创建订单中...")
        content = {
            'member_id':'1107602',
            'terminal_id':'32912',
            'member_trans_date':tradedate(),
            'member_trans_id':membertrans(),
            'notify_url':'',
            'user_id':'苏宁',
            'id_card':ID,
            'user_name':name
        }
        content_rsa = getEncryptData('1107602@32912_pri.pfx', 'xy5434', content).replace('"', '')
        data = {'member_id': '1107602', 'terminal_id': '32912', 'data_content': content_rsa}
        headers = {"content-type": "application/json"}
        urllib3.disable_warnings()
        response = json.loads(requests.post('https://api.xinyan.com/gateway-data/zhixing/v1/task/create',data = json.dumps(data),headers = headers,verify = False).text)
        print("订单创建成功："+str(response))
        result.append({'tradeno':response["data"]["tradeNo"],'id_card':ID,'name':name,'mobile':mobile})
    except:
        continue

#订单状态查询

print("获取订单状态中.....")
num = 0
for trade in result:
    try:
        header = {
            'memberId':'1107602',
            'terminalId':'32912'
        }
        status = json.loads(requests.get('https://api.xinyan.com/gateway-data/zhixing/v1/task/status/'+str(trade["tradeno"]),headers = header,verify = False).text)
        if status["errorMsg"] == None:
            errorMsgs = ''
            out = 0
            while status["data"]["phase"] != 'DONE':
                status = json.loads(requests.get('https://api.xinyan.com/gateway-data/zhixing/v1/task/status/' + str(trade["tradeno"]),headers=header, verify=False).text)
                print('订单状态为：'+str(status["data"]["description"]))
                if status["errorCode"] != None:
                    errorMsgs = str(status['errorMsg'])
                    out = 1
                    break
                elif status["data"]["phase_status"] == 'DONE_FAIL':
                    errorMsgs = str(status["data"]["description"])
                    out = 1
                    break
        else:
            errorMsgs = str(status['errorMsg'])
            out = 1
        print("订单状态为："+str(status["data"]["description"]))
        if out == 0:
            # 根据订单号查询法院被执行人信息
            print("查询订单结果中....")
            headers = {
            'memberId':'1107602',
            'terminalId':'32912'
            }
            res = requests.get('https://api.xinyan.com/data/zhixing/v2/info/'+str(trade["tradeno"]),headers = headers,verify=False).text
            print('结果为：'+str(res))
            work_sheet.write(num, 0, trade["name"])
            work_sheet.write(num, 1, trade["id_card"])
            work_sheet.write(num, 2, trade["mobile"])
            work_sheet.write(num, 3, res)
        else:
            work_sheet.write(num, 0, trade["name"])
            work_sheet.write(num, 1, trade["id_card"])
            work_sheet.write(num, 2, trade["mobile"])
            work_sheet.write(num, 3, errorMsgs)
        num = num + 1
        print('开始查询下一条数据，数据标志为：'+str(num))
    except:
        continue

workbook.close()
print(u'结果文件生成成功')


