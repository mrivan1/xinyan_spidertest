import requests
from common.Rsa.EncryptData import getEncryptData
import time
from common.excel.Obtain_excel import getColumnIndex,readExcelDataByIndex
from common.excel.Security_report import report,close_workbook,create_ownreport
from common.rand import membertrans,tradedate
import readconfig
import xlrd
import json
import urllib3

def social():
    localRead = readconfig.ReadConfig()
    path = localRead.get_address('CASE')



    #获取案例所在的table,以及各项所在的列
    table_n = readExcelDataByIndex(path,3)[0]

    ID = getColumnIndex(table_n,'ID')
    DESCRIBE = getColumnIndex(table_n,'DESCRIBE')
    CASE_ID = getColumnIndex(table_n,'CASE_ID')
    IS_RUN = getColumnIndex(table_n,'IS_RUN')
    ENV = getColumnIndex(table_n,'ENV')
    notify_url = getColumnIndex(table_n,'notify_url')
    user_id = getColumnIndex(table_n,'user_id')
    area_name = getColumnIndex(table_n,'area_name')
    account = getColumnIndex(table_n,'account')
    password = getColumnIndex(table_n,'password')
    login_type = getColumnIndex(table_n,'login_type')
    id_card = getColumnIndex(table_n,'id_card')
    mobile = getColumnIndex(table_n,'mobile')
    real_name = getColumnIndex(table_n,'real_name')
    sub_area = getColumnIndex(table_n,'sub_area')
    corp_account = getColumnIndex(table_n,'corp_account')
    corp_name = getColumnIndex(table_n,'corp_name')
    origin = getColumnIndex(table_n,'origin')
    ip = getColumnIndex(table_n,'ip')
    URL = getColumnIndex(table_n,'URL')

    #解析案例
    data = xlrd.open_workbook(path)
    data_sheet_env = data.sheets()[0]


    #记录生成的tradeno
    trade_NO = []
    rownum = 0
    rownum1 = 0
    data_sheet = data.sheets()[3]

    for rown in range(data_sheet.nrows-1):
        rows = data_sheet.row_values(rown + 1)
        out = 0
        #判断需要执行才进行以下操作
        if rows[IS_RUN] == 1:

            #获取环境变量信息
            if rows[ENV] == 'TEST':
                row_test = data_sheet_env.row_values(1)
                env = row_test[1]
                memberId_v = row_test[2]
                member_id_v = row_test[2]
                terminalId_v = row_test[3]
                terminal_id_v =row_test[3]
                key_pfx_v = row_test[4]
                key_password_v =row_test[5]
            elif rows[ENV] == 'PRO':
                row_test = data_sheet_env.row_values(2)
                env = row_test[1]
                memberId_v = row_test[2]
                member_id_v = row_test[2]
                terminalId_v = row_test[3]
                terminal_id_v =row_test[3]
                key_pfx_v = row_test[4]
                key_password_v =row_test[5]

            #获取接口参数信息
            ID_v = rows[ID]
            DESCRIBE_v = rows[DESCRIBE]
            CASE_ID_v = rows[CASE_ID]
            ENV_v = rows[ENV]
            notify_url_v = rows[notify_url]
            user_id_v = rows[user_id]
            area_name_v = rows[area_name]
            account_v = rows[account]
            password_v = rows[password]
            login_type_v = rows[login_type]
            id_card_v = rows[id_card]
            mobile_v = rows[mobile]
            real_name_v = rows[real_name]
            sub_area_v = rows[sub_area]
            corp_account_v = rows[corp_account]
            origin_v = rows[origin]
            ip_v = rows[ip]
            corp_name_v = rows[corp_name]
            url_v = rows[URL]

            #获取社保支持地区信息
            print("获取地区信息中...")
            header_s = {'memberId':memberId_v,'terminalId':terminalId_v}
            urllib3.disable_warnings()
            response = requests.get(env+"/gateway-data/security/v1/arealist",headers = header_s,verify = False)
            print("请求的地区为："+area_name_v)
            response_json = json.loads(response.text)

            for row in response_json["data"]:
                if area_name_v == row["city_name"]:
                    area_code = row["area_code"]
                    print("地区编码为："+area_code)
                else:
                    pass

            #创建社保订单
            content = {'member_id': member_id_v,
                       'terminal_id': terminal_id_v,
                       'member_trans_date': tradedate(),
                       'member_trans_id': membertrans(),
                       'notify_url': notify_url_v,
                       'user_id': user_id_v,
                       'area_code': area_code,
                       'account': account_v,
                       'password': password_v,
                       'login_type': login_type_v,
                       'id_card': id_card_v,
                       'mobile': mobile_v,
                       'real_name': real_name_v,
                       'sub_area': sub_area_v,
                       'corp_account': corp_account_v,
                       'corp_name': corp_name_v,
                       'origin': origin_v,
                       'ip': ip_v
                       }

            print("创建任务中,请求参数为：" + str(content))
            print("报文加密中......................")
            content_rsa = getEncryptData(key_pfx_v, key_password_v, content).replace('"', '')
            print("报文加密成功......................")

            data = {'member_id': member_id_v, 'terminal_id': terminal_id_v, 'data_content': content_rsa}
            headers = {"content-type": "application/json"}
            response = requests.post(env+'/gateway-data/security/v1/task/create',data= json.dumps(data),headers =headers,verify = False)
            response_json = json.loads(response.text)
            print("任务创建成功："+str(response_json))
            if response_json["data"] != None:
                tradeno = response_json["data"]["tradeNo"]
                err = ''
            else:
                tradeno = '999'
                err = response_json['errorMsg']
            LOG_INFO = account_v+','+password_v+','+ id_card_v+mobile_v+','+real_name_v
            trade_NO.append({'ID':ID_v,'tradeno': tradeno,'err':err,"env":env,"DESCRIBE":DESCRIBE_v,'CASE_ID':CASE_ID_v,'area':area_name_v,'log_info':LOG_INFO,'url':url_v})
        else:
            print("执行下一条...")


    print("开始查询结果信息...")

    rownum = 0
    for row_trande in range(len(trade_NO)):
        trade_NO_s = trade_NO[row_trande]
        errorMsgs = ''
        out = 0
        if trade_NO_s["tradeno"] != 999:
            rownum = rownum + 1
            #根据订单号查询订单状态
            header_s = {'memberId':memberId_v,'terminalId':terminalId_v}
            status = json.loads(requests.get(str(trade_NO_s["env"])+'/gateway-data/security/v1/task/status/'+str(trade_NO_s["tradeno"]),headers = header_s  ).text)
            print("订单状态为："+str(status))
            if status["errorMsg"] == None:
                while status["data"]["phase"] != 'DONE':
                    time.sleep(3)
                    status = json.loads(requests.get(str(trade_NO_s["env"])+'/gateway-data/security/v1/task/status/'+str(trade_NO_s["tradeno"]),headers = header_s  ).text)
                    print("订单状态为："+str(status))
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

            if out == 0:
                #根据订单号查询社保信息
                header_s = {'memberId': memberId_v, 'terminalId': terminalId_v}
                result = json.loads(requests.get(str(trade_NO_s["env"])+ '/data/security/v1/info/'+str(trade_NO_s["tradeno"]),headers = header_s).text)
                verify = 'Pass'
                report(rownum,trade_NO_s["ID"],trade_NO_s["DESCRIBE"],trade_NO_s["CASE_ID"],trade_NO_s["env"],trade_NO_s["tradeno"],trade_NO_s["area"],result,verify,errorMsgs)
                create_ownreport(trade_NO_s["area"],result,trade_NO_s["url"],trade_NO_s["log_info"])
            else:
                verify = 'Fail'
                report(rownum, trade_NO_s["ID"], trade_NO_s["DESCRIBE"], trade_NO_s["CASE_ID"], trade_NO_s["env"],trade_NO_s["tradeno"], trade_NO_s["area"], '', verify, errorMsgs)
        else:
            report(rownum, trade_NO_s["ID"], trade_NO_s["DESCRIBE"], trade_NO_s["CASE_ID"], trade_NO_s["env"],trade_NO_s["tradeno"], trade_NO_s["area"], '', verify, trade_NO_s['err'])
            print("执行下一条...")






    print("执行完毕.......")
    close_workbook()




