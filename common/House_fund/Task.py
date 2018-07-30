import json
import requests
from common.Rsa.EncryptData import getEncryptData
from common.rand import membertrans,tradedate
import time
import urllib3


def task_create(env,member_id,terminal_id,key_pfx,key_password,notify_url,user_id,area_code,account,password,login_type,id_card,mobile,real_name,sub_area,corp_account,corp_name,origin,ip):
    content = {'member_id': member_id ,
               'terminal_id': terminal_id,
               'member_trans_date': tradedate(),
               'member_trans_id':membertrans() ,
               'notify_url':notify_url,
               'user_id':user_id,
               'area_code':area_code,
               'account':account,
               'password':password,
               'login_type':login_type,
               'id_card':id_card,
               'mobile':mobile,
               'real_name':real_name,
               'sub_area':sub_area,
               'corp_account':corp_account,
               'corp_name':corp_name,
               'origin':origin,
               'ip':ip
               }

    print("创建任务中,请求参数为："+str(content))
    content_rsa = getEncryptData(key_pfx,key_password,content).replace('"','')

    url =env + '/gateway-data/fund/v1/task/create'
    data  =  {'member_id': member_id,'terminal_id': terminal_id,'data_content': content_rsa}
    headers = {"content-type": "application/json"}
    time_s = time.time()
    urllib3.disable_warnings()
    response = requests.post(url,data= json.dumps(data),headers =headers,verify = False)
    consu = response.elapsed.microseconds
    response = json.loads(response.text)

    if response["data"] != None:
        tradeno = response["data"]["tradeNo"]
        err = ''
    else:
        tradeno = '999'
        err = response['errorMsg']
    return  tradeno,response,time_s,consu,err

if __name__ == '__main__':
    trade_no_h = task_create('http://test.xinyan.com','8000013112','8000013112', '8000013112_pri.pfx', '217531', '', '1','524000', '440825198606250012', 'a.123456', '13', '', '', '','', '', '', '2', '')
    print(trade_no_h[1])