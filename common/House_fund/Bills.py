import requests
import json
import urllib3

def get_bills(env,memberId,terminalId,tradeno,year,page,size):
    print('获取账单信息中：请求参数为：'+env,memberId,terminalId,tradeno,year,page,size)
    url = env + '/data/fund/v2/bills/'+tradeno
    header = {'terminalId':terminalId,'memberId':memberId}
    params = {'year':year,
              'page':page,
              'size':size
              }
    urllib3.disable_warnings()
    response = requests.get(url,headers = header,params=params,verify = False)
    consu = response.elapsed.microseconds
    response = json.loads(response.text)
    print("接口返回为：" + str(response))
    return  response,consu

#trade_no_h = get_bills('http://test.xinyan.com','8000013112','8000013112', '201807121228560103022272','2018','1','20')