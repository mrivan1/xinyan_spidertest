import requests
import json
import urllib3

def get_loan(env,memberId,terminalId,tradeno):
    print('获取个人贷款信息中：请求参数为：' + env, memberId, terminalId, tradeno)
    url = env + '/data/fund/v2/loan/'+tradeno
    header = {'terminalId':terminalId,'memberId':memberId}
    urllib3.disable_warnings()
    response = requests.get(url,headers = header,verify = False)
    consu = response.elapsed.microseconds
    response = json.loads(response.text)
    print("接口返回为：" + str(response))
    return  response,consu

