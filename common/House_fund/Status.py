import requests
import json
import urllib3

def get_status(env,memberId,terminalId,tradeno):
    print("获取登录状态中，参数为：" + env, memberId, terminalId, tradeno)
    url = env + '/gateway-data/fund/v1/task/status/'+tradeno
    header = {'memberId':memberId,'terminalId':terminalId}
    urllib3.disable_warnings()
    response = requests.get(url,headers = header,verify = False)
    consu = response.elapsed.microseconds
    response = json.loads(response.text)
    print("接口返回为：" + str(response))
    return response,consu


#trade_no_h = get_status('http://test.xinyan.com','8000013112','8000013112', '201807121228560103022272')

