import requests
from common.House_fund.Arealist import Get_Arealist
import json
import urllib3

#根据城市名获取登录方式：login_type
def Get_Logindata(env,memberId,terminalId,areacode):
    #q请求url
    url = env+'/gateway-data/fund/v1/login/'+areacode
    #请求参数
    headers = {'memberId': memberId, 'terminalId':terminalId}
    print("获取登录方式中，请求参数为："+env,memberId,terminalId,areacode)
    urllib3.disable_warnings()
    response = requests.get(url,headers = headers,verify = False)
    consu = response.elapsed.microseconds
    response = json.loads(response.text)
    area = response['data']
    print("接口返回为：" + str(response))
    logtype =[]
    for i in range(len(area)):
        logtype.append(area[i]['login_type'])
    return logtype,response,consu



