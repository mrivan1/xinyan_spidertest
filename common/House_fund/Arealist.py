import requests
import json
import urllib3

def Get_Arealist(env,memberId,terminalId,areaname):
    #q请求url
    url = env+'/gateway-data/fund/v1/arealist'
    #请求参数
    headers = {'memberId': memberId, 'terminalId': terminalId}

    print("获取登录地区信息中，参数为："+env,memberId,terminalId,areaname)
    urllib3.disable_warnings()
    response = requests.get(url,headers = headers,verify = False)
    consu = response.elapsed.microseconds
    response = json.loads(response.text)
    for row in response["data"]:
        if areaname == row["city_name"]:
            return row['area_code'],response,consu
            break
        else:
            continue

if __name__ == '__main__':
    a = Get_Arealist("https://api.xinyan.com","1107602","32912","运城")
    print(a[0])





