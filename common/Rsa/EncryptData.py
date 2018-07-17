import requests
import json

#rsa加密的接口地址
url ='http://10.0.20.60:21102/xinyantest-demo/xinyan_test/getEncryptData.do'

#定义数据加密函数，返回内容为加密后的数据
def getEncryptData(key_pfx,key_password,data):
    paramas = {"key_pfx": key_pfx,"key_password": key_password,"data_content": data}
    headers = {"content-type": "application/json"}

    response = requests.post(url,data = json.dumps(paramas),headers = headers).text

    return response

