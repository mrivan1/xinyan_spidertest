import datetime
import random
#随机生成请求号
def ranreq():
    req = 'req'+str(datetime.datetime.now()).replace('-','').replace(' ','').replace(':','').replace('.','')
    return req

#生成交易日期函数：15位，年月日时分秒
def tradedate():
    d = str(datetime.datetime.now().strftime('%Y%m%d%H%M%S'))
    return d

#随机生成member_transid
def membertrans():
    transid = str(datetime.datetime.now().strftime('%Y%m%d%H%M%S'))+str(random.randint(1,100))+str(random.randint(1,100))
    return transid


#对excel中值去空格和特殊符号
def data_fc(x):
    data = str(x).replace(' ','').replace('-','')
    return data


