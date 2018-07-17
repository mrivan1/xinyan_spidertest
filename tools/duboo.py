# -*- coding:utf-8 -*-

from pyhessian.client import HessianProxy

from pyhessian import protocol

import json


def InvokeHessian(service, interface, method, req, retcode='000000'):
    url = 'dubbo://10.0.23.66:20882/' + service + '.' + interface

    print('URL:\t%s' % url)

    print('Method:\t%s' % method)


    res = getattr(HessianProxy(url), method)(req)

    print('Res:\t%s' % json.dumps(res, ensure_ascii=False))




if __name__ == '__main__':
    service = 'com.xinyan.channel.facade'

    interface = 'AuthBankCardFacade'

    method = 'authBankCardService'

    req = protocol.object_factory('com.xinyan.channel.facade.AuthBankCardFacade', channelId='1104901110',
                                  reqTransId_length='32', accNo='6212262610002607004', certificateType='123',
                                  cardType='101', idNo='612730198501061126', idName='高利娜', mobile='', validTime='',
                                  cvv2='')

    InvokeHessian(service, interface, method, req)
