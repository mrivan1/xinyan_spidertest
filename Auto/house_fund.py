from common.excel.Obtain_excel import getColumnIndex,readExcelDataByIndex
from common.House_fund.Arealist import Get_Arealist
from common.House_fund.Logindata import Get_Logindata
from common.House_fund.Task import task_create
from common.House_fund.Status import get_status
from common.House_fund.Result import get_result
from common.House_fund.Bills import get_bills
from common.House_fund.User import get_user
from common.House_fund.Loan import get_loan
from common.House_fund.Repay import get_repay
from common.excel.Report_excel import report,reports,close_workbook,report_x,create_ownreport
import readconfig
import xlrd
import time

def house_fund():
    #设置参数
    iterations = 30  #迭代轮数
    sleep_time = 30 #创建任务后的休眠时间，若任务较少则可以讲此数据调至100左右，若任务较多则调至20-30即可
    localRead = readconfig.ReadConfig()
    path = localRead.get_address('CASE')

    #获取案例所在的table,以及各项所在的列
    table_n = readExcelDataByIndex(path,2)[0]
    ID = getColumnIndex(table_n,'ID')
    DESCRIBE = getColumnIndex(table_n,'DESCRIBE')
    CASE_ID = getColumnIndex(table_n,'CASE_ID')
    IS_RUN = getColumnIndex(table_n,'IS_RUN')
    URI_NAME = getColumnIndex(table_n,'URI_NAME')
    ENV = getColumnIndex(table_n,'ENV')
    notify_url = getColumnIndex(table_n,'notify_url')
    user_id = getColumnIndex(table_n,'user_id')
    area_code = getColumnIndex(table_n,'area_code')
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
    size = getColumnIndex(table_n,'size')
    page = getColumnIndex(table_n,'page')
    year = getColumnIndex(table_n,'year')
    trade_no = getColumnIndex(table_n,'trade_no')
    Case_type = getColumnIndex(table_n,'Case_type')
    rely_on = getColumnIndex(table_n,'rely_on')
    errorMsg = getColumnIndex(table_n,'errorMsg')
    url_gjj = getColumnIndex(table_n,'URL')

    #解析案例
    data = xlrd.open_workbook(path)
    data_sheet_env = data.sheets()[0]

    #记录生成的tradeno
    trade_NO = []
    rownum = 0
    data_sheet = data.sheets()[2]
    for rown in range(data_sheet.nrows-1):
        rows = data_sheet.row_values(rown + 1)

        #判断需要执行（1）才进行以下操作
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
            URI_NAME_v = rows[URI_NAME]
            notify_url_v = rows[notify_url]
            user_id_v = rows[user_id]
            area_code_v = rows[area_code]
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
            size_v = rows[size]
            page_v = rows[page]
            year_v = rows[year]
            corp_name_v = rows[corp_name]
            trade_no_v = rows[trade_no]
            Case_type_v = rows[Case_type]
            rely_on_v = rows[rely_on]
            errorMsg_v = rows[errorMsg]
            url_gjj_v = rows[url_gjj]

            #只有案例类型为跑数据的时候才进行全自动跑任务
            if Case_type_v =='跑数据':
                rownum = rownum + 1
                areacode = Get_Arealist(env,memberId_v,terminalId_v,area_name_v)
                trade_no_h = task_create(env,member_id_v,terminal_id_v, key_pfx_v, key_password_v, notify_url_v, '1',
                                       areacode[0], account_v, password_v, login_type_v, id_card_v, mobile_v, real_name_v,
                                       sub_area_v, corp_account_v, corp_name_v, '2', ip_v)

                #讲创建的任务放入一个队列，任务全部创建完成后，在依次轮询队列里的任务，获取任务结果
                Trade_no = trade_no_h[0]

                trade_NO.append(
                    {"url_gjj": url_gjj_v, "rownum": rownum, "memberId": memberId_v, "terminalId": terminalId_v,
                     "tradeno": Trade_no, "ID": ID_v, "DESCRIBE": DESCRIBE_v, "CASE_ID": CASE_ID_v,
                     "Case_type": Case_type_v, "area_name": area_name_v, "env": env, "ENV": ENV_v,
                      "account": account_v, "password": password_v,"login_type":login_type_v,'areacode':areacode[0],
                     "id_card": id_card_v, "mobile": mobile_v, "real_name": real_name_v,"errorMsgs":trade_no_h[4],"corp_account":corp_account_v,"corp_name":corp_name_v})


                #控制创建任务的频率
                # time.sleep(3)


                # time_e = time.time()
                # work_sonsu = '%.2f'% (time_e - trade_no_h[2])




            # else:
            #     rownum1 = rownum1+1
            #     if rows[URI_NAME] == '获取城市公积金列表':
            #         area_code_rows = Get_Arealist(env,memberId_v,terminalId_v,area_name_v)
            #         consu = area_code_rows[2]
            #         if area_code_rows[1]["errorMsg"] != errorMsg_v and errorMsg_v != '' and area_code_rows[1]["errorMsg"] != None:
            #             verify = 'Fail'
            #         else:
            #             verify = 'Pass'
            #         report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v,'',verify,consu,errorMsg_v,str(area_code_rows[1]["errorMsg"]))
            #
            #     elif rows[URI_NAME] == '根据 areaCode 获取公积金登录信息':
            #         Logintype= Get_Logindata(env,memberId_v,terminalId_v,area_code_v)
            #         consu = Logintype[2]
            #         if Logintype[1]["errorMsg"] != errorMsg_v and errorMsg_v != '' and Logintype[1]["errorMsg"] != None:
            #             verify = 'Fail'
            #         else:
            #             verify = 'Pass'
            #         report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, '', verify, consu,errorMsg_v,str(Logintype[1]["errorMsg"]))
            #
            #     elif rows[URI_NAME] == '创建订单':
            #         trade_no_c  = task_create(env,member_id_v,terminal_id_v,key_pfx_v,key_password_v,notify_url_v,user_id_v,area_code_v,account_v,password_v,login_type_v,id_card_v,mobile_v,real_name_v,sub_area_v,corp_account_v,corp_name_v,origin_v,ip_v)
            #         trade_NO[ID_v] = trade_no_c[0]
            #         Trade_no = trade_no_c[0]
            #         consu = trade_no_c[3]
            #         if trade_no_c[1]["errorMsg"] != errorMsg_v and errorMsg_v != ''  :
            #             verify = 'Fail'
            #         elif trade_no_c[1]["errorMsg"] != None and trade_no_c[1]["errorMsg"] != errorMsg_v:
            #             verify = 'Fail'
            #         else:
            #             verify = 'Pass'
            #         report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, '', verify, consu,errorMsg_v,str(trade_no_c[1]["errorMsg"]))
            #
            #     elif rows[URI_NAME] == '订单执行状态查询':
            #         status_rows = get_status(env,memberId_v,terminalId_v,trade_NO[rely_on_v])
            #         Trade_no = trade_NO[rely_on_v]
            #         consu = status_rows[1]
            #         if status_rows[0]["errorMsg"]  !=  '订单不存在':
            #             while status_rows[0]["data"]["phase"] != 'DONE':
            #                 time.sleep(1)
            #                 status_rows = get_status(env, memberId_v, terminalId_v, Trade_no)
            #                 if status_rows[0]["errorCode"] != None:
            #                     break
            #                 elif status_rows[0]["data"]["phase_status"] == 'DONE_FAIL':
            #                     break
            #
            #         if status_rows[0]["errorMsg"] != errorMsg_v and errorMsg_v != '' :
            #             errr = str(status_rows[0]["errorMsg"])
            #             verify = 'Fail'
            #         elif  status_rows[0]["errorMsg"] != None and status_rows[0]["errorMsg"] != errorMsg_v :
            #             errr = str(status_rows[0]["errorMsg"])
            #             verify = 'Fail'
            #         elif  status_rows[0]["data"]["phase_status"] == 'DONE_FAIL' :
            #             errr = str(status_rows[0]["data"]["description"])
            #             verify = 'Fail'
            #         else:
            #             verify = 'Pass'
            #         report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, Trade_no, verify,consu,errorMsg_v,errr)
            #
            #     elif rows[URI_NAME] == '获取公积金信息':
            #         if rely_on_v != '':
            #             result = get_result(env,memberId_v,terminalId_v,trade_NO[rely_on_v])
            #             Trade_no = trade_NO[rely_on_v]
            #         else:
            #             result = get_result(env,memberId_v,terminalId_v,trade_no_v)
            #             Trade_no = trade_no_v
            #         consu = result[1]
            #         if result[0]["errorMsg"] != errorMsg_v and errorMsg_v != '' :
            #             verify = 'Fail'
            #         elif result[0]["errorMsg"] != None and result[0]["errorMsg"] != errorMsg_v:
            #             verify = 'Fail'
            #         else:
            #             verify = 'Pass'
            #         report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, Trade_no, verify,
            #                      consu,errorMsg_v,str(result[0]["errorMsg"]))
            #
            #     elif rows[URI_NAME] == '根据订单号、年份分页查询公积金缴纳信息':
            #         if rely_on_v != '':
            #             bills = get_bills(env, memberId_v, terminalId_v, trade_NO[rely_on_v], year_v, page_v, size_v)
            #             Trade_no = trade_NO[rely_on_v]
            #         else:
            #             bills = get_bills(env, memberId_v, terminalId_v, trade_no_v, year_v, page_v, size_v)
            #             Trade_no = trade_no_v
            #
            #         consu = bills[1]
            #         if bills[0]["errorMsg"] != errorMsg_v and errorMsg_v != '' :
            #             verify = 'Fail'
            #         elif bills[0]["errorMsg"] != None and bills[0]["errorMsg"] != errorMsg_v :
            #             verify = 'Fail'
            #         else:
            #             verify = 'Pass'
            #         report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, Trade_no, verify,
            #                      consu,errorMsg_v,str(bills[0]["errorMsg"]))
            #
            #     elif rows[URI_NAME] == '根据订单号查询公积金账户信息':
            #         if rely_on_v != '':
            #             userinfo_rows = get_user(env, memberId_v, terminalId_v, trade_NO[rely_on_v])
            #             Trade_no = trade_NO[rely_on_v]
            #         else:
            #             userinfo_rows = get_user(env, memberId_v, terminalId_v, trade_no_v)
            #             Trade_no = trade_no_v
            #
            #
            #         consu = userinfo_rows[1]
            #         if userinfo_rows[0]["errorMsg"] != errorMsg_v and errorMsg_v != '' :
            #             verify = 'Fail'
            #         elif userinfo_rows[0]["errorMsg"] != None and userinfo_rows[0]["errorMsg"] != errorMsg_v:
            #             verify = 'Fail'
            #         else:
            #             verify = 'Pass'
            #         report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, Trade_no, verify,
            #                      consu,errorMsg_v,str(userinfo_rows[0]["errorMsg"]))
            #
            #     elif rows[URI_NAME] == '根据订单号查询公积金贷款信息':
            #         if rely_on_v != '':
            #             loaninfo_rows = get_loan(env, memberId_v, terminalId_v, trade_NO[rely_on_v])
            #             Trade_no = trade_NO[rely_on_v]
            #         else:
            #             loaninfo_rows = get_loan(env, memberId_v, terminalId_v, trade_no_v)
            #             Trade_no = trade_no_v
            #
            #
            #         consu = loaninfo_rows[1]
            #         if loaninfo_rows[0]["errorMsg"] != errorMsg_v and errorMsg_v != '' :
            #             verify = 'Fail'
            #         elif loaninfo_rows[0]["errorMsg"] != None and loaninfo_rows[0]["errorMsg"] != errorMsg_v :
            #             verify = 'Fail'
            #         else:
            #             verify = 'Pass'
            #         report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, Trade_no, verify,
            #                      consu,errorMsg_v,str(loaninfo_rows[0]["errorMsg"]))
            #
            #     elif rows[URI_NAME] == '根据订单号查询公积金还款信息':
            #         if rely_on_v != '':
            #             repayinfo_rows = get_repay(env, memberId_v, terminalId_v, trade_NO[rely_on_v])
            #             Trade_no = trade_NO[rely_on_v]
            #         else:
            #             repayinfo_rows = get_repay(env, memberId_v, terminalId_v, trade_no_v)
            #             Trade_no = trade_no_v
            #
            #         consu = repayinfo_rows[1]
            #         if repayinfo_rows[0]["errorMsg"] != errorMsg_v and errorMsg_v != '' :
            #             verify = 'Fail'
            #         elif repayinfo_rows[0]["errorMsg"] != None and repayinfo_rows[0]["errorMsg"] != errorMsg_v:
            #             verify = 'Fail'
            #         else:
            #             verify = 'Pass'
            #         report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, Trade_no, verify,
            #                      consu,errorMsg_v,str(repayinfo_rows[0]["errorMsg"]))
            # print('开始下一条执行..........')

        else:
            continue
    #轮询订单队列，依次获取订单信息

    time.sleep(sleep_time)
    def get_results(trade_NO,trade_NO_re):
        for ii in range(len(trade_NO)):
            out = 0
            Trade_no = trade_NO[ii]["tradeno"]
            account_v = trade_NO[ii]["account"]
            password_v = trade_NO[ii]["password"]
            id_card_v = trade_NO[ii]["id_card"]
            mobile_v = trade_NO[ii]["mobile"]
            real_name_v = trade_NO[ii]["real_name"]
            env = trade_NO[ii]["env"]
            memberId_v = trade_NO[ii]["memberId"]
            terminalId_v = trade_NO[ii]["terminalId"]
            rownum = trade_NO[ii]["rownum"]
            ID_v = trade_NO[ii]["ID"]
            DESCRIBE_v = trade_NO[ii]["DESCRIBE"]
            CASE_ID_v = trade_NO[ii]["CASE_ID"]
            Case_type_v = trade_NO[ii]["Case_type"]
            ENV_v = trade_NO[ii]["ENV"]
            area_name_v = trade_NO[ii]["area_name"]
            url_gjj_v = trade_NO[ii]["url_gjj"]
            login_type_v = trade_NO[ii]["login_type"]
            areacode = trade_NO[ii]["areacode"]
            corp_account_v = trade_NO[ii]["corp_account"]
            corp_name_v = trade_NO[ii]["corp_name"]
            try:
                if Trade_no == '999':
                    out = 1
                    errorMsgs = trade_NO[ii]["errorMsgs"]
                else:
                    print("开始查询 %s 地区公积金....." % area_name_v)
                    status = get_status(env, memberId_v, terminalId_v, Trade_no)[0]
                    retry_times = 0
                    try_times = 0
                    if status["errorMsg"] == None:
                        while status["data"]["phase"] != 'DONE':
                            status = get_status(env, memberId_v, terminalId_v, Trade_no)[0]
                            errorMsgs = ''
                            try_times = try_times + 1
                            #对返回错误类型为官网繁忙或者验证码错误的时候增加重试机制
                            if '验证码' in status["data"]["description"] or '官网' in status["data"]["description"]  :
                                print("验证码错误，正在重试")
                                trade_no_h = task_create(env, member_id_v, terminal_id_v, key_pfx_v, key_password_v,
                                                         '', '1',
                                                         areacode, account_v, password_v, login_type_v,
                                                         id_card_v, mobile_v, real_name_v,
                                                         '',corp_account_v,corp_name_v, '2', '')
                                trade_NO_re.append(
                                    {"url_gjj": url_gjj_v, "rownum": rownum, "memberId": memberId_v,
                                     "terminalId": terminalId_v,
                                     "tradeno": trade_no_h[0], "ID": ID_v, "DESCRIBE": DESCRIBE_v, "CASE_ID": CASE_ID_v,
                                     "Case_type": Case_type_v, "area_name": area_name_v, "env": env, "ENV": ENV_v,
                                     "account": account_v, "password": password_v, "login_type": login_type_v,
                                     'areacode': areacode,
                                     "id_card": id_card_v, "mobile": mobile_v, "real_name": real_name_v,
                                     "errorMsgs": trade_no_h[4],"corp_account":corp_account_v,"corp_name":corp_name_v})
                                out = 2
                                break
                            elif status["errorCode"] != None:
                                errorMsgs = str(status['errorMsg'])
                                out = 1
                                break
                            elif status["data"]["phase_status"] == 'DONE_FAIL':
                                try:
                                    errorMsgs = str(status["data"]["description"])
                                except:
                                    pass
                                out = 1
                                break
                            elif try_times > 40:
                                print("任务超时")
                                out = 1
                                errorMsgs = "任务超时"
                                break
                    else:
                        errorMsgs = str(status['errorMsg'])
                        out = 1
            except:
                pass

            if out == 0:
                load_info = account_v + ',' + password_v + ',' + id_card_v + mobile_v + ',' + real_name_v
                result = get_result(env, memberId_v, terminalId_v, Trade_no)[0]
                bills_2014 = get_bills(env, memberId_v, terminalId_v, Trade_no, '2014', '1', '20')[0]
                bills_2015 = get_bills(env, memberId_v, terminalId_v, Trade_no, '2015', '1', '20')[0]
                bills_2016 = get_bills(env, memberId_v, terminalId_v, Trade_no, '2016', '1', '20')[0]
                bills_2017 = get_bills(env, memberId_v, terminalId_v, Trade_no, '2017', '1', '20')[0]
                bills_2018 = get_bills(env, memberId_v, terminalId_v, Trade_no, '2018', '1', '20')[0]
                userinfo = get_user(env, memberId_v, terminalId_v, Trade_no)[0]
                loaninfo = get_loan(env, memberId_v, terminalId_v, Trade_no)[0]
                repayinfo = get_repay(env, memberId_v, terminalId_v, Trade_no)[0]
                verify = 'Pass'
                report(rownum, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, '', ENV_v, area_name_v, Trade_no, result,
                       bills_2014, bills_2015, bills_2016, bills_2017, bills_2018, userinfo, loaninfo, repayinfo, verify,
                       0, '')
                if result["errorCode"] == None:
                    create_ownreport(area_name_v, ID_v, Trade_no, result, url_gjj_v, load_info,rownum)
                else:
                    pass
                time.sleep(4)
            elif out == 1:
                verify = 'Fail'
                report(rownum, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, area_name_v, Trade_no, '', '', '',
                       '', '', '', '', '', '', verify, 0, errorMsgs)
            else:
                pass
            print('%s 地区公积金查询完毕，开始查询下一个地区......' % area_name_v)


        return trade_NO_re

    names = {'trade_NO_re' + str(iname): 'trade_NO_re' + str(iname) for iname in range(iterations) }
    names['trade_NO_re0'] = []
    get_results(trade_NO,names['trade_NO_re0'])
    for tr in range(iterations):
        names['trade_NO_re%s'%str(tr+1)] = []
        if len(names['trade_NO_re%s'% str(tr)]) != 0:
            time.sleep(sleep_time*2)
            get_results(names['trade_NO_re%s'%str(tr)],names['trade_NO_re%s'%str(tr+1)])
        else:
            break


    print("执行完毕.......")
    reports(data_sheet.nrows,rownum)
    close_workbook()









