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

localRead = readconfig.ReadConfig()
path = localRead.get_address('CASE')

#获取案例所在的table,以及各项所在的列
table_n = readExcelDataByIndex(path,1)[0]
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

#解析案例
data = xlrd.open_workbook(path)
data_sheet_env = data.sheets()[0]

#记录生成的tradeno
trade_NO = {}
rownum = 0
rownum1 = 0
data_sheet = data.sheets()[1]
for rown in range(data_sheet.nrows-1):
    rows = data_sheet.row_values(rown + 1)
    out = 0
    #判断需要执行才进行以下操作
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


        if Case_type_v =='跑数据':
            rownum = rownum + 1
            areacode = Get_Arealist(env,memberId_v,terminalId_v,area_name_v)
            trade_no_h = task_create(env,member_id_v,terminal_id_v, key_pfx_v, key_password_v, notify_url_v, '1',
                                   areacode[0], account_v, password_v, login_type_v, id_card_v, mobile_v, real_name_v,
                                   sub_area_v, corp_account_v, corp_name_v, '2', ip_v)

            Trade_no = trade_no_h[0]
            if Trade_no == '999':
                out =1
                errorMsgs = trade_no_h[4]
            else:
                status = get_status(env, memberId_v, terminalId_v, Trade_no)[0]
                if status["errorMsg"] == None:
                    while status["data"]["phase"] != 'DONE':
                        time.sleep(1)
                        status = get_status(env, memberId_v, terminalId_v, Trade_no)[0]
                        if status["errorCode"] != None :
                            errorMsgs = str(status['errorMsg'])
                            out = 1
                            break
                        elif status["data"]["phase_status"] == 'DONE_FAIL':
                            errorMsgs = str(status["data"]["description"])
                            out = 1
                            break
                else:
                    errorMsgs = str(status['errorMsg'])
                    out = 1

            time_e = time.time()
            work_sonsu = '%.2f'% (time_e - trade_no_h[2])

            if out == 0:
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
                report(rownum,ID_v,DESCRIBE_v,CASE_ID_v,Case_type_v,URI_NAME_v,ENV_v,area_name_v,Trade_no,result,bills_2014,bills_2015,bills_2016,bills_2017,bills_2018,userinfo,loaninfo,repayinfo,verify,work_sonsu,'')
                create_ownreport(area_name_v,ID_v,Trade_no,result)
            else:
                verify = 'Fail'
                report(rownum, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v,area_name_v, Trade_no,'','','','','','','','','',verify,0,errorMsgs)

        else:
            rownum1 = rownum1+1
            if rows[URI_NAME] == '获取城市公积金列表':
                area_code_rows = Get_Arealist(env,memberId_v,terminalId_v,area_name_v)
                consu = area_code_rows[2]
                if area_code_rows[1]["errorMsg"] != errorMsg_v and errorMsg_v != '' and area_code_rows[1]["errorMsg"] != None:
                    verify = 'Fail'
                else:
                    verify = 'Pass'
                report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v,'',verify,consu,errorMsg_v,str(area_code_rows[1]["errorMsg"]))







            elif rows[URI_NAME] == '根据 areaCode 获取公积金登录信息':
                Logintype= Get_Logindata(env,memberId_v,terminalId_v,area_code_v)
                consu = Logintype[2]
                if Logintype[1]["errorMsg"] != errorMsg_v and errorMsg_v != '' and Logintype[1]["errorMsg"] != None:
                    verify = 'Fail'
                else:
                    verify = 'Pass'
                report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, '', verify, consu,errorMsg_v,str(Logintype[1]["errorMsg"]))

            elif rows[URI_NAME] == '创建订单':
                trade_no_c  = task_create(env,member_id_v,terminal_id_v,key_pfx_v,key_password_v,notify_url_v,user_id_v,area_code_v,account_v,password_v,login_type_v,id_card_v,mobile_v,real_name_v,sub_area_v,corp_account_v,corp_name_v,origin_v,ip_v)
                trade_NO[ID_v] = trade_no_c[0]
                Trade_no = trade_no_c[0]
                consu = trade_no_c[3]
                if trade_no_c[1]["errorMsg"] != errorMsg_v and errorMsg_v != ''  :
                    verify = 'Fail'
                elif trade_no_c[1]["errorMsg"] != None and trade_no_c[1]["errorMsg"] != errorMsg_v:
                    verify = 'Fail'
                else:
                    verify = 'Pass'
                report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, '', verify, consu,errorMsg_v,str(trade_no_c[1]["errorMsg"]))

            elif rows[URI_NAME] == '订单执行状态查询':
                status_rows = get_status(env,memberId_v,terminalId_v,trade_NO[rely_on_v])
                Trade_no = trade_NO[rely_on_v]
                consu = status_rows[1]
                if status_rows[0]["errorMsg"]  !=  '订单不存在':
                    while status_rows[0]["data"]["phase"] != 'DONE':
                        time.sleep(1)
                        status_rows = get_status(env, memberId_v, terminalId_v, Trade_no)
                        if status_rows[0]["errorCode"] != None:
                            break
                        elif status_rows[0]["data"]["phase_status"] == 'DONE_FAIL':
                            break

                if status_rows[0]["errorMsg"] != errorMsg_v and errorMsg_v != '' :
                    errr = str(status_rows[0]["errorMsg"])
                    verify = 'Fail'
                elif  status_rows[0]["errorMsg"] != None and status_rows[0]["errorMsg"] != errorMsg_v :
                    errr = str(status_rows[0]["errorMsg"])
                    verify = 'Fail'
                elif  status_rows[0]["data"]["phase_status"] == 'DONE_FAIL' :
                    errr = str(status_rows[0]["data"]["description"])
                    verify = 'Fail'
                else:
                    verify = 'Pass'
                report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, Trade_no, verify,consu,errorMsg_v,errr)

            elif rows[URI_NAME] == '获取公积金信息':
                if rely_on_v != '':
                    result = get_result(env,memberId_v,terminalId_v,trade_NO[rely_on_v])
                    Trade_no = trade_NO[rely_on_v]
                else:
                    result = get_result(env,memberId_v,terminalId_v,trade_no_v)
                    Trade_no = trade_no_v
                consu = result[1]
                if result[0]["errorMsg"] != errorMsg_v and errorMsg_v != '' :
                    verify = 'Fail'
                elif result[0]["errorMsg"] != None and result[0]["errorMsg"] != errorMsg_v:
                    verify = 'Fail'
                else:
                    verify = 'Pass'
                report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, Trade_no, verify,
                             consu,errorMsg_v,str(result[0]["errorMsg"]))

            elif rows[URI_NAME] == '根据订单号、年份分页查询公积金缴纳信息':
                if rely_on_v != '':
                    bills = get_bills(env, memberId_v, terminalId_v, trade_NO[rely_on_v], year_v, page_v, size_v)
                    Trade_no = trade_NO[rely_on_v]
                else:
                    bills = get_bills(env, memberId_v, terminalId_v, trade_no_v, year_v, page_v, size_v)
                    Trade_no = trade_no_v

                consu = bills[1]
                if bills[0]["errorMsg"] != errorMsg_v and errorMsg_v != '' :
                    verify = 'Fail'
                elif bills[0]["errorMsg"] != None and bills[0]["errorMsg"] != errorMsg_v :
                    verify = 'Fail'
                else:
                    verify = 'Pass'
                report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, Trade_no, verify,
                             consu,errorMsg_v,str(bills[0]["errorMsg"]))

            elif rows[URI_NAME] == '根据订单号查询公积金账户信息':
                if rely_on_v != '':
                    userinfo_rows = get_user(env, memberId_v, terminalId_v, trade_NO[rely_on_v])
                    Trade_no = trade_NO[rely_on_v]
                else:
                    userinfo_rows = get_user(env, memberId_v, terminalId_v, trade_no_v)
                    Trade_no = trade_no_v


                consu = userinfo_rows[1]
                if userinfo_rows[0]["errorMsg"] != errorMsg_v and errorMsg_v != '' :
                    verify = 'Fail'
                elif userinfo_rows[0]["errorMsg"] != None and userinfo_rows[0]["errorMsg"] != errorMsg_v:
                    verify = 'Fail'
                else:
                    verify = 'Pass'
                report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, Trade_no, verify,
                             consu,errorMsg_v,str(userinfo_rows[0]["errorMsg"]))

            elif rows[URI_NAME] == '根据订单号查询公积金贷款信息':
                if rely_on_v != '':
                    loaninfo_rows = get_loan(env, memberId_v, terminalId_v, trade_NO[rely_on_v])
                    Trade_no = trade_NO[rely_on_v]
                else:
                    loaninfo_rows = get_loan(env, memberId_v, terminalId_v, trade_no_v)
                    Trade_no = trade_no_v


                consu = loaninfo_rows[1]
                if loaninfo_rows[0]["errorMsg"] != errorMsg_v and errorMsg_v != '' :
                    verify = 'Fail'
                elif loaninfo_rows[0]["errorMsg"] != None and loaninfo_rows[0]["errorMsg"] != errorMsg_v :
                    verify = 'Fail'
                else:
                    verify = 'Pass'
                report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, Trade_no, verify,
                             consu,errorMsg_v,str(loaninfo_rows[0]["errorMsg"]))

            elif rows[URI_NAME] == '根据订单号查询公积金还款信息':
                if rely_on_v != '':
                    repayinfo_rows = get_repay(env, memberId_v, terminalId_v, trade_NO[rely_on_v])
                    Trade_no = trade_NO[rely_on_v]
                else:
                    repayinfo_rows = get_repay(env, memberId_v, terminalId_v, trade_no_v)
                    Trade_no = trade_no_v

                consu = repayinfo_rows[1]
                if repayinfo_rows[0]["errorMsg"] != errorMsg_v and errorMsg_v != '' :
                    verify = 'Fail'
                elif repayinfo_rows[0]["errorMsg"] != None and repayinfo_rows[0]["errorMsg"] != errorMsg_v:
                    verify = 'Fail'
                else:
                    verify = 'Pass'
                report_x(rownum1, ID_v, DESCRIBE_v, CASE_ID_v, Case_type_v, URI_NAME_v, ENV_v, Trade_no, verify,
                             consu,errorMsg_v,str(repayinfo_rows[0]["errorMsg"]))
        print('开始下一条执行..........')

    else:
        continue

print("执行完毕.......")
reports(data_sheet.nrows,rownum)
close_workbook()








