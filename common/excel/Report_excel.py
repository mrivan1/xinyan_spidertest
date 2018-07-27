#encoding='gb18030'
import xlsxwriter
from common.rand import tradedate

# 新建结果集
now = tradedate()
filename = '测试报告' + now + '.xlsx'
workbook = xlsxwriter.Workbook('D:/DevCode/data/spider/report/Housefound/'+filename)

# 创建sheet
work_sheet1 = workbook.add_worksheet('公积金-报告')
work_sheet = workbook.add_worksheet('公积金回归明细')
work_sheet2 = workbook.add_worksheet('公积金-案例详情')

# 结果包含的信息
row = ['ID', '案例描述', '案例ID', '案例类型', '接口名称', '环境','地区', '订单号', '公积金信息','2014年缴纳信息','2015年缴纳信息','2016年缴纳信息','2017年缴纳信息','2018年缴纳信息','个人信息','贷款信息','还款信息','是否通过','爬取耗时（S）','失败原因']
for s in range(len(row)):
    work_sheet.write(0, s, row[s])
work_sheet.freeze_panes(1, 0)

row1 = ['ID', '案例描述', '案例ID', '案例类型', '接口名称', '环境', '订单号','是否通过','接口耗时(μs)','期望值','实际值']
for sa in range(len(row1)):
    work_sheet2.write(0, sa, row1[sa])
work_sheet2.freeze_panes(1, 0)

def report(rownum,ID,DESCRIBE,CASE_ID,Case_type,URI_NAME,ENV,area,Trade_no,result,bills_2014,bills_2015,bills_2016,bills_2017,bills_2018,userinfo,loaninfo,repayinfo,verify,work_sonsu,err):
    row1 = [ID,DESCRIBE,CASE_ID,Case_type,URI_NAME,ENV,area,Trade_no,result,bills_2014,bills_2015,bills_2016,bills_2017,bills_2018,userinfo,loaninfo,repayinfo,verify,work_sonsu,err]
    for k in range(len(row1)):
        work_sheet.write(rownum, k, str(row1[k]))
    work_sheet.write_number(rownum ,18,float(row1[18]))
    if verify == 'Pass':
        work_sheet.write_url(rownum ,20, url = 'internal:'+str(area)+str(ID)+'报告详情!A1',string='报告详情')
    else:
        pass

def report_x(rownum,ID,DESCRIBE,CASE_ID,Case_type,URI_NAME,ENV,Trade_no,verify,consu,expect,actual):
    row1 = [ID,DESCRIBE,CASE_ID,Case_type,URI_NAME,ENV,Trade_no,verify,consu,expect,actual]
    for k in range(len(row1)):
        work_sheet2.write(rownum, k, str(row1[k]))
    work_sheet2.write_number(rownum, 8,float(row1[8]))

#创建公积金个人报告
def create_ownreport(areaname,ID_v,trande_no,result_s,url_gjj,loan_info,rownum):
    print("生成个人报告中............")
    if result_s !='':
        work_sheets = workbook.add_worksheet(areaname+ID_v + '报告详情')
        work_sheets.hide_gridlines(2)
        merge_format_t = workbook.add_format({'font_size': 20, 'align': 'center', 'valign': 'vcenter', 'font_name': '宋体'})
        merge_format_b = workbook.add_format({'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'border': 2, 'font_name': '宋体','fg_color':'#CD2626'})
        cell_format_b = workbook.add_format({'font_size': 10, 'font_color': 'black', 'font_name': '宋体', 'align': 'justify', 'border': 1,'fg_color':'#7CCD7C'})
        cell_format_bv = workbook.add_format({'font_size': 10, 'font_color': 'black', 'font_name': '宋体', 'align': 'right', 'border': 1})
        cell_format_bc = workbook.add_format({'font_size': 10, 'font_color': 'black', 'font_name': '宋体', 'align': 'justify', 'border': 1,'fg_color':'#7CCD7C'})
        border_l = workbook.add_format({'left': 2})
        border_t = workbook.add_format({'top': 2})
        border_h2 = workbook.add_format({'top': 2, 'left': 2})

        work_sheets.write_url('A1',url=  'internal:公积金回归明细!U'+str(rownum+1),string='明细页')
        work_sheets.merge_range('C3:O6', '公积金个人报告', merge_format_t)
        work_sheets.merge_range('C7:D7', '新颜订单号：', cell_format_bc)
        work_sheets.merge_range('E7:G7', str(trande_no), cell_format_bv)
        work_sheets.merge_range('H7:I7', '官网网址：', cell_format_bc)
        work_sheets.merge_range('J7:O7',url_gjj , cell_format_bv)
        work_sheets.merge_range('C8:D8', '登录信息：', cell_format_bc)
        work_sheets.merge_range('E8:O8',loan_info , cell_format_bv)

        #基本信息
        work_sheets.merge_range('C9:O9', '基本信息', merge_format_b)
        userinfo = result_s["data"]["user_info"]
        work_sheets.merge_range('C10:D10', '姓名', cell_format_b)
        work_sheets.merge_range('G10:H10', '补贴公积金账户余额', cell_format_b)
        work_sheets.merge_range('K10:L10', '个人月度缴存', cell_format_b)
        work_sheets.merge_range('C11:D11', '性别', cell_format_b)
        work_sheets.merge_range('G11:H11', '补贴月缴存', cell_format_b)
        work_sheets.merge_range('K11:L11', '月度总缴存', cell_format_b)
        work_sheets.merge_range('C12:D12', '出生日期', cell_format_b)
        work_sheets.merge_range('G12:H12', '缴存状态', cell_format_b)
        work_sheets.merge_range('K12:L12', '企业缴存比例', cell_format_b)
        work_sheets.merge_range('C13:D13', '手机号码', cell_format_b)
        work_sheets.merge_range('G13:H13', '身份证号码', cell_format_b)
        work_sheets.merge_range('K13:L13', '个人缴存比例', cell_format_b)
        work_sheets.merge_range('C14:D14', '邮箱', cell_format_b)
        work_sheets.merge_range('G14:H14', '证件类型', cell_format_b)
        work_sheets.merge_range('K14:L14', '补贴公积金公司缴存比例', cell_format_b)
        work_sheets.merge_range('C15:D15', '客户号', cell_format_b)
        work_sheets.merge_range('G15:H15', '通讯地址', cell_format_b)
        work_sheets.merge_range('K15:L15', '补贴公积金个人缴存比例', cell_format_b)
        work_sheets.merge_range('C16:D16', '公积金账号', cell_format_b)
        work_sheets.merge_range('G16:H16', '企业账户号码', cell_format_b)
        work_sheets.merge_range('K16:L16', '缴存基数', cell_format_b)
        work_sheets.merge_range('C17:D17', '账户余额', cell_format_b)
        work_sheets.merge_range('G17:H17', '当前缴存企业名称', cell_format_b)
        work_sheets.merge_range('K17:L17', '最新缴存日期', cell_format_b)
        work_sheets.merge_range('C18:D18', '公积金余额', cell_format_b)
        work_sheets.merge_range('G18:H18', '企业月度缴存', cell_format_b)
        work_sheets.merge_range('K18:L18', '开户日期', cell_format_b)
        if len(userinfo) != 0:
            work_sheets.merge_range('E10:F10', userinfo["real_name"], cell_format_bv)
            work_sheets.merge_range('I10:J10', userinfo["subsidy_balance"], cell_format_bv)
            work_sheets.merge_range('M10:O10', userinfo["monthly_customer_income"], cell_format_bv)
            work_sheets.merge_range('E11:F11', userinfo["gender"], cell_format_bv)
            work_sheets.merge_range('I11:J11', userinfo["subsidy_income"], cell_format_bv)
            work_sheets.merge_range('M11:O11', userinfo["monthly_total_income"], cell_format_bv)
            work_sheets.merge_range('E12:F12', userinfo["birthday"], cell_format_bv)
            work_sheets.merge_range('I12:J12', userinfo["pay_status"], cell_format_bv)
            work_sheets.merge_range('M12:O12', userinfo["corporation_ratio"], cell_format_bv)
            work_sheets.merge_range('E13:F13', userinfo["mobile"], cell_format_bv)
            work_sheets.merge_range('I13:J13', userinfo["id_card"], cell_format_bv)
            work_sheets.merge_range('M13:O13', userinfo["customer_ratio"], cell_format_bv)
            work_sheets.merge_range('E14:F14', userinfo["email"], cell_format_bv)
            work_sheets.merge_range('I14:J14', userinfo["card_type"], cell_format_bv)
            work_sheets.merge_range('M14:O14', userinfo["subsidy_corporation_ratio"], cell_format_bv)
            work_sheets.merge_range('E15:F15', userinfo["customer_number"], cell_format_bv)
            work_sheets.merge_range('I15:J15', userinfo["home_address"], cell_format_bv)
            work_sheets.merge_range('M15:O15', userinfo["subsidy_customer_ratio"], cell_format_bv)
            work_sheets.merge_range('E16:F16', userinfo["gjj_number"], cell_format_bv)
            work_sheets.merge_range('I16:J16', userinfo["corporation_number"], cell_format_bv)
            work_sheets.merge_range('M16:O16', userinfo["base_number"], cell_format_bv)
            work_sheets.merge_range('E17:F17', userinfo["balance"], cell_format_bv)
            work_sheets.merge_range('I17:J17', userinfo["corporation_name"], cell_format_bv)
            work_sheets.merge_range('M17:O17', userinfo["last_pay_date"], cell_format_bv)
            work_sheets.merge_range('E18:F18', userinfo["fund_balance"], cell_format_bv)
            work_sheets.merge_range('I18:J18', userinfo["monthly_corporation_income"], cell_format_bv)
            work_sheets.merge_range('M18:O18', userinfo["begin_date"], cell_format_bv)
        else:
            work_sheets.merge_range('E10:F10', '无信息', cell_format_bv)

        #缴存明细
        work_sheets.merge_range('C19:O19', '缴存明细', merge_format_b)
        bills = result_s["data"]["bill_record"]
        bill_2014 = 1
        if len(bills) !=  0:
            work_sheets.write('C20', '缴存时间', cell_format_bc)
            work_sheets.write('D20', '缴存年月', cell_format_bc)
            work_sheets.write('E20', '出账', cell_format_bc)
            work_sheets.write('F20', '补贴出账', cell_format_bc)
            work_sheets.write('G20', '入账', cell_format_bc)
            work_sheets.write('H20', '补贴入账', cell_format_bc)
            work_sheets.write('I20', '余额', cell_format_bc)
            work_sheets.write('J20', '缴存公司名称', cell_format_bc)
            work_sheets.write('K20', '公司缴存金额', cell_format_bc)
            work_sheets.write('L20', '个人缴存金额', cell_format_bc)
            work_sheets.write('M20', '公司缴存比例', cell_format_bc)
            work_sheets.write('N20', '个人缴存比例', cell_format_bc)
            work_sheets.write('O20', '补缴', cell_format_bc)
            for bills_len in range(len(bills)):
                bill_2014  = bill_2014+1
                work_sheets.write(20 + bills_len, 2, bills[bills_len]["deal_time"], cell_format_bv)
                work_sheets.write(20 + bills_len, 3, bills[bills_len]["month"], cell_format_bv)
                work_sheets.write(20 + bills_len, 4, bills[bills_len]["outcome"], cell_format_bv)
                work_sheets.write(20 + bills_len, 5, bills[bills_len]["subsidy_outcome"], cell_format_bv)
                work_sheets.write(20 + bills_len, 6, bills[bills_len]["income"], cell_format_bv)
                work_sheets.write(20 + bills_len, 7, bills[bills_len]["subsidy_income"], cell_format_bv)
                work_sheets.write(20 + bills_len, 8, bills[bills_len]["balance"], cell_format_bv)
                work_sheets.write(20 + bills_len, 9, bills[bills_len]["corporation_name"], cell_format_bv)
                work_sheets.write(20 + bills_len, 10, bills[bills_len]["corporation_income"], cell_format_bv)
                work_sheets.write(20 + bills_len, 11, bills[bills_len]["customer_income"], cell_format_bv)
                work_sheets.write(20 + bills_len, 12, bills[bills_len]["corporation_ratio"], cell_format_bv)
                work_sheets.write(20 + bills_len, 13, bills[bills_len]["customer_ratio"], cell_format_bv)
                work_sheets.write(20 + bills_len, 14, bills[bills_len]["additional_income"], cell_format_bv)
        else:
            work_sheets.write(19 , 2, '没有缴存记录', cell_format_bv)


        #贷款信息
        work_sheets.merge_range('C'+str(20+bill_2014)+':O'+str(20+bill_2014), '贷款信息', merge_format_b)
        loan = result_s["data"]["loan_info"]
        work_sheets.merge_range('C' + str(21 + bill_2014) + ':D' + str(21 + bill_2014), '贷款人姓名', cell_format_b)
        work_sheets.merge_range('G'+str(21+bill_2014)+':H'+str(21+bill_2014), '贷款开始时间', cell_format_b)
        work_sheets.merge_range('K'+str(21+bill_2014)+':L'+str(21+bill_2014), '第二还款人姓名', cell_format_b)
        work_sheets.merge_range('C'+str(22+bill_2014)+':D'+str(22+bill_2014), '贷款 - 联系手机', cell_format_b)
        work_sheets.merge_range('G'+str(22+bill_2014)+':H'+str(22+bill_2014), '贷款结束日期', cell_format_b)
        work_sheets.merge_range('K'+str(22+bill_2014)+':L'+str(22+bill_2014), '第二还款人身份证', cell_format_b)
        work_sheets.merge_range('C'+str(23+bill_2014)+':D'+str(23+bill_2014), '贷款状态', cell_format_b)
        work_sheets.merge_range('G'+str(23+bill_2014)+':H'+str(23+bill_2014), '还款方式', cell_format_b)
        work_sheets.merge_range('K'+str(23+bill_2014)+':L'+str(23+bill_2014), '第二还款人手机', cell_format_b)
        work_sheets.merge_range('C'+str(24+bill_2014)+':D'+str(24+bill_2014), '承办银行', cell_format_b)
        work_sheets.merge_range('G'+str(24+bill_2014)+':H'+str(24+bill_2014), '每月还款日', cell_format_b)
        work_sheets.merge_range('K'+str(24+bill_2014)+':L'+str(24+bill_2014), '第二还款人工作单位', cell_format_b)
        work_sheets.merge_range('C'+str(25+bill_2014)+':D'+str(25+bill_2014), '贷款类型', cell_format_b)
        work_sheets.merge_range('G'+str(25+bill_2014)+':H'+str(25+bill_2014), '扣款账号', cell_format_b)
        work_sheets.merge_range('K'+str(25+bill_2014)+':L'+str(25+bill_2014), '贷款余额', cell_format_b)
        work_sheets.merge_range('C'+str(26+bill_2014)+':D'+str(26+bill_2014), '贷款人身份证', cell_format_b)
        work_sheets.merge_range('G'+str(26+bill_2014)+':H'+str(26+bill_2014), '扣款银行账号姓名', cell_format_b)
        work_sheets.merge_range('K'+str(26+bill_2014)+':L'+str(26+bill_2014), '剩余期数', cell_format_b)
        work_sheets.merge_range('C'+str(27+bill_2014)+':D'+str(27+bill_2014), '当前贷款购房地址', cell_format_b)
        work_sheets.merge_range('G'+str(27+bill_2014)+':H'+str(27+bill_2014), '贷款利率', cell_format_b)
        work_sheets.merge_range('K'+str(27+bill_2014)+':L'+str(27+bill_2014), '最后还款日期', cell_format_b)
        work_sheets.merge_range('C'+str(28+bill_2014)+':D'+str(28+bill_2014), '通讯地址', cell_format_b)
        work_sheets.merge_range('G'+str(28+bill_2014)+':H'+str(28+bill_2014), '罚息利率', cell_format_b)
        work_sheets.merge_range('K'+str(28+bill_2014)+':L'+str(28+bill_2014), '逾期本金', cell_format_b)
        work_sheets.merge_range('C'+str(29+bill_2014)+':D'+str(29+bill_2014), '贷款合同号', cell_format_b)
        work_sheets.merge_range('G'+str(29+bill_2014)+':H'+str(29+bill_2014), '商业贷款合同编号', cell_format_b)
        work_sheets.merge_range('K'+str(29+bill_2014)+':L'+str(29+bill_2014), '逾期利息', cell_format_b)
        work_sheets.merge_range('C'+str(30+bill_2014)+':D'+str(30+bill_2014), '贷款期数', cell_format_b)
        work_sheets.merge_range('G'+str(30+bill_2014)+':H'+str(30+bill_2014), '商业贷款银行', cell_format_b)
        work_sheets.merge_range('K'+str(30+bill_2014)+':L'+str(30+bill_2014), '逾期罚息', cell_format_b)
        work_sheets.merge_range('C'+str(31+bill_2014)+':D'+str(31+bill_2014), '贷款金额', cell_format_b)
        work_sheets.merge_range('G'+str(31+bill_2014)+':H'+str(31+bill_2014), '商业贷款金额', cell_format_b)
        work_sheets.merge_range('K'+str(31+bill_2014)+':L'+str(31+bill_2014), '逾期天数', cell_format_b)
        work_sheets.merge_range('C'+str(32+bill_2014)+':D'+str(32+bill_2014), '月还款额度', cell_format_b)
        work_sheets.merge_range('G'+str(32+bill_2014)+':H'+str(32+bill_2014), '第二还款人银行账号', cell_format_b)
        work_sheets.merge_range('K'+str(32+bill_2014)+':L'+str(32+bill_2014), '-----', cell_format_b)
        if len(loan) == 0:
            work_sheets.merge_range('E' + str(21 + bill_2014) + ':F' + str(21 + bill_2014), '没有贷款记录', cell_format_bv)
        else:
            loan = result_s["data"]["loan_info"][0]
            work_sheets.merge_range('E'+str(21+bill_2014)+':F'+str(21+bill_2014), loan["name"], cell_format_bv)
            work_sheets.merge_range('I'+str(21+bill_2014)+':J'+str(21+bill_2014), loan["start_date"], cell_format_bv)
            work_sheets.merge_range('M'+str(21+bill_2014)+':O'+str(21+bill_2014), loan["second_bank_account_name"], cell_format_bv)
            work_sheets.merge_range('E'+str(22+bill_2014)+':F'+str(22+bill_2014), loan["phone"], cell_format_bv)
            work_sheets.merge_range('I'+str(22+bill_2014)+':J'+str(22+bill_2014), loan["end_date"], cell_format_bv)
            work_sheets.merge_range('M'+str(22+bill_2014)+':O'+str(22+bill_2014), loan["second_id_card"], cell_format_bv)
            work_sheets.merge_range('E'+str(23+bill_2014)+':F'+str(23+bill_2014), loan["status"], cell_format_bv)
            work_sheets.merge_range('I'+str(23+bill_2014)+':J'+str(23+bill_2014), loan["repay_type"], cell_format_bv)
            work_sheets.merge_range('M'+str(23+bill_2014)+':O'+str(23+bill_2014), loan["second_phone"], cell_format_bv)
            work_sheets.merge_range('E'+str(24+bill_2014)+':F'+str(24+bill_2014), loan["bank"], cell_format_bv)
            work_sheets.merge_range('I'+str(24+bill_2014)+':J'+str(24+bill_2014), loan["deduct_day"], cell_format_bv)
            work_sheets.merge_range('M'+str(24+bill_2014)+':O'+str(24+bill_2014), loan["second_corporation_name"], cell_format_bv)
            work_sheets.merge_range('E'+str(25+bill_2014)+':F'+str(25+bill_2014), loan["loan_type"], cell_format_bv)
            work_sheets.merge_range('I'+str(25+bill_2014)+':J'+str(25+bill_2014), loan["bank_account"], cell_format_bv)
            work_sheets.merge_range('M'+str(25+bill_2014)+':O'+str(25+bill_2014), loan["remain_amount"], cell_format_bv)
            work_sheets.merge_range('E'+str(26+bill_2014)+':F'+str(26+bill_2014), loan["id_card"], cell_format_bv)
            work_sheets.merge_range('I'+str(26+bill_2014)+':J'+str(26+bill_2014), loan["bank_account_name"], cell_format_bv)
            work_sheets.merge_range('M'+str(26+bill_2014)+':O'+str(26+bill_2014), loan["remain_periods"], cell_format_bv)
            work_sheets.merge_range('E'+str(27+bill_2014)+':F'+str(27+bill_2014), loan["house_address"], cell_format_bv)
            work_sheets.merge_range('I'+str(27+bill_2014)+':J'+str(27+bill_2014), loan["loan_interest_percent"], cell_format_bv)
            work_sheets.merge_range('M'+str(27+bill_2014)+':O'+str(27+bill_2014), loan["last_repay_date"], cell_format_bv)
            work_sheets.merge_range('E'+str(28+bill_2014)+':F'+str(28+bill_2014), loan["mailing_address"], cell_format_bv)
            work_sheets.merge_range('I'+str(28+bill_2014)+':J'+str(28+bill_2014), loan["penalty_interest_percent"], cell_format_bv)
            work_sheets.merge_range('M'+str(28+bill_2014)+':O'+str(28+bill_2014), loan["overdue_capital"], cell_format_bv)
            work_sheets.merge_range('E'+str(29+bill_2014)+':F'+str(29+bill_2014), loan["contract_number"], cell_format_bv)
            work_sheets.merge_range('I'+str(29+bill_2014)+':J'+str(29+bill_2014), loan["commercial_contract_number"], cell_format_bv)
            work_sheets.merge_range('M'+str(29+bill_2014)+':O'+str(29+bill_2014), loan["overdue_interest"], cell_format_bv)
            work_sheets.merge_range('E'+str(30+bill_2014)+':F'+str(30+bill_2014), loan["periods"], cell_format_bv)
            work_sheets.merge_range('I'+str(30+bill_2014)+':J'+str(30+bill_2014), loan["commercial_bank"], cell_format_bv)
            work_sheets.merge_range('M'+str(30+bill_2014)+':O'+str(30+bill_2014), loan["overdue_penalty"], cell_format_bv)
            work_sheets.merge_range('E'+str(31+bill_2014)+':F'+str(31+bill_2014), loan["loan_amount"], cell_format_bv)
            work_sheets.merge_range('I'+str(31+bill_2014)+':J'+str(31+bill_2014), loan["commercial_amount"], cell_format_bv)
            work_sheets.merge_range('M'+str(31+bill_2014)+':O'+str(31+bill_2014), loan["overdue_days"], cell_format_bv)
            work_sheets.merge_range('E'+str(32+bill_2014)+':F'+str(32+bill_2014), loan["monthly_repay_amount"], cell_format_bv)
            work_sheets.merge_range('I'+str(32+bill_2014)+':J'+str(32+bill_2014), loan["second_bank_account"], cell_format_bv)
            work_sheets.merge_range('M'+str(32+bill_2014)+':O'+str(32+bill_2014), '-----', cell_format_bv)


        #还款明细
        work_sheets.merge_range('C' + str(33 + bill_2014) + ':O' + str(33 + bill_2014), '还款明细', merge_format_b)
        loan_repay_record_v = result_s["data"]["loan_repay_record"]
        repay_num = 1
        if len(loan_repay_record_v) == 0:
            work_sheets.write(33 + bill_2014, 2, '没有还款信息', cell_format_bv)
        else:
            work_sheets.merge_range('C' + str(34 + bill_2014) + ':D' + str(34 + bill_2014), '还款日期', cell_format_bc)
            work_sheets.merge_range('E' + str(34 + bill_2014) + ':F' + str(34 + bill_2014), '记账日期', cell_format_bc)
            work_sheets.merge_range('G' + str(34 + bill_2014) + ':H' + str(34 + bill_2014), '还款金额', cell_format_bc)
            work_sheets.merge_range('I' + str(34 + bill_2014) + ':J' + str(34 + bill_2014), '还款本金', cell_format_bc)
            work_sheets.merge_range('K' + str(34 + bill_2014) + ':L' + str(34 + bill_2014), '还款利息', cell_format_bc)
            work_sheets.merge_range('M' + str(34 + bill_2014) + ':N' + str(34 + bill_2014), '还款罚息', cell_format_bc)
            work_sheets.write('O'+ str(34 + bill_2014) , '贷款合同号', cell_format_bc)
            for repay_v in range(len(loan_repay_record_v)):
                repay_num = repay_num+1
                work_sheets.merge_range('C' + str(35 + bill_2014 + repay_v) + ':D' + str(35 + bill_2014 + repay_v),loan_repay_record_v[repay_v]['repay_date'], cell_format_bv)
                work_sheets.merge_range('E' + str(35 + bill_2014 + repay_v) + ':F' + str(35 + bill_2014 + repay_v),loan_repay_record_v[repay_v]['accounting_date'], cell_format_bv)
                work_sheets.merge_range('G' + str(35 + bill_2014 + repay_v) + ':H' + str(35 + bill_2014 + repay_v),loan_repay_record_v[repay_v]['repay_amount'], cell_format_bv)
                work_sheets.merge_range('I' + str(35 + bill_2014 + repay_v) + ':J' + str(35 + bill_2014 + repay_v),loan_repay_record_v[repay_v]['repay_capital'], cell_format_bv)
                work_sheets.merge_range('K' + str(35 + bill_2014 + repay_v) + ':L' + str(35 + bill_2014 + repay_v),loan_repay_record_v[repay_v]['repay_interest'], cell_format_bv)
                work_sheets.merge_range('M' + str(35 + bill_2014 + repay_v) + ':N' + str(35 + bill_2014 + repay_v),loan_repay_record_v[repay_v]['repay_penalty'], cell_format_bv)
                work_sheets.write('O' + str(35 + bill_2014 + repay_v), loan_repay_record_v[repay_v]['contract_number'],cell_format_bv)
        print("个人报告生成成功............")
    else:
        print(".....没有记录....")

    for ls in range(35 +bill_2014+ repay_num):
        work_sheets.write('B' + str(ls + 2), '', border_l)
        work_sheets.write('Q' + str(ls + 2), '', border_l)

    for ts in range(15):
        work_sheets.write(1, ts + 1, '', border_t)
        work_sheets.write(35 +bill_2014+ repay_num, ts + 1, '', border_t)

        work_sheets.write('B2', '', border_h2)


def reports(total,row):
    # 生成报告
    print("详情信息收集完毕，创建报告中......")
    merge_format_t = workbook.add_format({'font_size': 20, 'align': 'center', 'valign': 'vcenter', 'font_name': '宋体'})
    merge_format_b = workbook.add_format(
        {'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'border': 2, 'font_name': '宋体'})
    merge_format_r = workbook.add_format(
        {'font_size': 10, 'align': 'right', 'valign': 'vcenter', 'border': 1, 'font_name': '宋体'})

    cell_format_e = workbook.add_format(
        {'font_size': 10, 'font_color': 'red', 'font_name': '宋体', 'align': 'justify', 'border': 1})
    cell_format = workbook.add_format(
        {'font_size': 10, 'font_color': 'green', 'font_name': '宋体', 'align': 'justify', 'border': 1})
    cell_format_b = workbook.add_format(
        {'font_size': 10, 'font_color': 'black', 'font_name': '宋体', 'align': 'justify', 'border': 1})
    cell_format_yh = workbook.add_format({'font_size': 11, 'font_color': 'black', 'font_name': '微软雅黑', 'align': 'left'})

    border_l = workbook.add_format({'left': 2})
    border_t = workbook.add_format({'top': 2})
    border_h2 = workbook.add_format({'top': 2, 'left': 2})

    for lc in range(74):
        work_sheet1.write('H' + str(lc + 2), '', border_l)
        work_sheet1.write('P' + str(lc + 2), '', border_l)

    for lc1 in range(5):
        work_sheet1.write('O' + str(lc1 + 13), '', border_l)

    for tc in range(8):
        work_sheet1.write(1, tc + 7, '', border_t)
        work_sheet1.write(75, tc + 7, '', border_t)

    for tcl in range(5):
        work_sheet1.write(17, tcl + 9, '', border_t)

    work_sheet1.write('H2', '', border_h2)

    work_sheet1.set_column('I:K', 14)
    work_sheet1.set_column('N:N', 14)
    work_sheet1.merge_range('I5:N4', '测试报告', merge_format_t)
    work_sheet1.merge_range('I12:N12', '测试概览', merge_format_b)
    work_sheet1.merge_range('I13:I17', '公积金API', merge_format_b)
    work_sheet1.merge_range('K16:N16', now, merge_format_r)
    work_sheet1.merge_range('K17:N17', '任欢欢', merge_format_r)
    work_sheet1.write('J14', '反案例数', cell_format_e)
    work_sheet1.write('J13', '正案例数', cell_format)
    work_sheet1.write('J15', '回归条数', cell_format_b)
    work_sheet1.write('J16', '执行日期', cell_format_b)
    work_sheet1.write('J17', '执行人', cell_format_b)
    work_sheet1.write_formula('K14', '=COUNTIF(\'公积金-案例详情\'!D2:D' + str(total) + ',"反案例")', cell_format_e)
    work_sheet1.write_formula('K13','=COUNTIF(\'公积金-案例详情\'!D2:D' + str(total) + ',"正案例")', cell_format)
    work_sheet1.write_formula('K15', '=COUNTIF(\'公积金回归明细\'!D2:D' + str(total) + ',"跑数据")', cell_format_b)
    work_sheet1.write('L13', '通过数', cell_format_b)
    work_sheet1.write('L14', '失败数', cell_format_e)
    work_sheet1.write('L15', '用例总数', cell_format_b)
    work_sheet1.write_formula('M13', '=COUNTIF(\'公积金回归明细\'!R2:R' + str(total) + ',"Pass")+COUNTIF(\'公积金-案例详情\'!H2:H' + str(total) + ',"Pass")', cell_format_b)
    work_sheet1.write_formula('M14', '=COUNTIF(\'公积金回归明细\'!R2:R' + str(total) + ',"Fail")+COUNTIF(\'公积金-案例详情\'!H2:H' + str(total) + ',"Fail")', cell_format_e)
    work_sheet1.write_formula('M15', '=COUNTIF(\'公积金回归明细\'!A2:A' + str(total) + ',"<>")+COUNTIF(\'公积金-案例详情\'!A2:A' + str(total) + ',"<>")', cell_format_b)
    work_sheet1.write_formula('N13', '=M13/M15', cell_format_b)
    work_sheet1.write_formula('N14', '=M14/M15', cell_format_e)
    work_sheet1.write_formula('N15', '=M15/M15', cell_format_b)

    work_sheet1.conditional_format('N13:N15', {'type': 'data_bar', 'bar_color': '#63C384'})

    work_sheet1.write('I22', '综合分析：', cell_format_yh)
    print("报告分析中............")
    print("正在分析案例分布情况..........")
    #生成案例分布环图
    # 设置图表
    chart1 = workbook.add_chart({'type': 'doughnut'})

    # 选择数据区域
    chart1.add_series({
        'name': '测试报告',
        'categories': '=\'公积金-报告\'!$J$13:$J$15',
        'values': '=\'公积金-报告\'!$K$13:$K$15',
    })

    # 给图表命名
    chart1.set_title({'name': '案例分布'})
    chart1.set_rotation(90)

    work_sheet1.insert_chart('I24', chart1, {'x_offset': 25, 'y_offset': 10})
    print("案例分布情况分析完毕..........")

    print("正在分析接口耗时情况..........")
    #生成接口耗时点图
    work_sheet1.write('H185','获取城市公积金列表', cell_format_b)
    work_sheet1.write_formula('I185', '=AVERAGEIF(\'公积金-案例详情\'!E2:E' + str(total) + ',"获取城市公积金列表",\'公积金-案例详情\'!I2:I' + str(total)+')', cell_format_b)
    work_sheet1.write_formula('J185', '{=MAX(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="获取城市公积金列表",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)
    work_sheet1.write_formula('K185', '{=MIN(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="获取城市公积金列表",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)

    work_sheet1.write('H186','根据 areaCode 获取公积金登录信息', cell_format_b)
    work_sheet1.write_formula('I186', '=AVERAGEIF(\'公积金-案例详情\'!E2:E' + str(total) + ',"根据 areaCode 获取公积金登录信息",\'公积金-案例详情\'!I2:I' + str(total)+')', cell_format_b)
    work_sheet1.write_formula('J186', '{=MAX(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="根据 areaCode 获取公积金登录信息",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)
    work_sheet1.write_formula('K186', '{=MIN(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="根据 areaCode 获取公积金登录信息",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)

    work_sheet1.write('H187','创建订单', cell_format_b)
    work_sheet1.write_formula('I187', '=AVERAGEIF(\'公积金-案例详情\'!E2:E' + str(total) + ',"创建订单",\'公积金-案例详情\'!I2:I' + str(total)+')', cell_format_b)
    work_sheet1.write_formula('J187', '{=MAX(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="创建订单",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)
    work_sheet1.write_formula('K187', '{=MIN(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="创建订单",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)

    work_sheet1.write('H188','订单执行状态查询', cell_format_b)
    work_sheet1.write_formula('I188', '=AVERAGEIF(\'公积金-案例详情\'!E2:E' + str(total) + ',"订单执行状态查询",\'公积金-案例详情\'!I2:I' + str(total)+')', cell_format_b)
    work_sheet1.write_formula('J188', '{=MAX(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="订单执行状态查询",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)
    work_sheet1.write_formula('K188', '{=MIN(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="订单执行状态查询",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)

    work_sheet1.write('H189','获取公积金信息', cell_format_b)
    work_sheet1.write_formula('I189', '=AVERAGEIF(\'公积金-案例详情\'!E2:E' + str(total) + ',"获取公积金信息",\'公积金-案例详情\'!I2:I' + str(total)+')', cell_format_b)
    work_sheet1.write_formula('J189', '{=MAX(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="获取公积金信息",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)
    work_sheet1.write_formula('K189', '{=MIN(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="获取公积金信息",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)

    work_sheet1.write('H190','根据订单号、年份分页查询公积金缴纳信息', cell_format_b)
    work_sheet1.write_formula('I190', '=AVERAGEIF(\'公积金-案例详情\'!E2:E' + str(total) + ',"根据订单号、年份分页查询公积金缴纳信息",\'公积金-案例详情\'!I2:I' + str(total)+')', cell_format_b)
    work_sheet1.write_formula('J190', '{=MAX(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="根据订单号、年份分页查询公积金缴纳信息",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)
    work_sheet1.write_formula('K190', '{=MIN(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="根据订单号、年份分页查询公积金缴纳信息",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)

    work_sheet1.write('H191','根据订单号查询公积金账户信息', cell_format_b)
    work_sheet1.write_formula('I191', '=AVERAGEIF(\'公积金-案例详情\'!E2:E' + str(total) + ',"根据订单号查询公积金账户信息",\'公积金-案例详情\'!I2:I' + str(total)+')', cell_format_b)
    work_sheet1.write_formula('J191', '{=MAX(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="根据订单号查询公积金账户信息",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)
    work_sheet1.write_formula('K191', '{=MIN(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="根据订单号查询公积金账户信息",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)

    work_sheet1.write('H192','根据订单号查询公积金贷款信息', cell_format_b)
    work_sheet1.write_formula('I192', '=AVERAGEIF(\'公积金-案例详情\'!E2:E' + str(total) + ',"根据订单号查询公积金贷款信息",\'公积金-案例详情\'!I2:I' + str(total)+')', cell_format_b)
    work_sheet1.write_formula('J192', '{=MAX(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="根据订单号查询公积金贷款信息",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)
    work_sheet1.write_formula('K192', '{=MIN(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="根据订单号查询公积金贷款信息",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)

    work_sheet1.write('H193','根据订单号查询公积金还款信息', cell_format_b)
    work_sheet1.write_formula('I193', '=AVERAGEIF(\'公积金-案例详情\'!E2:E' + str(total) + ',"根据订单号查询公积金还款信息",\'公积金-案例详情\'!I2:I' + str(total)+')', cell_format_b)
    work_sheet1.write_formula('J193', '{=MAX(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="根据订单号查询公积金还款信息",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)
    work_sheet1.write_formula('K193', '{=MIN(IF(\'公积金-案例详情\'!E2:E' + str(total) + '="根据订单号查询公积金还款信息",\'公积金-案例详情\'!I2:I' + str(total)+'))}', cell_format_b)

    chart3 = workbook.add_chart({'type': 'column'})

    # 选择数据区域
    chart3.add_series({
        'name': '最大耗时（μs）',
        'categories': '=$H$185:$H$193',
        'values': '=$J$185:$J$193',
    })
    chart3.add_series({
        'name': '平均耗时（μs）',
        'categories': '=$H$185:$H$193',
        'values': '=$I$185:$I$193',
    })

    chart3.add_series({
        'name': '最小耗时（μs）',
        'categories': '=$H$185:$H$193',
        'values': '=$K$185:$K$193',
    })

    # 给图表命名
    chart3.set_title({'name': '接口耗时'})


    work_sheet1.insert_chart('I58', chart3, {'x_offset': 25, 'y_offset': 10})


    print("接口耗时情况分析完毕..........")


    print("正在分析爬取耗时情况..........")
    #生成爬取耗时分布图
    # 设置图表
    chart2 = workbook.add_chart({'type': 'column'})

    # 选择数据区域
    chart2.add_series({
        'name': '耗时（s）',
        'categories': '=\'公积金回归明细\'!$G$2:$G$'+str(row+1),
        'values': '=\'公积金回归明细\'!$S$2:$S$'+str(row+1),
    })

    # 给图表命名
    chart2.set_title({'name': '地区爬取耗时'})


    work_sheet1.insert_chart('I41', chart2, {'x_offset': 25, 'y_offset': 10})

    work_sheet1.hide_gridlines(2)

    print("爬取耗时情况分析完毕..........")

    print("报告生成成功......")

def close_workbook():
    workbook.close()