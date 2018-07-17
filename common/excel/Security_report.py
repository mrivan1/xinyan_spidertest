#encoding='gb18030'
import xlsxwriter
from common.rand import tradedate

# 新建结果集
now = tradedate()
filename = '测试报告' + now + '.xlsx'
workbook = xlsxwriter.Workbook('D:/DevCode/data/spider/report/Security/'+filename)

# 创建sheet
work_sheet = workbook.add_worksheet('社保-回归')

# 结果包含的信息
row = ['ID', '案例描述', '案例ID', '环境', '订单号','地区', '社保信息','是否通过','失败原因']
for s in range(len(row)):
    work_sheet.write(0, s, row[s])
work_sheet.freeze_panes(1, 0)

def report(rownum,ID,DESCRIBE,CASE_ID,ENV,tradeno,area,userinfo,result,err):
    row1 = [ID,DESCRIBE,CASE_ID,ENV,tradeno,area,userinfo,result,err]
    for k in range(len(row1)):
        work_sheet.write(rownum, k, str(row1[k]))


#创建公积金个人报告
def create_ownreport(areaname,result_s):
    print("生成个人报告中............")
    if result_s !='':
        work_sheets = workbook.add_worksheet(areaname + '报告详情')
        work_sheets.hide_gridlines(2)
        merge_format_t = workbook.add_format({'font_size': 20, 'align': 'center', 'valign': 'vcenter', 'font_name': '宋体'})
        merge_format_b = workbook.add_format({'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'border': 2, 'font_name': '宋体','fg_color':'#CD2626'})
        cell_format_b = workbook.add_format({'font_size': 10, 'font_color': 'black', 'font_name': '宋体', 'align': 'justify', 'border': 1,'fg_color':'#7CCD7C'})
        cell_format_bv = workbook.add_format({'font_size': 10, 'font_color': 'black', 'font_name': '宋体', 'align': 'right', 'border': 1})
        cell_format_bc = workbook.add_format({'font_size': 10, 'font_color': 'black', 'font_name': '宋体', 'align': 'justify', 'border': 1,'fg_color':'#7CCD7C'})
        border_l = workbook.add_format({'left': 2})
        border_t = workbook.add_format({'top': 2})
        border_h2 = workbook.add_format({'top': 2, 'left': 2})

        work_sheets.merge_range('C3:P6', '社保个人报告', merge_format_t)
        work_sheets.merge_range('C7:D7', '系统 id：', cell_format_bc)
        work_sheets.merge_range('E7:F7', str(result_s["data"]["task_id"]), cell_format_bv)
        work_sheets.merge_range('G7:H7', '地区编码：', cell_format_bc)
        work_sheets.merge_range('I7:K7', str(result_s["data"]["area_code"]), cell_format_bv)
        work_sheets.merge_range('L7:M7', '所属城市名称：', cell_format_bc)
        work_sheets.merge_range('N7:P7', str(result_s["data"]["city"]), cell_format_bv)
        #基本信息
        work_sheets.merge_range('C9:P9', '基本信息', merge_format_b)
        userinfo = result_s["data"]["base_info"]
        work_sheets.merge_range('C10:D10', '真实姓名', cell_format_b)
        work_sheets.merge_range('G10:H10', '用户信息 Id', cell_format_b)
        work_sheets.merge_range('K10:L10', '社保帐号', cell_format_b)
        work_sheets.merge_range('C11:D11', '个人编号', cell_format_b)
        work_sheets.merge_range('G11:H11', '民族', cell_format_b)
        work_sheets.merge_range('K11:L11', '证件类型', cell_format_b)
        work_sheets.merge_range('C12:D12', '证件号码', cell_format_b)
        work_sheets.merge_range('G12:H12', '家庭住址', cell_format_b)
        work_sheets.merge_range('K12:L12', '缴存基数', cell_format_b)
        work_sheets.merge_range('C13:D13', '开户日期', cell_format_b)
        work_sheets.merge_range('G13:H13', '出生日期', cell_format_b)
        work_sheets.merge_range('K13:L13', '首次参保时间', cell_format_b)
        work_sheets.merge_range('C14:D14', '户口性质', cell_format_b)
        work_sheets.merge_range('G14:H14', '参保单位', cell_format_b)
        work_sheets.merge_range('K14:L14', '参保单位编号', cell_format_b)
        work_sheets.merge_range('C15:D15', '最新缴存日期', cell_format_b)
        work_sheets.merge_range('G15:H15', '缴存状态', cell_format_b)
        work_sheets.merge_range('K15:L15', '人员状态', cell_format_b)
        work_sheets.merge_range('C16:D16', '手机号', cell_format_b)
        work_sheets.merge_range('G16:H16', '性别', cell_format_b)
        work_sheets.merge_range('K16:L16', '单位类型', cell_format_b)
        work_sheets.merge_range('C17:D17', '参加工作时间', cell_format_b)
        work_sheets.merge_range('G17:H17', '工伤保险', cell_format_b)
        work_sheets.merge_range('K17:L17', '失业保险', cell_format_b)
        work_sheets.merge_range('C18:D18', '医疗保险', cell_format_b)
        work_sheets.merge_range('G18:H18', '医疗保险余额', cell_format_b)
        work_sheets.merge_range('K18:L18', '养老保险', cell_format_b)
        work_sheets.merge_range('C19:D19', '医疗保险', cell_format_b)
        work_sheets.merge_range('G19:H19', '医疗保险余额', cell_format_b)
        if len(userinfo) != 0:
            work_sheets.merge_range('E10:F10', userinfo["real_name"], cell_format_bv)
            work_sheets.merge_range('I10:J10', userinfo["user_info_id"], cell_format_bv)
            work_sheets.merge_range('M10:P10', userinfo["social_security_no"], cell_format_bv)
            work_sheets.merge_range('E11:F11', userinfo["personal_no"], cell_format_bv)
            work_sheets.merge_range('I11:J11', userinfo["nation"], cell_format_bv)
            work_sheets.merge_range('M11:P11', userinfo["id_type"], cell_format_bv)
            work_sheets.merge_range('E12:F12', userinfo["id_card"], cell_format_bv)
            work_sheets.merge_range('I12:J12', userinfo["address"], cell_format_bv)
            work_sheets.merge_range('M12:P12', userinfo["base_number"], cell_format_bv)
            work_sheets.merge_range('E13:F13', userinfo["begin_date"], cell_format_bv)
            work_sheets.merge_range('I13:J13', userinfo["birth_day"], cell_format_bv)
            work_sheets.merge_range('M13:P13', userinfo["first_insured_date"], cell_format_bv)
            work_sheets.merge_range('E14:F14', userinfo["household_registration"], cell_format_bv)
            work_sheets.merge_range('I14:J14', userinfo["insured_unit"], cell_format_bv)
            work_sheets.merge_range('M14:P14', userinfo["insured_unit_code"], cell_format_bv)
            work_sheets.merge_range('E15:F15', userinfo["last_pay_date"], cell_format_bv)
            work_sheets.merge_range('I15:J15', userinfo["pay_status"], cell_format_bv)
            work_sheets.merge_range('M15:P15', userinfo["personnel_status"], cell_format_bv)
            work_sheets.merge_range('E16:F16', userinfo["phone"], cell_format_bv)
            work_sheets.merge_range('I16:J16', userinfo["sex"], cell_format_bv)
            work_sheets.merge_range('M16:P16', userinfo["unit_type"], cell_format_bv)
            work_sheets.merge_range('E17:F17', userinfo["work_time"], cell_format_bv)
            work_sheets.merge_range('I17:J17', userinfo["industrial_insurance"], cell_format_bv)
            work_sheets.merge_range('M17:P17', userinfo["unemployment_insurance"], cell_format_bv)
            work_sheets.merge_range('E18:F18', userinfo["medical_insurance"], cell_format_bv)
            work_sheets.merge_range('I18:J18', userinfo["medical_insurance_balance"], cell_format_bv)
            work_sheets.merge_range('M18:P18', userinfo["endowment_insurance"], cell_format_bv)
            work_sheets.merge_range('E19:F19', userinfo["maternity_insurance"], cell_format_bv)
            work_sheets.merge_range('I19:J19', userinfo["fetch_time"], cell_format_bv)
        else:
            work_sheets.merge_range('E10:F10', '无信息', cell_format_bv)

        #保险种类
        work_sheets.merge_range('C20:P20', '保险种类', merge_format_b)
        bills = result_s["data"]["insurances"]
        bill_2014 = 1
        if len(bills) !=  0:
            work_sheets.write('C21', '缴纳基数', cell_format_bc)
            work_sheets.write('D21', '缴存公司名称', cell_format_bc)
            work_sheets.write('E21', '公司缴存比例', cell_format_bc)
            work_sheets.write('F21', '个人缴存比例', cell_format_bc)
            work_sheets.write('G21', '描述信息', cell_format_bc)
            work_sheets.write('H21', '首次参保时间', cell_format_bc)
            work_sheets.write('I21', '险种编号', cell_format_bc)
            work_sheets.write('J21', '参保状态', cell_format_bc)
            work_sheets.write('K21', '保险id', cell_format_bc)
            work_sheets.write('L21', '保险类型', cell_format_bc)
            work_sheets.write('M21', '公司缴存金额', cell_format_bc)
            work_sheets.write('N21', '个人缴存金额', cell_format_bc)
            work_sheets.write('O21', '缴存月数', cell_format_bc)
            work_sheets.write('P21', '---', cell_format_bc)
            for bills_len in range(len(bills)):
                bill_2014  = bill_2014+1
                work_sheets.write(21 + bills_len, 2, bills[bills_len]["base_number"], cell_format_bv)
                work_sheets.write(21 + bills_len, 3, bills[bills_len]["corporation_name"], cell_format_bv)
                work_sheets.write(21 + bills_len, 4, bills[bills_len]["corporation_scale"], cell_format_bv)
                work_sheets.write(21 + bills_len, 5, bills[bills_len]["customer_scale"], cell_format_bv)
                work_sheets.write(21 + bills_len, 6, bills[bills_len]["description"], cell_format_bv)
                work_sheets.write(21 + bills_len, 7, bills[bills_len]["first_insured_date"], cell_format_bv)
                work_sheets.write(21 + bills_len, 8, bills[bills_len]["insurance_code"], cell_format_bv)
                work_sheets.write(21 + bills_len, 9, bills[bills_len]["insurance_status"], cell_format_bv)
                work_sheets.write(21 + bills_len, 10, bills[bills_len]["insurance_id"], cell_format_bv)
                work_sheets.write(21 + bills_len, 11, bills[bills_len]["insurance_type"], cell_format_bv)
                work_sheets.write(21 + bills_len, 12, bills[bills_len]["monthly_corporation_income"], cell_format_bv)
                work_sheets.write(21 + bills_len, 13, bills[bills_len]["monthly_customer_income"], cell_format_bv)
                work_sheets.write(21 + bills_len, 14, bills[bills_len]["total_months"], cell_format_bv)
                work_sheets.write(21 + bills_len, 15, '---', cell_format_bv)
        else:
            work_sheets.write(20 , 2, '没有记录', cell_format_bv)



        #保险缴存记录
        work_sheets.merge_range('C' + str(21 + bill_2014) + ':P' + str(21 + bill_2014), '保险缴存记录', merge_format_b)
        insurance_record_v = result_s["data"]["insurance_record"]
        repay_num = 1
        if len(insurance_record_v) == 0:
            work_sheets.write(22 + bill_2014, 2, '没有保险缴存记录信息', cell_format_bv)
        else:
            work_sheets.write('C'+str(22+bill_2014 ), '缴存总额', cell_format_bc)
            work_sheets.write('D'+str(22+bill_2014 ), '缴纳基数', cell_format_bc)
            work_sheets.write('E'+str(22+bill_2014 ), '缴存公司名称', cell_format_bc)
            work_sheets.write('F'+str(22+bill_2014 ), '公司缴存金额', cell_format_bc)
            work_sheets.write('G'+str(22+bill_2014 ), '公司缴存比例', cell_format_bc)
            work_sheets.write('H'+str(22+bill_2014 ), '个人缴存金额', cell_format_bc)
            work_sheets.write('I'+str(22+bill_2014 ), '个人缴存比例', cell_format_bc)
            work_sheets.write('J'+str(22+bill_2014 ), '发生时间', cell_format_bc)
            work_sheets.write('K'+str(22+bill_2014 ), '描述信息', cell_format_bc)
            work_sheets.write('L'+str(22+bill_2014 ), '所属保险', cell_format_bc)
            work_sheets.write('M'+str(22+bill_2014 ), '险种编号', cell_format_bc)
            work_sheets.write('N'+str(22+bill_2014 ), '缴存月份', cell_format_bc)
            work_sheets.write('O'+str(22+bill_2014 ), '缴存月份(end)', cell_format_bc)
            work_sheets.write('P'+str(22+bill_2014 ), '缴存状态标记', cell_format_bc)
            for repay_v in range(len(insurance_record_v)):
                repay_num = repay_num+1
                work_sheets.write(22 + bill_2014 + repay_v, 2, insurance_record_v[repay_v]["amount"], cell_format_bv)
                work_sheets.write(22 + bill_2014+ repay_v, 3, insurance_record_v[repay_v]["base_number"], cell_format_bv)
                work_sheets.write(22 + bill_2014+ repay_v, 4, insurance_record_v[repay_v]["corporation_name"], cell_format_bv)
                work_sheets.write(22 + bill_2014+ repay_v, 5, insurance_record_v[repay_v]["corporation_payment"], cell_format_bv)
                work_sheets.write(22 + bill_2014+ repay_v, 6, insurance_record_v[repay_v]["corporation_scale"], cell_format_bv)
                work_sheets.write(22 + bill_2014+ repay_v, 7, insurance_record_v[repay_v]["personal_payment"], cell_format_bv)
                work_sheets.write(22 + bill_2014+ repay_v, 8, insurance_record_v[repay_v]["customer_scale"], cell_format_bv)
                work_sheets.write(22 + bill_2014+ repay_v, 9, insurance_record_v[repay_v]["deal_time"], cell_format_bv)
                work_sheets.write(22 + bill_2014+ repay_v, 10, insurance_record_v[repay_v]["description"], cell_format_bv)
                work_sheets.write(22 + bill_2014+ repay_v, 11, insurance_record_v[repay_v]["insurance_type"], cell_format_bv)
                work_sheets.write(22 + bill_2014+ repay_v, 12, insurance_record_v[repay_v]["insurance_code"], cell_format_bv)
                work_sheets.write(22 + bill_2014+ repay_v, 13, insurance_record_v[repay_v]["month"], cell_format_bv)
                work_sheets.write(22 + bill_2014+ repay_v, 14, insurance_record_v[repay_v]["month_end"], cell_format_bv)
                work_sheets.write(22 + bill_2014+ repay_v, 15, insurance_record_v[repay_v]["status"], cell_format_bv)
            work_sheets.autofilter('M'+str(22+bill_2014)+':M'+str(22 + bill_2014+repay_num))


        #医保消费信息
        work_sheets.merge_range('C' + str(22 + bill_2014+repay_num) + ':P' + str(22 + bill_2014+repay_num), '医保消费信息', merge_format_b)
        medical_insurance_record_v = result_s["data"]["medical_insurance_record"]
        medical_num = 0
        if len(medical_insurance_record_v) == 0:
            work_sheets.write(22 + bill_2014+repay_num, 2, '没有医保消费信息', cell_format_bv)
        else:
            work_sheets.merge_range('C' + str(23 + bill_2014+repay_num) + ':E' + str(23 + bill_2014+repay_num), '医疗机构名称', cell_format_bc)
            work_sheets.merge_range('F' + str(23 + bill_2014+repay_num) + ':H' + str(23 + bill_2014+repay_num), '医疗机构类别', cell_format_bc)
            work_sheets.merge_range('I' + str(23 + bill_2014+repay_num) + ':K' + str(23 + bill_2014+repay_num), '医保结算时间', cell_format_bc)
            work_sheets.merge_range('L' + str(23 + bill_2014+repay_num) + ':P' + str(23 + bill_2014+repay_num), '结算金额', cell_format_bc)
            for medical_v in range(len(medical_insurance_record_v)):
                work_sheets.merge_range('C' + str(24 + bill_2014 + repay_num+medical_v) + ':E' + str(24 + bill_2014 + repay_num+medical_v),medical_insurance_record_v[medical_v]['organization_name'], cell_format_bv)
                work_sheets.merge_range('F' + str(24 + bill_2014 + repay_num+medical_v) + ':H' + str(24 + bill_2014 + repay_num+medical_v),medical_insurance_record_v[medical_v]['type'], cell_format_bv)
                work_sheets.merge_range('I' + str(24 + bill_2014 + repay_num+medical_v) + ':K' + str(24 + bill_2014 + repay_num+medical_v),medical_insurance_record_v[medical_v]['settlemen_time'], cell_format_bv)
                work_sheets.merge_range('L' + str(24 + bill_2014 + repay_num+medical_v) + ':P' + str(24 + bill_2014 + repay_num+medical_v),medical_insurance_record_v[medical_v]['money'], cell_format_bv)
                medical_num = medical_num + 1

        print("个人报告生成成功............")


    else:
        print(".....没有记录....")

    for ls in range(24 + bill_2014 + repay_num + medical_num):
        work_sheets.write('B' + str(ls + 2), '', border_l)
        work_sheets.write('R' + str(ls + 2), '', border_l)

    for ts in range(16):
        work_sheets.write(1, ts + 1, '', border_t)
        work_sheets.write(25 + bill_2014 + repay_num + medical_num, ts + 1, '', border_t)

        work_sheets.write('B2', '', border_h2)


def close_workbook():
    # 关闭文件
    workbook.close()