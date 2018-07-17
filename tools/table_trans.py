import csv
import codecs
from common.table_model import table_mode
import readconfig

print("----------开始执行------------")
localRead = readconfig.ReadConfig()

table_re_path = localRead.get_address('table_re')
file = codecs.open(table_re_path,'w','utf-8')

print("获取内容中......")
table_path = localRead.get_address('table')
with codecs.open(table_path,'r','utf-8') as g:
    lines = len(g.readlines())
g.close()

print("内容解析中......")
with codecs.open(table_path,'r','utf-8') as f:
    reader = csv.reader(f)
    for i,rows in enumerate(reader):
        if i>0:
            col_name = rows[0]
            lab_name = rows[1]
            if i != lines-1:
                table_re = table_mode(col_name, lab_name) + ','
            else:
                table_re = table_mode(col_name, lab_name)
            print("第"+str(i)+"行解析成功...")
            file.write(table_re)

print("解析结束")
f.close()
file.close()
print("----------文件已生成------------")