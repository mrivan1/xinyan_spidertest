import csv
import codecs
from common.table_model import member_mode
import readconfig

print("----------开始执行------------")
localRead = readconfig.ReadConfig()

member_re_path = localRead.get_address('member_re')
file = codecs.open(member_re_path,'w','utf-8')

print("获取内容中......")
member_path = localRead.get_address('member')
with codecs.open(member_path,'r','utf-8') as g:
    lines = len(g.readlines())
g.close()
print("内容解析中......")
with codecs.open(member_path,'r','utf-8') as f:
    reader = csv.reader(f)
    for i,rows in enumerate(reader):
        if i>0:
            table_name = rows[0]
            column_name = rows[1]
            out_name = rows[2]
            if i != lines-1:
                member_re = member_mode(table_name, column_name,out_name) + ','
            else:
                member_re = member_mode(table_name, column_name,out_name)
            print("第"+str(i)+"行解析成功...")
            file.write(member_re)

print("解析结束")
f.close()
file.close()
print("----------文件已生成------------")


if __name__ == '__main__':
    pass