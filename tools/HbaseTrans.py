import csv
import codecs
import re
from common.table_model import hbase_mode
import readconfig

print("----------开始执行------------")
localRead = readconfig.ReadConfig()

table_re_path = localRead.get_address('hbase_re')
file = codecs.open(table_re_path,'w','utf-8')

print("获取内容中......")
table_path = localRead.get_address('hbase')


print("内容解析中......")
with codecs.open(table_path,'r','utf-8') as f:
    reader = csv.reader(f)
    for i,rows in enumerate(reader):
        if i>=0:
            #.....................................
            #生成时候注意改表名和rowkey！！！！
            #......................................
            table_name = 'CREDIT_MINING:Product_CardInfo'#表名-------------
            clume_name = re.sub(r'.*info:',"",rows[0])
            rowkey = '100002f52600ab40'#rowkey----------------
            if (len(rows)>3):
                value = re.sub(r'.*value=',"",rows[2])+','+re.sub(r' .*$',"",rows[3])
            else:
                value = re.sub(r' .*$', "", re.sub(r'.*value=', "", rows[2]))
            hbase_re = hbase_mode(table_name,rowkey,clume_name,value)
            print("第"+str(i+1)+"行解析成功...")
            file.write(hbase_re)

print("解析结束")
f.close()
file.close()
print("----------文件已生成------------")