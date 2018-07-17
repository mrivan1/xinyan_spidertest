import json
def table_mode(x,y):
    table_j = '{"col_name": "'+x+'","lab_name": "'+y+'","created_at": "1510209233508","created_by": "string","dict_status": "NORMAL","id": 0,"table_id": 0,"updated_at": "1510209233508","updated_by": "string"}'
    return table_j

def member_mode(x,y,z):
    member_j = '{"box_type": "","cfg_id": 0,"column_family": "info","column_status": "NORMAL","column_type": "SYSTEM","created_at": "1510209233508","created_by": "string","id": 0,"table_name": "'+x+'","column_name": "'+y+'","out_name": "'+z+'","split_box": "","updated_at": "1510209233508","updated_by": "string"}'
    return member_j

def hbase_mode(x,y,z,v):
    hbase_j = "put'"+x+"','"+y+"','info:"+z+"','"+v+"'"+"\n"
    return hbase_j