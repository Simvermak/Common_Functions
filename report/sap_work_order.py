import json
from datetime import datetime, timedelta
import requests
from w3lib.http import basic_auth_header
import pandas as pd
import common.database as database

'''
SAP系统工单取数
'''

# 根据需求所修改的参数
start_date = str(20230601)
end_date = str(20230627)
ORDERS = []  # 查询制定单据功能（可为空） 示例：[{"OBJ":"110000000856"}]

test_path = '/DEV/REST_PRODUCTION_ORDER_QUERY'
produce_path = '/PRD/REST_PRODUCTION_ORDER_QUERY'
root = 'https://voopoo-dq66ezwa.it-cpi010-rt.cpi.cn40.apps.platform.sapcloud.cn/http'

username = "sb-68ec7809-f61f-4e94-81aa-cbf4aa0aa013!b1343|it-rt-voopoo-dq66ezwa!b39"
password = "536e0fb9-d93b-4d67-be08-884caa8fa1d1$ATTXuy-4sLJZ9OPMaBuNexlnjR2sidQ_7Se9nJ8keMU="
auth_header = basic_auth_header(username, password)

url = root + produce_path
headers = {'Authorization': auth_header}
body = {"DATE_FROM": start_date, "DATE_TO": end_date, "ORDERS": ORDERS}

req = requests.post(url, headers=headers, json=body)

info = json.loads(req.text)
items = info['DATA']
json_list = []
if isinstance(items, dict):
    json_list.append(items)
else:
    json_list = items

df = pd.DataFrame(json_list)

# 应射系统字段名
df.rename(
    columns={'AUFNR': 'Order_Id', 'MATNR': 'Product_Code', 'MAKTX': 'Product_Name', 'ZSPEC': 'Product_Specification',
             'ZMODEL': 'model', 'DISPO': 'MRP_Control', 'GAMNG': 'Product_Qty', 'GWEMG': 'In_Qty',
             'GSTRP': 'Start_Date', 'GLTRP': 'End_Date', 'FTRMI': 'Approve_Date', 'GETRI': 'In_Date',
             'GASTAT': 'Product_Status', 'GWSTAT': 'In_Status', 'OVERDUE': 'Overdue', 'ZBRAND': 'Business_Group',
             'AEDAT': 'Sap_timestamp', 'AUART': 'Order_Type'}, inplace=True)
df = df.replace('0000-00-00', None)


# 获得预计完工日期范围
def get_date_range(date_str):
    rels_str = date_str
    if date_str:
        input_date = datetime.strptime(date_str, "%Y-%m-%d")
        weekday = input_date.weekday()

        # 计算与上周六的日期差
        days_to_subtract = (weekday + 2) % 7
        start_date = input_date - timedelta(days=days_to_subtract)

        # 计算与这周五的日期差
        days_to_add = (4 - weekday) % 7
        end_date = input_date + timedelta(days=days_to_add)

        rels_str = f'{start_date.strftime("%Y%m%d")}-{end_date.strftime("%Y%m%d")}'
    return rels_str


# 系统字段
df['week_date'] = df['End_Date'].apply(get_date_range)
df['week_date'] = df['week_date'].astype(str)
df['week_date'] = df['week_date'].str.replace(r"[()'-]", "", regex=True)
df['week_date'] = df['week_date'].str.replace(", ", "-")

now = datetime.now()
df['timestamp'] = now

formatted_datetime = now.strftime("%Y%m%d%H%M%S")
new_index = [formatted_datetime + str(num) for num in range(len(df['Order_Id']))]
# 将格式化后的日期时间设置为索引
df = df.reset_index().rename(columns={'index': 'id'})
df['id'] = new_index

# df.to_excel('test.xlsx', index=False)
engine_scc_data = database.engine('scc_data')

df.to_sql('work_order', con=engine_scc_data, if_exists='append', index=False)

# 若出现订单号重复，则用新数据替换旧数据
database.execute(
    "DELETE FROM work_order WHERE NOT EXISTS (SELECT 1 FROM (SELECT Order_Id, MAX(id) AS max_id FROM work_order GROUP BY Order_Id) AS t WHERE work_order.Order_Id = t.Order_Id AND work_order.id = t.max_id)",
    engine_scc_data)
