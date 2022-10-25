from turtle import onclick
import pandas as pd, xlwings as xw
import os

path='raw/'

Tmall_list = pd.DataFrame()
Tmall_detail = pd.DataFrame()
Alipay = pd.DataFrame()

for file in os.listdir(path):
    if file.endswith('.xlsx'):
        Tmall_list = Tmall_list.append(pd.read_excel(path+file, converters={"订单编号": str,"支付单号": str}))

for file in os.listdir(path):
    if file.startswith(('0', '1', '2', '3', '4', '5', '6', '7', '8', '9')):
        Alipay = Alipay.append(pd.read_csv(path+file, converters={"Partner_transaction_id": str, "Transaction_id": str}))

for file in os.listdir(path):
    if file.startswith('ExportOrderDetailList'):
        Tmall_detail = Tmall_detail.append(pd.read_csv(path+file,encoding='gbk'))


Tmall_list = Tmall_list[['订单编号','确认收货时间']] #筛选只需要的一些列


# # 最好先把支付宝的列明修改与天猫一直的话，合并的时候，就不会多一列
Tmall_list.rename(columns={'订单编号':'Partner_transaction_id'}, inplace=True)
list_meger = pd.merge(Alipay,Tmall_list,how='left')
list_meger = list_meger['确认收货时间'].fillna(axis=1,method='ffill')

print(list_meger)