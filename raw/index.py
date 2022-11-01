import pandas as pd, xlwings as xw, numpy as np
import os
import datetime

# 指定测试文件路径
path = 'raw/'

#创建DataFrame用于保存读取进来的excel数据

# 天猫订单列表
TMOrderList = pd.DataFrame()
# 天猫宝贝列表
TMOrderDetailList = pd.DataFrame()
# 支付宝列表
AlipayLilst = pd.DataFrame()
# 淘宝客列表
CPSOrderLilst = pd.DataFrame()
# 地宫列表
DGOrderLilst = pd.DataFrame()

for file in os.listdir(path):
    # 将文件后缀为.xlsx的文件全部合并到TmgOrderLilst
    if file.endswith('ExportOrderList'):
        TMOrderList = TMOrderList.append(pd.read_excel(path + file, converters={"订单编号": str, "支付单号": str}))

    if file.startswith('ExportOrderDetailList'):
        TMOrderDetailList = TMOrderDetailList.append(pd.read_csv(path+file,encoding='gbk'))

    # 将文件名字开头为数字的文件全部合并到AlipayLilst
    if file.startswith(('0', '1', '2', '3', '4', '5', '6', '7', '8', '9')):
        AlipayLilst = AlipayLilst.append(pd.read_csv(path + file, converters={"Partner_transaction_id": str, "Transaction_id": str}))

    if file.endswith('订单结算明细报表.csv'):
        CPSOrderLilst = CPSOrderLilst.append(pd.read_csv(path + file, converters={"淘宝父订单编号": str}))


