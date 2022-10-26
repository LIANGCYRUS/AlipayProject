import pandas as pd, xlwings as xw
import os
import datetime

# 指定测试文件路径
path = 'raw/'

#创建DataFrame用于保存读取进来的excel数据

# 天猫订单列表
TmgOrderLilst = pd.DataFrame()
# 天猫宝贝列表
TmgOrderDetailLilst = pd.DataFrame()
# 支付宝列表
AlipayLilst = pd.DataFrame()
# 淘宝客列表
CPSOrderLilst = pd.DataFrame()

'''
根据文件名称,自动把同种类的文件合并在一起
'''
for file in os.listdir(path):
    # 将文件后缀为.xlsx的文件全部合并到TmgOrderLilst
    if file.endswith('.xlsx'):
        TmgOrderLilst = TmgOrderLilst.append(pd.read_excel(path + file, converters={"订单编号": str, "支付单号": str}))
    # 将文件名字开头为数字的文件全部合并到AlipayLilst
    if file.startswith(('0', '1', '2', '3', '4', '5', '6', '7', '8', '9')):
        AlipayLilst = AlipayLilst.append(pd.read_csv(path + file, converters={"Partner_transaction_id": str, "Transaction_id": str}))
    # 将名字开头为ExportOrderDetailList的文件全部合并到
    if file.startswith('ExportOrderDetailList'):
        TmgOrderDetailLilst = TmgOrderDetailLilst.append(pd.read_csv(path+file,encoding='gbk'))

'''
********** 代码说明 **********
1、先把必要的列留下来，没必要的列可以去除。
2、修改列明：原因 > 合并的时候key列的名称不一样，那么就会多一列出来，到时候还要删除。
3、把【确认收货时间】的空值先填充，以你为pandas合并的时候，空值和缺失值无法区分。一般来说，空值是没有确认时间，缺失值是分销订单
'''

NAN = 'N/A'
NA = '无确认收货时间'
AS = '售后退款'
DG = '分销订单'


TmgOrderLilst = TmgOrderLilst[['订单编号','确认收货时间']] #筛选只需要的一些列
TmgOrderLilst.rename(columns={'订单编号':'Partner_transaction_id'}, inplace=True)
TmgOrderLilst['确认收货时间'] = TmgOrderLilst['确认收货时间'].fillna(NAN) #将空值用N/A替换

# print(TmgOrderLilst)


'''
********** 代码说明 **********
1、将确认收货时间合并到支付宝列表中
2、修改列明：原因 > 合并的时候key列的名称不一样，那么就会多一列出来，到时候还要删除。
3、将Type为R的标记成为售后退款
4、将标记好111的在备注列上标记无确认收货时间
5、将确认收货时间仍然为空白的在备注列上标记分销订单
'''

Confirmation_time_merge = pd.merge(AlipayLilst,TmgOrderLilst,on='Partner_transaction_id',how='left')

Confirmation_time_merge.loc[(Confirmation_time_merge['Type'] == 'R') ,'备注'] = AS
Confirmation_time_merge.loc[(Confirmation_time_merge['确认收货时间'] == NAN) ,'备注'] = NA
Confirmation_time_merge.loc[((Confirmation_time_merge['确认收货时间'].isnull()) & (Confirmation_time_merge['备注'].isnull())) ,'备注'] = DG


'''
********** 代码说明 **********
1、将确认收货时间合并到支付宝列表中
2、修改列明：原因 > 合并的时候key列的名称不一样，那么就会多一列出来，到时候还要删除。
3、将Type为R的标记成为售后退款
4、将标记好111的在备注列上标记无确认收货时间
5、将确认收货时间仍然为空白的在备注列上标记分销订单
'''

Split_List = ['',DG]





Confirmation_time_merge.to_excel('output.xlsx',index=False)
