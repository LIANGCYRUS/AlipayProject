import pandas as pd, xlwings as xw, numpy as np
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
    if file.startswith('CPS'):
        CPSOrderLilst = CPSOrderLilst.append(pd.read_csv(path + file, converters={"淘宝父订单编号": str}))

'''
********** 代码说明 **********
1、先把必要的列留下来，没必要的列可以去除。
2、修改列明：原因 > 合并的时候key列的名称不一样，那么就会多一列出来，到时候还要删除。
3、把【确认收货时间】的空值先填充，以你为pandas合并的时候，空值和缺失值无法区分。一般来说，空值是没有确认时间，缺失值是分销订单
'''

NAN = 0
NA = '无确认收货时间'
AS = '售后退款'
DG = '分销订单'


TmgOrderLilst = TmgOrderLilst[['订单编号','确认收货时间']] #筛选只需要的一些列
CPSOrderLilst = CPSOrderLilst[['淘宝父订单编号']]

TmgOrderLilst.rename(columns={'订单编号':'Partner_transaction_id'}, inplace=True)

# CPS文件自带\t，所以要去掉
CPSOrderLilst = CPSOrderLilst.replace(r'\t','', regex=True)

A = pd.merge(TmgOrderLilst,CPSOrderLilst,how='outer',left_on='Partner_transaction_id',right_on='淘宝父订单编号')


def check_cps(a,b):
    if a == b:
        return '是'
    else:
        return '否'


A['是否淘宝客单'] = A.apply(lambda x : check_cps(x['Partner_transaction_id'],x['淘宝父订单编号']),axis = 1)

TmgOrderLilst = A[['Partner_transaction_id','确认收货时间','是否淘宝客单']]


TmgOrderLilst['确认收货时间'] = TmgOrderLilst['确认收货时间'].fillna(NAN) #将空值用N/A替换

'''
********** 代码说明 **********
1、将确认收货时间合并到支付宝列表中
2、修改列明：原因 > 合并的时候key列的名称不一样，那么就会多一列出来，到时候还要删除。
3、将Type为R的标记成为售后退款
4、将标记好111的在备注列上标记无确认收货时间
5、将确认收货时间仍然为空白的在备注列上标记分销订单
6、最后将备注上剩下的标注成为时间，以后的话，按照输入的年月，进行分类
'''

Confirmation_time_merge = pd.merge(AlipayLilst,TmgOrderLilst,on='Partner_transaction_id',how='left')

# 先把Type为R的在备注上标记AS
Confirmation_time_merge.loc[(Confirmation_time_merge['Type'] == 'R') ,'备注'] = AS
# 在合并之前,在天猫订单列表中,把空格都用0来表示,所以这里如果等于0的话,就表示没有确认收货时间.为什么要写0是因为下面的操作,要换算成为时间,如果用文本来填充的话,会出错.
Confirmation_time_merge.loc[(Confirmation_time_merge['确认收货时间'] == 0) ,'备注'] = NA
# 再把确认时间为空的以及备注也是空的,那么可以判断为分销订单了
Confirmation_time_merge.loc[((Confirmation_time_merge['确认收货时间'].isnull()) & (Confirmation_time_merge['备注'].isnull())) ,'备注'] = DG
# 把所有订单种类标记后，再把确认收货时间为空的单元格填充为0，以便下一步转换成为时间类型
Confirmation_time_merge.loc[(Confirmation_time_merge['确认收货时间'].isnull()) ,'确认收货时间'] = 0

'''
********** 将淘宝客订单区分 **********
'''
# Confirmation_time_merge['备注'] = np.where(Confirmation_time_merge['Partner_transaction_id'] == CPSOrderLilst['Partner_transaction_id'],'淘宝客订单','')

# for i in CPSOrderLilst['Partner_transaction_id']:
#     if Confirmation_time_merge['Partner_transaction_id'] == i:
#         Confirmation_time_merge['备注'] = '111'



# 把文本型的时间转换成为时间类型
Confirmation_time_merge['确认收货时间'] = pd.to_datetime(Confirmation_time_merge['确认收货时间'])
# 提取时间的年月,标记到新的一列,这一列准备用于区分确认收货时间
Confirmation_time_merge['month'] = Confirmation_time_merge['确认收货时间'].apply(lambda x:x.strftime('%Y-%m'))

'''
********** 以上所有的订单类型已经区分完成 **********
'''

this_month_input = input('请输入需要结算的年月份（2022-10）:')
this_month = Confirmation_time_merge.loc[(Confirmation_time_merge['month'] == this_month_input) & (Confirmation_time_merge['是否淘宝客单'] == '否')]


AAS = Confirmation_time_merge.loc[(Confirmation_time_merge['备注'] == '售后退款')]
DG = Confirmation_time_merge.loc[(Confirmation_time_merge['备注'] == DG)]
CPS = Confirmation_time_merge.loc[(Confirmation_time_merge['是否淘宝客单'] == '是')]
#
#
# # print(AAS)
#
file_name = 'RAW_MERGE.xlsx'
Confirmation_time_merge.to_excel(file_name)
print(file_name+'导出成功')
#

wb = xw.Book('raw/template.xlsx')
ws = wb.sheets[0]
ws2 = wb.sheets[1]
ws3 = wb.sheets[2]
ws4 = wb.sheets[3]

range2 = ws.range('C:D')
range2.api.NumberFormat ="@"

range3 = ws2.range('C:D')
range3.api.NumberFormat ="@"

range4 = ws3.range('C:D')
range4.api.NumberFormat ="@"

range5 = ws4.range('C:D')
range5.api.NumberFormat ="@"

# # 进行赋值
ws.range('B9').options(pd.DataFrame, index=True).value = this_month
ws.range('B1:Z10000').columns.autofit()


ws2.range('B9').options(pd.DataFrame, index=True).value = DG
ws2.range('B1:Z10000').columns.autofit()

ws3.range('B9').options(pd.DataFrame, index=True).value = AAS
ws3.range('B1:Z10000').columns.autofit()

ws4.range('B9').options(pd.DataFrame, index=True).value = CPS
ws4.range('B1:Z10000').columns.autofit()

wb.save('final.xlsx')
print('成功导出')

'''
********** 代码说明 **********
1、将确认收货时间合并到支付宝列表中
2、修改列明：原因 > 合并的时候key列的名称不一样，那么就会多一列出来，到时候还要删除。
3、将Type为R的标记成为售后退款
4、将标记好111的在备注列上标记无确认收货时间
5、将确认收货时间仍然为空白的在备注列上标记分销订单
6、最后将备注上剩下的标注成为时间，以后的话，按照输入的年月，进行分类
'''