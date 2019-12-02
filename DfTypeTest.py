#coding:utf8

import pandas as pd
import os

df1=pd.read_excel(os.getcwd()+'\\data\\数据导出.xlsx')
#print(df1)
df1[['提交时间']]=df1[['提交时间']].astype('datetime64[ns]')
df1.to_excel(os.getcwd()+'\\data\\数据导出.xlsx')
df1=pd.read_excel(os.getcwd()+'\\data\\数据导出.xlsx',index=False)
print(df1.dtypes)
