#coding: utf-8

import pandas as pd
import os

file=os.getcwd()+"\\data\\数据导出.xlsx"
df=pd.read_excel(file)
df1=df.groupby('提交人',as_index=False)['耗费时间'].sum()
print(df1)
