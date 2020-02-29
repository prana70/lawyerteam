#coding:utf8

import pandas as pd
import os

#准备拟添加的法律日志报表数据，清理、排序，便于合并！
df1=pd.read_excel(os.getcwd()+'\\data\\法律日志报表20200101-20200131.xls') 
df1.rename(columns={'填报人':'提交人','填报时间':'提交时间','备注：':'备注'},inplace=True)
df1.drop(['工号','部门','最后一次修改时间','图片地址','地址','评论信息'],axis=1,inplace=True)
df1['提交时间']=df1['提交时间'].str.replace('[年月]','-',).str.replace('日','')
df1=df1[['客户名称','项目名称','办理日期','事务类型','耗费时间','涉及标的','工作内容','提交人','提交时间','备注']]


#导入，备份原“数据导出.xlsx”
df2=pd.read_excel(os.getcwd()+'\\data\\数据导出.xlsx')
os.rename(os.getcwd()+'\\data\\数据导出.xlsx',os.getcwd()+'\\data\\数据导出_备份20200101.xlsx')

#合并数据表，并生成新“数据导出.xlsx”
df=pd.concat([df1,df2])
df.dropna(axis=0,how='all',inplace=True)
df[['提交时间']]=df[['提交时间']].astype('datetime64[ns]')
print(df)
print(df.dtypes)
df.to_excel(os.getcwd()+'\\data\\数据导出.xlsx',index=False)
