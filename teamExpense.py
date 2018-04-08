#coding:utf-8

import os
import pandas as pd

dir=os.getcwd()+'\\report\\salary\\'
fileList=os.listdir(dir)
#print(fileList)


finalDf=pd.DataFrame(columns=['项目','姓名','合计（元）'])

#print(finalDf)
for file in fileList:
    fileName=dir+file
    #print(fileName)
    df=pd.read_csv(fileName,encoding='gb2312')
    #print(df)
    df['项目']=file[:-4]
    df1=df[['项目','姓名','合计（元）']]
    #print(df1)
    finalDf=pd.concat([finalDf,df1],ignore_index=True)
print(finalDf)
finalDf.to_csv(os.getcwd()+'\\report\\团队工资支出.csv',index=False)
