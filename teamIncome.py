#coding:utf-8

import os
import pandas as pd

dir=os.getcwd()+'\\report\\lawyerfee\\'
fileList=os.listdir(dir)
#print(fileList)
for file in fileList:
    fileName=dir+file
    #print(fileName)
    df=pd.read_csv(fileName,encoding='gbk')
    df['项目']=file[:-10]
    print(df.columns)
    print(df[' 应分配税后律师费 '])
    #df2=df[['分配主体','应分配税后律师费']]
    #print(df2)
