#coding:utf-8

import os
import pandas as pd

dir=os.getcwd()+'\\report\\lawyerfee\\'
fileList=os.listdir(dir)
#print(fileList)
finalDf=pd.DataFrame(columns=['分配主体', '可分配律师费', '耗费时间', '有效工作时间', '分配比例', '应分配税前律师费', '扣税比例',
       '应分配税后律师费', '项目'])
#print(finalDf)
for file in fileList:
    fileName=dir+file
    #print(fileName)
    df=pd.read_csv(fileName,encoding='gb2312')
    df['项目']=file[:-10]
    finalDf=pd.concat([finalDf,df],ignore_index=True)
print(finalDf[finalDf.分配主体.str.contains('团队')][['项目','分配主体','应分配税后律师费']])
finalDf[finalDf.分配主体.str.contains('团队')][['项目','分配主体','应分配税后律师费']].to_csv(os.getcwd()+'\\report\\团队收入.csv',index=False)
