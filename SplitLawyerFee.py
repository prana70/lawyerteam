#coding: utf8

import pandas as pd
import re
import os
import salary as sl
import lawyerfee as lf

pd.options.display.float_format = '{:,.2f}'.format



def MakeRegularExpression(str): #根据输入的关键字str生成正则表达式
    SplitChar=r'[\s\,\;\，\；\、\.\。]'
    print(str)
    SubStrs=re.split('\|',str)
    expressions=[]
    for SubStr in SubStrs:
        KeyWords=re.split(SplitChar,SubStr)
        #print('KeyWords:',KeyWords)
        expression=''
        for element in KeyWords:
            expression+='(?=.*?'+element+')'
        #print('expression:',expression)
        expressions.append(expression)
        #print(expressions)
    FinalExpression='|'.join(expressions)
    return FinalExpression
      
    
    '''
    if list:
        expression=r''
        for element in list:
            expression+='(?=.*?'+element+')'
        return expression
    else:
        return None
    '''

def GetAbilityFactor(name,date): #根据律师（提交人）姓名、办理日期，返回当条法律事务记录的能力等级
    return lf.get_acoefficient(sl.get_ability_level(name,date))


def GetOwner(name,date): #根据律师（提交人）姓名、办理日期，返回当条法律事务记录的对应分配主体
    df1=pd.read_excel(os.getcwd()+'\\config\\律师身份.xlsx')
    df2=df1[df1.姓名==name].sort_values(by=['取得日期'],ascending=False)
    for index,row in df2.iterrows():
        if date>=str(row['取得日期'])[:10]:
            if row['身份']=='授薪律师':
                if date>='2017-10-01':
                    return '团队（%s）'%name
                else:
                    return '王彬（%s）'%name
            return name
    return None
            
def GetTaxRatio(name): #根分配主体的名称，返回相的扣税比例
    df1=pd.read_excel(os.getcwd()+'\\config\\律师身份.xlsx')
    if name not in list(df1['姓名']): #如果分配主体为团队，则按返回0.21的扣税率
        #print('未查到身份，按0.21扣税')
        return 0.21
    else:
        df2=df1[df1.姓名==name].sort_values(by=['取得日期'],ascending=False)
        if df2.iloc[0]['身份']=='独立律师': #如果分配主体是独立律师，则返回0.27的扣税率
            #print('是独立律师，按0.27扣税')
            return 0.27
        elif df2.iloc[0]['身份']=='授薪律师': #如果分配主体是授薪，则表明之前的统计有错，提示查错！
            #print('是授薪律师，不能参加分配，请检查错误')
            return None
        else: #如果分配主体是非独立律师，则返回0.21的扣税率
            #print('非独立律师，按0.21扣税')
            return 0.21
    

if __name__=='__main__':
    df1=pd.read_excel(os.getcwd()+'\\config\\项目库.xlsx')
    df2=pd.read_excel(os.getcwd()+'\\data\\数据导出.xlsx')
    df1=df1[df1.办案律师费>0] #剔除未完成的项目，或者已完成但没有律师费的项目。
    for index,row in df1.iterrows():
        file=os.getcwd()+'\\report\\lawyerfee\\'+row['客户名称']+'_'+row['项目名称']+'_律师费分配.csv' #生成最终要生成的项目律师费分配文件名
        if not os.path.exists(file) and not pd.isnull(row['终止日期']): # and '浦' in row['客户名称']: #如果具有生成文件名的文件不存在，并且项目已终止，则往下生成，否则不管
            
            print(50*'*')
            print(row['客户名称'],row['项目名称'])
            print(50*'*')


            if row['办案律师费']>0:    #获取项目可分配的律师费
                LawyerFee=row['办案律师费']
            elif row['总律师费']>0:
                LawyerFee=row['总律师费']/2
            else:
                LawyerFee=0
            
            CustomerExpression=MakeRegularExpression(row['客户名称关键字']) #客户名称关键字的正则表达式
            #print('客户名称关键字正则表达式：',CustomerExpression)
            ProgramExpression=MakeRegularExpression(row['项目名称关键字'])
            #print('项目名称关键字正则表达式：',ProgramExpression)


            df3=df2[(df2.客户名称.str.match(CustomerExpression))& #通过设定客户名称、项目名称等条件生成项目明细库
                    (df2.项目名称.str.match(ProgramExpression))&
                    (df2.办理日期>=str(row['起始日期'])[:10])& #时间戳转换成str去掉多余的时间，以便于比较,下同
                    (df2.办理日期<=str(row['终止日期'])[:10])]. \
                    fillna({'耗费时间':0,'涉及标的':0}) 
            #print(df3)
            df3['能力系数']=df3.apply(lambda row:GetAbilityFactor(row['提交人'],row['办理日期']),axis=1)
            df3['分配主体']=df3.apply(lambda row:GetOwner(row['提交人'],row['办理日期']),axis=1)
            df3['有效工作时间']=df3['耗费时间']*df3['能力系数']
            print(df3)
            df4=df3.groupby('分配主体',as_index=False)['耗费时间','有效工作时间'].sum()
            df4['总时间']=df4['有效工作时间'].sum()
            df4['分配比例']=df4['有效工作时间']/df4['总时间']
            df4['可分配律师费']=LawyerFee
            df4['应分配税前律师费']=df4['可分配律师费']*df4['分配比例']
            df4['扣税比例']=df4.apply(lambda row:GetTaxRatio(row['分配主体']),axis=1)
            df4['应分配税后律师费']= df4['应分配税前律师费']*(1-df4['扣税比例'])
            #print(df4[['分配主体','可分配律师费','有效工作时间','分配比例','应分配税前律师费','扣税比例','应分配税后律师费']])
            df4[['分配主体','可分配律师费','耗费时间','有效工作时间','分配比例','应分配税前律师费','扣税比例','应分配税后律师费']]. \
            to_csv(file,index=False)
            print('生成完毕！')
    
