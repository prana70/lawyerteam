#coding:utf-8


import pandas as pd

import os


#根据姓名、起止日期返回律师工作时间
def get_time_consuming(name,StartDate,EndDate):
    df=pd.read_excel(os.getcwd()+'\\data\\数据导出.xlsx')
    return df[(df.提交人==name) & (df.提交时间>=StartDate) & (df.提交时间<EndDate)]['耗费时间'].sum()


#根据姓名,指定日期返回律师的当前能力定级
def get_ability_level(name,OppointedDate):
    df=pd.read_excel(os.getcwd()+'\\config\\律师能力定级.xlsx')
    df1=df[df.姓名==name].sort_values(by=['定级日期'],ascending=False)
    LevelDates=list(df1['定级日期'])
    for LevelDate in LevelDates:
        if OppointedDate>=str(LevelDate)[0:10]:
            return df1[df1.定级日期==LevelDate].iloc[0,2]
    return '法律实习生'



#根据能力定级返回律师的基本工资和奖金系数
def get_bsalary_bcoefficient(AbilityLevel):
    df=pd.read_excel(os.getcwd()+'\\config\\律师能力等级及薪酬.xlsx')
    df1=df[df.能力等级==AbilityLevel]
    return df1.iloc[0,1],df1.iloc[0,2]

if __name__=='__main__':

    year_=input('请输入年份：')
    month_=input('请输入月份：')

    StartYear=year_
    EndYear=year_


    StartMonth=str(int(month_)-1)
    EndMonth=month_

    if int(EndMonth)<10:
        EndMonth='0'+EndMonth
        
    if StartMonth=='0':
        StartMonth='12'
        StartYear=str(int(StartYear)-1)
    elif int(StartMonth)<10:
        StartMonth='0'+StartMonth

    StartDate=StartYear+'-'+StartMonth+'-26'
    EndDate=EndYear+'-'+EndMonth+'-26'
    print(StartDate,'至',EndDate)

    #获取授薪律师姓名列表
    df=pd.read_excel(os.getcwd()+'\\config\\授薪律师名单.xlsx')
    names=df['姓名']

    f=open(os.getcwd()+'\\report\\salary\\授薪律师工资（'+year_+'年'+month_+'月).csv','w')
    f.write('姓名,基本工资(元),工作时间（分钟）,奖金（元）,合计（元）\n')
    for name in names:
        TimeConsuming=get_time_consuming(name,StartDate,EndDate)
        AbilityLevel=get_ability_level(name,EndDate)
        BasicSalary,BonusCoefficient=get_bsalary_bcoefficient(AbilityLevel)
        print('计算并写入%s的工资......'%name)
        f.write(name+','+format(BasicSalary,'0.0f')+','+format(TimeConsuming,'0.0f')+','+format(TimeConsuming/60*BonusCoefficient,'0.0f')+','+format((BasicSalary+TimeConsuming/60*BonusCoefficient),'0.0f')+'\n')
    print('OK!所有人的工资已计算并写入完成！')
    f.close()
