#coding:utf-8

import pandas as pd
import salary as sl
import os

#根据能力定级返回律师的能力系数
def get_acoefficient(AbilityLevel):
    df=pd.read_excel('律师能力等级及薪酬.xlsx')
    df1=df[df.能力等级==AbilityLevel]
    return df1.iloc[0,7]


if __name__=='__main__':
    ProgramName=input('请输入项目名称：')
    LawyerFee=int(input('请输入可分配律师费总额：'))
    df=pd.read_excel(os.getcwd()+'\lawyerfee\\'+ProgramName+'.xlsx')[['提交人','办理日期','耗费时间']]


    ListOfPayedLawyers=list(pd.read_excel('授薪律师名单.xlsx')['姓名'])#提取授薪律师姓名构建列表




    #计算有效工作时间，并生成新dataframe---df1
    data=[]
    for xxx in df.values:
        yyy=list(xxx)
        AbilityLevel=sl.get_ability_level(yyy[0],yyy[1])
        AbilityCoefficient=get_acoefficient(AbilityLevel)
        yyy.append(AbilityCoefficient)
        yyy.append(yyy[2]*AbilityCoefficient)
        data.append(yyy)
        #区分授薪律师的工作时间，2017年10月以前的归属王彬，以后的归属团队
        if yyy[0] in ListOfPayedLawyers:
            if yyy[1]>='2017-10-01':
                yyy[0]='团队（%s）'%yyy[0]
            else:
                yyy[0]='王彬（%s）'%yyy[0]

    df1=pd.DataFrame(data,columns=['姓名','办理日期','工作时间','能力系数','有效时间'])
    print(df1)

    #按律师姓名进行将工作时间和有效工作时间进行分组汇总
    df2=df1.groupby('姓名').sum()[['工作时间','有效时间']]
    print(df2)

    #根据分组汇总结果计算工作时间比例
    df3=df2.groupby(level=0).agg({'有效时间':'sum'}).apply(lambda x: 100*x/ x.sum())
    df3.rename(columns={'有效时间':'时间比例'},inplace=True)
    print(df3)

    #拼接并格式化表
    df4=pd.merge(df2,df3,left_index=True,right_index=True)
    df4['可分律师费总额（元）']=LawyerFee
    df4['应分含税律师费（元）']=df4['时间比例']*LawyerFee/100
    df4.rename(columns={'工作时间':'工作时间（分钟）','有效时间':'有效时间（分钟）','时间比例':'时间比例（%）'},inplace=True)
    df4['扣税比例（%）']=21

    #根据律师身份判断扣税比例，并计算应分扣税律师费
    DFOfLawyerRole=pd.read_excel('律师身份.xlsx')

    ListOfIndependentLawyers=list(DFOfLawyerRole[DFOfLawyerRole.身份=='独立律师']['姓名'])
    for line in df4.index:
        if line in ListOfIndependentLawyers:
            df4.ix[line,['扣税比例（%）']]=27
        print(df4.ix[line,['扣税比例（%）']])
    df4['应分扣税律师费（元）']=df4['应分含税律师费（元）']*(1-df4['扣税比例（%）']/100)
    print(df4.round(2))

    df4.round(2).to_csv(os.getcwd()+'\lawyerfee\\'+ProgramName+'律师费分配.csv')

