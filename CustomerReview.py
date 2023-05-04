#coding:utf8

import unprojected as up
import pandas as pd
import os
from sqlalchemy import create_engine


if __name__=='__main__':
    project_df=pd.read_excel(os.getcwd()+'\\config\\项目库.xlsx')
    worklog_df=pd.read_sql('worklog',create_engine('sqlite:///data.db'))
    program_statistics_lists=[]
    for index,row in project_df.iterrows():
        customer_expression=up.MakeRegularExpression(row['客户名称关键字'])
        program_expression=up.MakeRegularExpression(row['项目名称关键字'])
        program_worklog_df=worklog_df[(worklog_df.客户名称.str.match(customer_expression))&
                                      (worklog_df.项目名称.str.match(program_expression))&
                                      (worklog_df.办理日期.dt.date>=row['起始日期'].date())&
                                      (worklog_df.办理日期.dt.date<=row['终止日期'].date())].\
                                      fillna({'耗费时间':0,'涉及标的':0})
        program_statistics_dict={'客户名称':row['客户名称'],
                                 '项目名称':row['项目名称'],
                                 '律师费总额':row['律师费总额'],
                                 '办案律师费':row['办案律师费'],
                                 '服费件次':len(program_worklog_df),
                                 '耗费时间':program_worklog_df['耗费时间'].sum()
                                 }
        program_statistics_lists.append(program_statistics_dict)
    program_statistics_df=pd.DataFrame(program_statistics_lists)
    print(program_statistics_df)
    program_statistics_df.to_csv(os.getcwd()+'\\report\\program_statistics.csv')

