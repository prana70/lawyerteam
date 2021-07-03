#coding:utf8

import pandas as pd
import re
import os
import salary as sl
import lawyerfee as lf
from sqlalchemy import create_engine
import datetime


engine=create_engine('sqlite:///data.db')
df=pd.read_sql('worklog',engine)
df=df[(df.办理日期.dt.date>=datetime.date(2020,11,1))&(df.办理日期.dt.date<=datetime.date(2020,11,7))]
print(df)
df.to_excel('20201101-20201117工作日志.xls')
