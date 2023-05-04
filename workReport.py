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
print(df)
df.to_excel('工作日志.xls')
