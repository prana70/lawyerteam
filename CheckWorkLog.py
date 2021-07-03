#coding:utf8

from SpiderDingding import get_worklog_from_dingding as gw

start='2021-05-10'
end='2021-05-10'

df=gw(start,end)

for index,row in df.iterrows():
    print(index,row)
