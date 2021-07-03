#coding:utf8

import requests as rq
import json
import pandas as pd
import sqlite3 as sql
from sqlalchemy import create_engine
import datetime


def get_worklog_from_dingding(startTime,endTime): #根据起止日期，从dingding调取工作日志，并清洗成规整的表
    #访问登录页，以取得访问许可
    print('调取工作日志：从 ',startTime,' 到 ',endTime)
    url_login='https://ynuf.aliapp.org/service/um.json'
    head={
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Origin': 'https://im.dingtalk.com',
        'Referer': 'https://im.dingtalk.com/',
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36'
        }
    data={
        'data': '106!UTEmc0cl/Q6Hc39fHmHI4lxmYX+lQmsE5UamzLHtKj7U5xHey3omNx6PzEHCnVWwSFfJGAmYu5oyv0+CYB0Dm8nfwThWeRkVMs5i/joeN59VnASX4fXSlhfm+InDfJz2AlrsKAH+yDwAS2pfMCMKCCh1URFsPsXJOU+c4v/SP7Egb6mdy23rFHZf5fJwsyXGDC/cvvY9IWcP/RaH6uOP//N1Qy989yZD8ykE4u9XsPYs/FkU2UVa1WglaZnevWgRUb4pXiRIC34W1lblKILX1ECYaodhvERqk3Uxv1t8UA1O1LblGRj3g8U4a5snvhVqVoClv/FEC3231kdlKMNN14C5llnPTsbKsh0HTPRPJXjMtOgS+w0xEYdG0ettPGkgVaHnxEmVWJ5mW5xXmExiD6EvtbTM733fuwbxR6l2Rw9bP0lXiJZKNaRYUg3QyeglmKJaxI1PYQyAn1BnsInv97Vu1bymTo6W+kYeUaD51aSeYfgoGrGC2AVqjO1sh1S9aB2L2Dxw2qWQyp113vQXUN9ddrdcoo/BGH7vsyXjRRG69akzJm4Qam3N+IDSHld6+azXFoA2/q3PIlL4hJOqS563SNtNBhhTBQYgst9sR0tB5pKnzkjD/cp7wZBbxH3WA4y5WvmqJar06ZowUcSt9SOrZOC2xw5/he9Nl4qccD+AFJ8Whp+vpFFzg+IdfonzCsgl',
        'xa': 'dingding',
        'xt': '' 
        }
    sn=rq.Session()
    #sn.post(url_login,data=data,headers=head)

    #根据设定日期，爬取日志数据，注意：日期期间不可设置太宽，因为每次调取有500条记录的限制，期间太宽，会导致超过500条的记录无法爬取。
    url='https://landray.dingtalkapps.com/alid/reportpc/client/getTotalDetail?templateid=157417b767aa236cb549d174541b58c6&corpid=dinge242679ea89dbb6f&startTime='+startTime+'&endTime='+endTime+'&pageNow=0&pageSize=0&templateid=157417b767aa236cb549d174541b58c6&_=1557889300543'
    headers={
        'authority': 'landray.dingtalkapps.com',
        'method': 'GET',
        'path': '/alid/reportpc/client/getTotalDetail?templateid=157417b767aa236cb549d174541b58c6&corpid=dinge242679ea89dbb6f&startTime=2020-01-01&endTime=2020-03-09&pageNow=0&pageSize=0&templateid=157417b767aa236cb549d174541b58c6&_=1557889300543',
        'scheme': 'https',
        'accept': 'application/json, text/javascript, */*; q=0.01',
        'accept-encoding': 'gzip, deflate, br',
        'accept-language': 'zh-CN,zh;q=0.9',
        'client-corpid': 'dinge242679ea89dbb6f',
        'content-type': 'application/json',
        'cookie': 'cna=u8QnE1h1zRACAbaVzieqNiJE; isg=BIeH6gKQl9ToMxPB51ilzoDwFjuRJER8eJ6ng1l0o5Y9yKeKYVzrvsWJboiWRDPm; _bl_uid=U8jOpv4O69zeqmrzOqb86aLyLdIU; JSESSIONID=31EEB8F5C6FFFDB0CE7CF94A316457DB; sid=31EEB8F5C6FFFDB0CE7CF94A316457DB; sid=31EEB8F5C6FFFDB0CE7CF94A316457DB; LR_TOKEN=eyJjb3JwSWQiOiJkaW5nZTI0MjY3OWVhODlkYmI2ZiIsInZhbHVlIjoiJENLX0tleUNlbnRlcl92MS4wJFNxbG00dThlQjk2M0h2eWNBTERUZHAwS3RDMzBxb3lBT2lmTmdxMzl5UWpXdEJmSHJkaG0wTEthVTZkYVdpOTgifQ==; LR_TOKEN=eyJjb3JwSWQiOiJkaW5nZTI0MjY3OWVhODlkYmI2ZiIsInZhbHVlIjoiJENLX0tleUNlbnRlcl92MS4wJFNxbG00dThlQjk2M0h2eWNBTERUZHAwS3RDMzBxb3lBT2lmTmdxMzl5UWpXdEJmSHJkaG0wTEthVTZkYVdpOTgifQ==; LR_TOKEN_N=JENLX0tleUNlbnRlcl92MS4wJEFyYTZZckpOMVVnYlk3dmh6UmQydXJJek1YVEdhMExGdVlaQ245TTh5empXSlpwR241RkxnNWQ4TlJJNWZCTEdma0M1cWd4aDI1NU9ucHJsVUMrY3JwUVM5KzNzeUwxK2RxSzYzeUVYOFRvPQ==; LR_TOKEN_N=JENLX0tleUNlbnRlcl92MS4wJEFyYTZZckpOMVVnYlk3dmh6UmQydXJJek1YVEdhMExGdVlaQ245TTh5empXSlpwR241RkxnNWQ4TlJJNWZCTEdma0M1cWd4aDI1NU9ucHJsVUMrY3JwUVM5KzNzeUwxK2RxSzYzeUVYOFRvPQ==',
        'random-numbers': '15578893005436648662',
        'referer': 'https://landray.dingtalkapps.com/alid/app/reportpc/tablereport.html?corpid=dinge242679ea89dbb6f',
        'user-agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36',
        'x-client-corpid': 'dinge242679ea89dbb6f',
        'x-requested-with': 'XMLHttpRequest',
        }

    resp=sn.get(url,headers=headers)
    print('访问状态代码：',resp.status_code)


    resp_json=resp.json()

    #将日志记录清理出来形成pandas表，此时列名为数字代码，需利用title_name字典将列名转换成文字名称
    print('生成初始日志记录表......')
    df2=pd.DataFrame(resp_json['reports'])

    #初始日志表是否是空记录，是的话，返回None，不是的话，继续处理
    record_count=len(df2)

    if record_count<=0:
        print(record_count,'条记录！')
        return None
    else:
        print(record_count,'条记录！')
    
        #将列名和列代码清理出来，生成以列代码为键名的字典
        print('生成列名字典......')
        df1=pd.DataFrame(resp_json['titleNames'])
        df1['sort']=df1['sort'].astype(str)
        title_name=dict(zip(df1['sort'],df1['titleName']))


        print('转换列名......')
        for sort in title_name:
            df2.rename(columns={sort:title_name[sort]},inplace=True)

        
        #将正式日志记录表中的\n\r\t等字符删除
        print('清理正式日志记录中\n\r\t')
        df2=df2.applymap(lambda x:x.strip())

        #将正式日志记录表中的时间、数字列，从object转换成相应类型
        df2['填报时间']=pd.to_datetime(df2['填报时间'],errors='raise',format='%Y年%m月%d日 %H:%M')
        df2['未读人数']=pd.to_numeric(df2['未读人数'])
        df2['已读率']=df2['已读率'].str.strip('%').astype(float)/100
        df2['办理日期']=pd.to_datetime(df2['办理日期'],errors='raise')
        df2['耗费时间']=pd.to_numeric(df2['耗费时间'])
        df2['涉及标的']=pd.to_numeric(df2['涉及标的'])
        df2['评论数']=pd.to_numeric(df2['评论数'])
        df2['点赞数']=pd.to_numeric(df2['点赞数'])
        df2['已读人数']=pd.to_numeric(df2['已读人数'])
        return df2

if __name__=='__main__':

    '''
    #将正式日志记录表写入test.xls文件
    print('将正式日志记录表写入test.xls文件......')
    df2.to_excel('test.xls',index=False)
    print('将正式日志记录表写入test.xls文件完成！')
    '''

    #print(get_worklog_from_dingding('2015-01-01','2016-12-31'))

    #设定工作日志爬取期间
    start=datetime.date(2020,1,1)
    end=datetime.date.today()

    #判断期间中的每一天，如果爬过，就不爬了；没爬过，则爬取，并在'已爬取工作日志的日期.txt'中记录；爬下来后，如果是零记录，则不处理，非也，则写入sqlite数据库。
    for i in range((end-start).days+1):
        day_raw=start+datetime.timedelta(days=i)
        if day_raw<datetime.date.today(): #判断爬取时间是否到了！
            day=str(day_raw)
            f=open('SpideredWorklogDate.txt','r')
            if day in f.read():
                print(day,'已爬取过，本次不爬取！')
                f.close()
            else:
                f=open('SpideredWorklogDate.txt','a')
                df=get_worklog_from_dingding(day,day)
                if not df is None:
                    #将正式日志记录表写入test.db中的worklog表
                    print('将正式日志记录表写入data.db文件......')
                    con=sql.connect('data.db')
                    df.to_sql('worklog',con=con,if_exists='append',index=False)
                    con.close()
                    print('将正式日志记录表写入data.db文件完成！')
                f.write(day+'\n')
                f.close()
        else:
            print(day_raw,'爬取时间未到，请过几天再试！')
        
        
    print('连接data.db...')
    engine=create_engine('sqlite:///data.db')
    df3=pd.read_sql('worklog',engine)
    print(df3.dtypes)

