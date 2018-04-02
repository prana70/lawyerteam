#coding: utf-8
import pandas as pd

def valuation_formula(x,y):
    return x*y

df=pd.DataFrame({'AAA':[1,2,1,3],
                 'BBB':[1,1,2,2],
                 'CCC':[2,1,3,1]})
print(df)

df1=df.apply(lambda x:x*10,axis=1)
print(df1)
