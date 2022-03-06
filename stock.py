import requests
import pandas as pd
import io 
# 網址
site = "https://www.twse.com.tw/exchangeReport/MI_5MINS_INDEX?response=csv&date=20220304&_=1646536509325"

# 利用 requests 來跟遠端 server 索取資料
response = requests.get(site)

df = pd.read_csv(io.StringIO(response.text))
print (df.head())