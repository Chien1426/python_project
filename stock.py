import requests
import pandas as pd
import xlwings as xw
from tkinter import*

# Tk視窗參數設定
win = Tk()
win.title("Python Stock")
# win.geometry(長x寬 +X +Y)
win.geometry("400x220+800+400")
win.config(bg="#323232")

title_text = Label(text="Python Stock", fg="skyblue", bg="#323232")
# obj.config(font="字型 大小")
title_text.config(font="微軟正黑體 15")
title_text.pack()

# 日期輸入框
min_range = Label(text="Date(yyyymmdd)", fg="white", bg="#323232")
min_range.pack()
min_entry = Entry()
min_entry.pack()

x_show = Label(text="", fg="white", bg="#323232")
x_show.pack()


# 把csv檔抓下來
def stock_csv ():
    min_val = str(min_entry.get())
    url = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=csv&date="+min_val+"&type=ALL"
    res = requests.get(url)
    data = res.text
    if len(data) == 0:
        x_show.config(text="no data")
        return
    else:
        # 把爬下來的資料整理乾淨
        cleaned_data = []
        for da in data.split('\n'):
            if len(da.split('","')) == 16 and da.split('","')[0][0] != '=':
                cleaned_data.append([ele.replace('",\r','').replace('"','') 
                                    for ele in da.split('","')])
        
        # 輸出成表格並呈現在excel上
        df = pd.DataFrame(cleaned_data, columns = cleaned_data[0])
        df = df.set_index('證券代號')[1:]
        xw.view(df)

    


# 生成及複製按鈕
generate_btn = Button(text="Excel for Stock", command= stock_csv)
generate_btn.pack()

win.mainloop()

 
