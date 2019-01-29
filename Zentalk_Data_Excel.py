# -*- coding: utf-8 -*-
"""
Created on Wed Oct 11 14:15:57 2017

@author: Jo_Lin
"""

import requests
from bs4 import BeautifulSoup
#import csv
    

#from scrapy.spider import CrawlSpider,Rule
#urlPath='https://www.asus.com/zentalk/tw/forum-287-'    #初始網址 (ze554kl)
#urlPath='https://www.asus.com/zentalk/tw/forum-274-'    #ZE553KL(ZenFone3 Zoom)
urlPath='https://www.asus.com/zentalk/tw/forum-75-'  #粉絲閒聊
urlPath='https://www.asus.com/zentalk/tw/forum-247-'  #ZE552KL
loop_Count = 1       #紀錄目前的頁數
total=1              #紀錄頁數的變數
url=urlPath+str(loop_Count)+'.html'

html=requests.get(url)                       #獲取頁面內容
sp=BeautifulSoup(html.text, 'html.parser')

#==========================================
def writeFile(wstr):
    f = open('ZE552_1KL.txt', 'a', encoding = 'UTF-8') 
    f.write(wstr)
    f.write("\r\n")
    f.close()

#============= 擷取網頁 內容item act1 ================
#data_Lev1=sp.select("th")
#print("data1 length:",len(data_Lev1))  #找到第一層table所在的div
#
#for item in data_Lev1:
#    if len(item.select("em")) > 0:
#         print(item.select("em")[0].text)
#         print(item.select(".xst")[0].text)
#         print("===================")
         
#===================獲取頁面資訊=======================   


TP_Tag=sp.select("#fd_page_top")  #取得上方資料列
#print("cp_Total Page : ",len(totalPage))
TP_Number_text=TP_Tag[0].find_all('a')  #取得上方資料列的項目數量(頁數)
#print("cp_page number: ",len(totalPage_number))
print(len(TP_Number_text))
TP_Number_item=len(TP_Number_text)-2

TP_Number_Page=TP_Number_text[TP_Number_item].text #擷取總頁數的字串
TP_Number_Page_text=TP_Number_Page.split(' ')[1]   #處理總頁數的字串
TP_Number_Page_Number=int(TP_Number_Page_text)
print("total Page:",TP_Number_Page_Number)         #計算總頁數 (數字)

#============= loop ====================

#=====原先寫入txt的方法====

#while loop_Count <= TP_Number_Page_Number:
#    if total>5000:
#        break;
#    else:
#        data_Lev1=sp.select("th")
#        for item in data_Lev1:
#            if len(item.select("em")) > 0:
#                print(str(total) + ','+item.select("em")[0].text+','+item.select(".xst")[0].text)
#                mystr=str(total) + ','+item.select("em")[0].text+','+item.select(".xst")[0].text
#                writeFile(mystr)
#                total=total+1
#        loop_Count=loop_Count+1
#        url=url=urlPath+str(loop_Count)+'.html'
#        html=requests.get(url)                       #獲取頁面內容
#        sp=BeautifulSoup(html.text, 'html.parser')

#========= 修改成 Excel的方法
dict_ZT=[]

while loop_Count <= TP_Number_Page_Number:
    if total>5000:
        break;
    else:
        data_Lev1=sp.select("th")
        for item in data_Lev1:
            if len(item.select("em")) > 0:
#                print(str(total) + ','+item.select("em")[0].text+','+item.select(".xst")[0].text)
                ZTData_text={}
                ZTData_text['ID']=str(total) 
                ZTData_text['Status']=item.select("em")[0].text
                ZTData_text['Title']=item.select(".xst")[0].text
#                mystr=str(total) + ','+item.select("em")[0].text+','+item.select(".xst")[0].text
                dict_ZT.append(ZTData_text)
                
                
                
#                writetoExcel(mystr)
                total=total+1
        loop_Count=loop_Count+1
        url=urlPath+str(loop_Count)+'.html'
        html=requests.get(url)                       #獲取頁面內容
        sp=BeautifulSoup(html.text, 'html.parser')
print(dict_ZT)
