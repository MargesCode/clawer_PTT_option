# -*- coding: utf-8 -*-
"""
Created on Wed Oct 11 14:15:57 2017

@author: Jo_Lin
"""

import requests
from bs4 import BeautifulSoup
import Wexcel
#import csv
    

#from scrapy.spider import CrawlSpider,Rule
#urlPath='https://www.asus.com/zentalk/tw/forum-287-'    #初始網址 (ze554kl)
#urlPath='https://www.asus.com/zentalk/tw/forum-314-'    #初始網址 (zs551kl)
#urlPath='https://www.asus.com/zentalk/tw/forum-315-'    #初始網址 (zc554kl)
#urlPath='https://www.asus.com/zentalk/tw/forum-288-'    #初始網址 (zD552kl)
#urlPath='https://www.asus.com/zentalk/tw/forum-274-'    #ZE553KL(ZenFone3 Zoom)
#urlPath='https://www.asus.com/zentalk/tw/forum-75-'  #粉絲閒聊
#urlPath='https://www.asus.com/zentalk/tw/forum-247-'  #ZE552KL

#pageInfo=[{"PhoneType":"ZenFone3","Model":"ZE553KL","url":"https://www.asus.com/zentalk/tw/forum-274-"},{"PhoneType":"ZenFone4","Model":"ZE554KL","url":"https://www.asus.com/zentalk/tw/forum-287-"},{"PhoneType":"ZenFone4","Model":"ZS551KL","url":"https://www.asus.com/zentalk/tw/forum-314-"}]

#pageInfo=[{"PhoneType":"ZenFone4","Model":"ZD552KL","url":"https://www.asus.com/zentalk/tw/forum-288-"},{"PhoneType":"ZenFone3","Model":"ZE553KL","url":"https://www.asus.com/zentalk/tw/forum-274-"}]
pageInfo=[{"PhoneType":"ZenFone4","Model":"ZD552KL","url":"https://www.asus.com/zentalk/tw/forum-288-"},{"PhoneType":"ZenFone4","Model":"ZC554KL","url":"https://www.asus.com/zentalk/tw/forum-315-"}]



def getPage(pInfo,pageCount):
#    loop_Count = 1       #紀錄目前的頁數
    urlPath=pInfo['url']           #獲取網頁網址
    url=urlPath+str(pageCount)+'.html'    
    html=requests.get(url)                       #獲取頁面內容
    sp=BeautifulSoup(html.text, 'html.parser')
    return sp

##==========================================
#def writeFile(wstr):
#    f = open('ZE552_1KL.txt', 'a', encoding = 'UTF-8') 
#    f.write(wstr)
#    f.write("\r\n")
#    f.close()

#============= 擷取網頁 內容item act1 ================
#data_Lev1=sp.select("th")
#print("data1 length:",len(data_Lev1))  #找到第一層table所在的div
#
#for item in data_Lev1:
#    if len(item.select("em")) > 0:
#         print(item.select("em")[0].text)
#         print(item.select(".xst")[0].text)
#         print("===================")
         
#===================獲取頁數資訊=======================   

def getTotalPage(l_pageData):
    TP_Tag=l_pageData.select("#fd_page_top")  #取得上方資料列
    TP_Number_text=TP_Tag[0].find_all('a')  #取得上方資料列的項目數量(頁數)
    TP_Number_item=len(TP_Number_text)-2    
    TP_Number_Page=TP_Number_text[TP_Number_item].text #擷取總頁數的字串    
    if ' ' in TP_Number_Page:
        TP_Number_Page_text=TP_Number_Page.split(' ')[1]   #處理總頁數的字串
    else:
        TP_Number_Page_text=TP_Number_Page 
    totalPage=int(TP_Number_Page_text)
    print("Total Pages:",totalPage)         #計算總頁數 (數字)
    return totalPage

#============= loop ====================

def getPageConent(l_pageInfo,l_totalPage):
    array_Data=[]
    dict_data={}
    loop_Count = 1       #紀錄目前的頁數
    total=1              #紀錄頁數的變數

    sp=getPage(l_pageInfo,loop_Count)
    while loop_Count <= l_totalPage:
#        print("array_Data:",array_Data)
        if total>5000:                  #當資料大於5000則停止
            break;
        else:
            data_Lev1=sp.select("th")
            for item in data_Lev1:
                dict_data={}
#                print("array_Data--1:",array_Data)
                if len(item.select("em")) > 0:
#                    print(str(total) + ','+item.select("em")[0].text+','+item.select(".xst")[0].text)
                    dict_data['ID']=str(total)
                    dict_data['Phone Model']=pInfo['Model'] 
                    dict_data['Status']=item.select("em")[0].text
                    dict_data['Summary']=item.select(".xst")[0].text
                    array_Data.append(dict_data)
                    total=total+1
            loop_Count=loop_Count+1
            sp=getPage(l_pageInfo,loop_Count)

    return array_Data
#=============  Main =====================

allData=[]
dict_allData={}
array_shName=[]
array_shName2=[]

for pInfo in pageInfo:
    phoneType=pInfo['PhoneType']   #獲取PhoneType
    phoneModel=pInfo['Model']      #獲取PhoneModel
    pageData=getPage(pInfo,1) #或許初始頁面
    totalPage=getTotalPage(pageData)
    dict_data=getPageConent(pInfo,totalPage)
    
    dict_allData['PhoneType']=pInfo['PhoneType']
    dict_allData['Model']=pInfo['Model'] 
    dict_allData['Data']=dict_data
#    dict_allData['Number']=
    
    allData.append(dict_allData)
    print("data length:",str(len(allData)))

    
    excelName="ZenTalkData_Test.xls"
    wExcel=Wexcel.Wexcel(excelName)
    wb=wExcel.NewExcel()

    
    for cdata in allData:
        shName=cdata['PhoneType']
        capName=['ID','Phone Model','Status','Summary']
        if shName not in array_shName:
            print("New Sheet:",shName)
            dict_shName={}
            array_shName.append(shName)
            dict_shName['shName']=shName
            sh=wExcel.createSheet(wb,shName)
            dict_shName['sh']=sh
            dict_shName['Number']=1
            array_shName2.append(dict_shName)
            wExcel.createCap(sh,capName)
            rowNumber=1
        else:
            print("old Sheet:",shName)
            for d in array_shName2:
                if d['shName']==shName:
                    sh=d['sh']
                    rowNumber=dict_shName['Number']+1          
        dictdata=cdata['Data']
        print("Crow:",rowNumber)
        rNumber=wExcel.writeTofile(sh,wb,dictdata,rowNumber)
        dict_shName['Number'] += rNumber
        
        print("return Row:",dict_shName['Number'])
#        dict_shName['Number']=dict_shName['Number']+rNumber
#        print("Rows:",rNumber)
#        print("Rows2:",dict_shName['Number'])
