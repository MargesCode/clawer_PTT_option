# -*- coding: utf-8 -*-
"""
Created on Wed Oct 11 14:15:57 2017

@author: Jo_Lin
"""

import requests
from bs4 import BeautifulSoup
import Wexcel
import json
#import csv
    

#======   Page Information ==================

ZF4_ZE4554KL='https://www.asus.com/zentalk/tw/forum-287-'    #初始網址 (ze554kl)
ZF4_ZS551KL='https://www.asus.com/zentalk/tw/forum-314-'    #初始網址 (zs551kl)
ZF4_ZC554KL='https://www.asus.com/zentalk/tw/forum-315-'    #初始網址 (zc554kl)
ZF4_ZD552KL='https://www.asus.com/zentalk/tw/forum-288-'    #初始網址 (zD552kl)
ZF3_ZE553KL='https://www.asus.com/zentalk/tw/forum-274-'    #ZE553KL(ZenFone3 Zoom)

ZF3_ZE552KL='https://www.asus.com/zentalk/tw/forum-247-'  #ZE552KL
ZF3_ZS570KL='https://www.asus.com/zentalk/tw/forum-246-'
ZF3_ZC553KL='https://www.asus.com/zentalk/tw/forum-264-'
ZF3_ZC551KL='https://www.asus.com/zentalk/tw/forum-260-'
ZF3_ZC520TL='https://www.asus.com/zentalk/tw/forum-263-'
ZF3_ZU680KL='https://www.asus.com/zentalk/tw/forum-255-'
ZF3_ZS550KL='https://www.asus.com/zentalk/tw/forum-265-'
ZF3_ZE520KL='https://www.asus.com/zentalk/tw/forum-248-'

NB_Zenbook='https://www.asus.com/zentalk/tw/forum-298-'
NB_AllNB='https://www.asus.com/zentalk/tw/forum-296-'
NB_Vivobook_N='https://www.asus.com/zentalk/tw/forum-302-'
NB_Vivobook_S='https://www.asus.com/zentalk/tw/forum-303-'
NB_Vivobook_K='https://www.asus.com/zentalk/tw/forum-304-'
NB_Vivobook_X='https://www.asus.com/zentalk/tw/forum-305-'
NB_Vivobook_E='https://www.asus.com/zentalk/tw/forum-306-'
NB_Vivobook_OTHERS='https://www.asus.com/zentalk/tw/forum-307-'

#urlPath='https://www.asus.com/zentalk/tw/forum-75-'  #粉絲閒聊

#pageInfo=[{"PhoneType":"ZenFone3","Model":"ZC553KL","url":ZF3_ZC553KL},{"PhoneType":"ZenFone3","Model":"ZC551KL","url":ZF3_ZC551KL},{"PhoneType":"ZenFone3","Model":"ZC520TL","url":ZF3_ZC520TL},{"PhoneType":"ZenFone3","Model":"ZU680KL","url":ZF3_ZU680KL},{"PhoneType":"ZenFone3","Model":"ZS550KL","url":ZF3_ZS550KL},{"PhoneType":"ZenFone3","Model":"ZE520KL","url":ZF3_ZE520KL},{"PhoneType":"ZenFone3","Model":"ZE553KL","url":ZF3_ZE553KL},{"PhoneType":"ZenFone4","Model":"ZE554KL","url":ZF4_ZE4554KL},{"PhoneType":"ZenFone4","Model":"ZS551KL","url":ZF4_ZS551KL},{"PhoneType":"ZenFone4","Model":"ZC554KL","url":ZF4_ZC554KL},{"PhoneType":"ZenFone4","Model":"ZD552KL","url":ZF4_ZD552KL},{"PhoneType":"ZenFone3","Model":"ZE552KL","url":ZF3_ZE552KL},{"PhoneType":"ZenFone3","Model":"ZS570KL","url":ZF3_ZS570KL},{"PhoneType":"NB","Model":"Zenbook","url":NB_Zenbook},{"PhoneType":"NB","Model":"Other NB","url":NB_AllNB},{"PhoneType":"NB","Model":"Vivobook_N","url":NB_Vivobook_N},{"PhoneType":"NB","Model":"Vivobook_S","url":NB_Vivobook_S},{"PhoneType":"NB","Model":"Vivobook_K","url":NB_Vivobook_K},{"PhoneType":"NB","Model":"Vivobook_X","url":NB_Vivobook_X},{"PhoneType":"NB","Model":"Vivobook_E","url":NB_Vivobook_E},{"PhoneType":"NB","Model":"Vivobook_OTHERS","url":NB_Vivobook_OTHERS},]
pageInfo=[{"PhoneType":"ZenFone3","Model":"ZE552KL","url":ZF3_ZE552KL},{"PhoneType":"ZenFone3","Model":"ZS570KL","url":ZF3_ZS570KL},{"PhoneType":"ZenFone3","Model":"ZC553KL","url":ZF3_ZC553KL},{"PhoneType":"ZenFone3","Model":"ZC551KL","url":ZF3_ZC551KL},{"PhoneType":"ZenFone3","Model":"ZC520TL","url":ZF3_ZC520TL},{"PhoneType":"ZenFone3","Model":"ZU680KL","url":ZF3_ZU680KL},{"PhoneType":"ZenFone3","Model":"ZS550KL","url":ZF3_ZS550KL},{"PhoneType":"ZenFone3","Model":"ZE520KL","url":ZF3_ZE520KL},{"PhoneType":"ZenFone3","Model":"ZE553KL","url":ZF3_ZE553KL},{"PhoneType":"ZenFone4","Model":"ZE554KL","url":ZF4_ZE4554KL},{"PhoneType":"ZenFone4","Model":"ZS551KL","url":ZF4_ZS551KL},{"PhoneType":"ZenFone4","Model":"ZC554KL","url":ZF4_ZC554KL},{"PhoneType":"ZenFone4","Model":"ZD552KL","url":ZF4_ZD552KL},{"PhoneType":"NB","Model":"Zenbook","url":NB_Zenbook},{"PhoneType":"NB","Model":"Other NB","url":NB_AllNB},{"PhoneType":"NB","Model":"Vivobook_N","url":NB_Vivobook_N},{"PhoneType":"NB","Model":"Vivobook_S","url":NB_Vivobook_S},{"PhoneType":"NB","Model":"Vivobook_K","url":NB_Vivobook_K},{"PhoneType":"NB","Model":"Vivobook_X","url":NB_Vivobook_X},{"PhoneType":"NB","Model":"Vivobook_E","url":NB_Vivobook_E},{"PhoneType":"NB","Model":"Vivobook_OTHERS","url":NB_Vivobook_OTHERS},]
#pageInfo=[{"PhoneType":"ZenFone3","Model":"ZC551KL","url":ZF3_ZC551KL},{"PhoneType":"ZenFone4","Model":"ZD552KL","url":ZF4_ZD552KL}]
pageInfo.sort(key=lambda d:d['PhoneType'])

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
    if(TP_Number_item > 0):
        TP_Number_Page=TP_Number_text[TP_Number_item].text #擷取總頁數的字串
        if ' ' in TP_Number_Page:
            TP_Number_Page_text=TP_Number_Page.split(' ')[1]   #處理總頁數的字串
        else:
            TP_Number_Page_text=TP_Number_Page 
        totalPage=int(TP_Number_Page_text)
    else:
        totalPage=1
#    print("Total Pages:",totalPage)         #計算總頁數 (數字)
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

#==========================  output data ========================
def outputToFile(list_Data):
    excelName="ZenTalkData_Test_1220.xls"
    wExcel=Wexcel.Wexcel(excelName)
    wb=wExcel.NewExcel()

    
    for cdata in list_Data:
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
            print("current Number:",dict_shName['Number'])
            for d in array_shName2:
                if d['shName']==shName:
                    sh=d['sh']
                    rowNumber=dict_shName['Number']      
        dictdata=cdata['Data']
#        print("Crow:",rowNumber)
        rNumber=wExcel.writeTofile(sh,wb,dictdata,rowNumber)
        dict_shName['Number'] = rNumber
        
#        print("return Row:",dict_shName['Number'])
#        dict_shName['Number']=dict_shName['Number']+rNumber
#        print("Rows:",rNumber)
        print("update row:",dict_shName['Number'])
        print(array_shName2)
        print("**********"+cdata['Model']+"結束 *********")
#============= write to json ============================================
def writeToJson(l_spam):
    jsonFileName='ZenTalkData_Test_1220.json'
    with open(jsonFileName, 'w') as jsFile:
        json.dump(l_spam, jsFile)


#===== JSON轉excel=====:

def jsonToExcel(spam):
#    columNum=1
    flg=True
    arr_phoneType=[]
    fileName='ZenTalkData_Test_1220.xls'
    wFile=Wexcel.Wexcel(fileName)
#    sh=wFile.openExcel()
    
    for cspam in spam:
        print(cspam['PhoneType'])
        if cspam['PhoneType'] not in arr_phoneType:
            flg=True
            arr_phoneType.append(cspam['PhoneType'])
            print(arr_phoneType)
            shName=cspam['PhoneType']
            dict_capName=['ID', 'Phone Model','Status','Summary']  
            sh=wFile.createSheet(shName,dict_capName)
            flg=False
            columNum=1
              
        columNum=wFile.jsonToExcel_Zentalk(sh,cspam['Data'],columNum)
    print("*** write to file completed....")
#=============  Main =====================

allData=[]
dict_allData={}
array_shName=[]
array_shName2=[]

array_comp1=[]

for pInfo in pageInfo:
    
    phoneType=pInfo['PhoneType']   #獲取PhoneType
    phoneModel=pInfo['Model']      #獲取PhoneModel       
    pageData=getPage(pInfo,1) #或取初始頁面
    
    totalPage=getTotalPage(pageData)
    print(phoneModel + " Total Page : " +str(totalPage))
    dict_data=getPageConent(pInfo,totalPage)
    dict_allData={}
    dict_allData['PhoneType']=pInfo['PhoneType']
    dict_allData['Model']=pInfo['Model'] 
    dict_allData['Data']=dict_data
#    dict_allData['Number']=
    
    allData.append(dict_allData)
    print("data length:",str(len(allData)))

writeToJson(allData)  #先寫入 JSOM
jsonToExcel(allData)
   
#outputToFile(allData)  #寫入Excel

#==========================


