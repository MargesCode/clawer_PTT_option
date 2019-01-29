# -*- coding: utf-8 -*-
"""
Created on Wed Nov 22 11:43:46 2017

@author: Jo_Lin
"""
#import xlrd
import xlwt
import json

class Wexcel:
    
    def  __init__(self,fileName):
        self.fileName=fileName
        self.wb=xlwt.Workbook()

        
        
    def createSheet(self,shName,capName):

        sh=self.wb.add_sheet(shName)
        columNum=0
        for capStr in capName:
            sh.write(0,columNum,capStr)
            columNum+=1
        return sh
            

            
    def writeTofileEntities(self):
        wb=xlwt.Workbook() 
        sh=wb.add_sheet('TestData')
        sh.write(0,0,"ID")
        sh.write(0,1,"Agent Name")
        sh.write(0,2,"Entities Name")
        sh.write(0,3,"Reference Value")
        sh.write(0,4,"Synonym")
        colnum=1
        for cd in self.dic_data:
            print(cd)
            sh.write(colnum,0,cd['ID'])
            sh.write(colnum,1,cd['AgentName'])
            sh.write(colnum,2,cd['EntitiesName'])
            sh.write(colnum,3,cd['ReferenceValue'])
            sh.write(colnum,4,cd['Synonym'])
            colnum+=1
        wb.save(self.fileName)
        
    def writeTofile(self):
        wb=xlwt.Workbook() 
        sh=wb.add_sheet(self.shName)
        cols=len(self.capName)-1
        colNum=0
#寫入Caption Name
        for col in self.capName:                #定義Caption Name
            sh.write(0,colNum,col)
            colNum+=1
        colsnum=1
#寫入資料        
        for cd in self.dic_data:
            colNum=0
            print(cd)
            for key,value in cd.items():
                sh.write(colsnum,colNum,value)
            colsnum+=1
            colNum+=1
        wb.save(self.fileName)
        
        
    def jsonToExcel(self,sh,dict_data):
        columNum=1
        for cd in dict_data:
            print(cd)
            sh.write(columNum,0,cd['ID'])
            sh.write(columNum,1,cd['SearchKeyWord'])
            sh.write(columNum,2,cd['Rate'])
            sh.write(columNum,3,cd['Date'])
            sh.write(columNum,4,cd['Title'])
            sh.write(columNum,5,cd['Author'])
            columNum+=1
        self.wb.save(self.fileName)

    def jsonToExcel_Zentalk(self,sh,dict_data,columNum):
#        columNum=1
        for cd in dict_data:
            print(cd)
            sh.write(columNum,0,cd['ID'])
            sh.write(columNum,1,cd['Phone Model'])
            sh.write(columNum,2,cd['Status'])
            sh.write(columNum,3,cd['Summary'])
            columNum+=1
        self.wb.save(self.fileName)
        return columNum
        
#        ['ID', 'SearchKeyWord','Rate','Date','Title','Author']
        
        

