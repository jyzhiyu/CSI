#!/usr/bin/env python
#  -*- coding:utf-8 -*-
# Created by Jyzhiyu on 2017-8-2
__version__ = '1.0'
__author__ = 'jyzhiyu@gmail.com'
__license__ = 'GPL'

#-----------------------------------------------------
# File: GetCsiMainData
# Date: 2017-8-2
# Description:
#-----------------------------------------------------

import datetime
import xlrd,xlwt
from xlutils.copy import copy
import requests
from html.parser import HTMLParser

CsiAddrHeader = "http://www.csindex.cn/zh-CN/downloads/industry-price-earnings-ratio?"

class ClassCsiJsonPayload:
    def __init__(self):
        self.type=''
        self.date=''

CsiJsonPayload=ClassCsiJsonPayload()

def ClassCsiJsonPayload2dict(std):
    return {
        'type': std.type,
        'date': std.date
    }

#Init the Buffer of datas
DataBuf = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]

maskhead = 0
mask = 0
class MyHTMLParser(HTMLParser):

    def handle_starttag(self, tag, attrs):
        global maskhead
        if tag == 'tr':
            maskhead = 1

    def handle_endtag(self, tag):
        global maskhead
        if tag == 'tr':
            maskhead = 0

    def handle_data(self, data):
        global maskhead, DataBuf, CsiJsonPayload,mask
        if maskhead > 0 and data.isspace() == False:
            #Mask the data
            if data == '上海A股':
                mask = 1;return
            if data == '深圳A股':
                mask = 2;return
            if data == '沪深A股':
                mask = 3;return
            if data == '深市主板':
                mask = 4;return
            if data == '中小板':
                mask = 5;return
            if data == '创业板':
                mask = 6;return

            if CsiJsonPayload.type == 'zy2' and mask==1:
                DataBuf[0]=data
            if CsiJsonPayload.type == 'zy3' and mask==1:
                DataBuf[1]=data
            if CsiJsonPayload.type == 'zy2' and mask==2:
                DataBuf[2]=data
            if CsiJsonPayload.type == 'zy3' and mask==2:
                DataBuf[3]=data
            if CsiJsonPayload.type == 'zy2' and mask==3:
                DataBuf[4]=data
            if CsiJsonPayload.type == 'zy3' and mask==3:
                DataBuf[5]=data
            if CsiJsonPayload.type == 'zy2' and mask==4:
                DataBuf[6]=data
            if CsiJsonPayload.type == 'zy3' and mask==4:
                DataBuf[7]=data
            if CsiJsonPayload.type == 'zy2' and mask==5:
                DataBuf[8]=data
            if CsiJsonPayload.type == 'zy3' and mask==5:
                DataBuf[9]=data
            if CsiJsonPayload.type == 'zy2' and mask==6:
                DataBuf[10]=data
            if CsiJsonPayload.type == 'zy3' and mask==6:
                DataBuf[11]=data

            mask=0
parser = MyHTMLParser()

#Main function
def main():
    global DataBuf
    #Setting the initial date
    ErrFlag=0

    StartDate = datetime.date(2012, 1, 1)
    StepDate = datetime.timedelta(days=1)
    EndDate = datetime.date.today()

    #Setting the excel file name for operation
    ExcelFileName = u'主要板块.xls'
    #If the Excel file is not exist, create a new one
    try:
        ReadExcelFileHandle = xlrd.open_workbook(ExcelFileName)
    except:
        WriteExcelFileHandle = xlwt.Workbook()
        WorkSheet = WriteExcelFileHandle.add_sheet(u'CSIOrinalData', cell_overwrite_ok=True)
        SpreadTitle = [u'日期',
                       u'上证PE', u'上证PB',
                       u'深证PE', u'深证PB',
                       u'沪深PE', u'沪深PB',
                       u'深市主板PE', u'深市主板PB',
                       u'中小板PE', u'中小板PB',
                       u'创业板PE', u'创业主板PB']
        for i in range(len(SpreadTitle)):
            WorkSheet.write(0, i, SpreadTitle[i])

        WriteExcelFileHandle.save(ExcelFileName)
        ReadExcelFileHandle = xlrd.open_workbook(ExcelFileName)

    StockOrigalSheet = ReadExcelFileHandle.sheet_by_index(0)  # Open default Sheet 1
    EndRowIndex = StockOrigalSheet.nrows #Get the end row index
    #print('EndRowIndex is',EndRowIndex)
    # Setting the date format
    DateFormat = xlwt.XFStyle()
    DateFormat.num_format_str = 'yyyy-mm-dd'
    try:
        TailDateInStockOrigalSheet = xlrd.xldate.xldate_as_datetime(
            StockOrigalSheet.cell(EndRowIndex-1, 0).value, 0).date()   #Get the Date in end row
        WorkDay = TailDateInStockOrigalSheet + StepDate
    except:
        WorkDay = StartDate

    if WorkDay >= EndDate:
        print('数据已经更新到今天,按任意键退出......')
        return 1
    else:
        WriteExcelFileHandle = copy(ReadExcelFileHandle)    #Transfor the ReadExcelFileHandle to WriteExcelFileHandle
        WriteSheet = WriteExcelFileHandle.get_sheet(0)
        NewRowIndex = StockOrigalSheet.nrows
        while (WorkDay <= EndDate):
            #Print message
            print('正在更新数据' + str(WorkDay))
            #Update CSI data
            CsiJsonPayload.date = str(WorkDay)
            CsiJsonPayload.type = 'zy2' #zy1 is Dynamic PE of CSI mainboard
            url = CsiAddrHeader + 'type=' + CsiJsonPayload.type + '&' + 'date=' + CsiJsonPayload.date
            while ErrFlag == 0:
                try:
                    CsiWebHtml = requests.get(url)
                    ErrFlag = 1
                except:
                    pass
            ErrFlag = 0
            parser.feed(CsiWebHtml.text)
            CsiJsonPayload.type = 'zy3' #zy3 is PE of CSI mainboard
            url = CsiAddrHeader + 'type=' + CsiJsonPayload.type + '&' + 'date=' + CsiJsonPayload.date
            while ErrFlag == 0:
                try:
                    CsiWebHtml = requests.get(url)
                    ErrFlag = 1
                except:
                    pass
            ErrFlag = 0
            parser.feed(CsiWebHtml.text)
            print(DataBuf)
            if  DataBuf[0] == 0 or DataBuf[0] == ' -- ':
                pass
            else:
                #Write data into worksheet
                WriteSheet.write(NewRowIndex, 0, WorkDay, DateFormat)
                for i in range(12):
                    WriteSheet.write(NewRowIndex, i+1, DataBuf[i])

                #update the new row index
                NewRowIndex = NewRowIndex + 1

            # init the databuffer to next operation
            DataBuf = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
            WriteExcelFileHandle.save(u'主要板块.xls')

            if WorkDay.weekday() == 4:
                WorkDay = WorkDay + StepDate + StepDate + StepDate  # Jump to next Monday
            else:
                WorkDay = WorkDay + StepDate

#Boot Segment
if __name__=='__main__':
    main()