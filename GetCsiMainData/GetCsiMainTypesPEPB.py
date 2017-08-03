#!/usr/bin/env python
#  -*- coding:utf-8 -*-
# Created by Jyzhiyu on 2017-8-2
__version__ = '1.0.1'
__author__ = 'jyzhiyu@gmail.com'
__license__ = 'GPL'

#-----------------------------------------------------
# File: GetCsiMainTypesData
# Date: 2017-8-3
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
DataBuf = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,]

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
            if data == '00':
                mask = 10;return
            if data == '01':
                mask = 11;return
            if data == '02':
                mask = 12;return
            if data == '03':
                mask = 13;return
            if data == '04':
                mask = 14;return
            if data == '05':
                mask = 15;return
            if data == '06':
                mask = 16;return
            if data == '07':
                mask = 17;return
            if data == '08':
                mask = 18;return
            if data == '09':
                mask = 19;return

            if data == '能源':
                mask = mask + 10;return
            if data == '原材料':
                mask = mask + 10;return
            if data == '工业':
                mask = mask + 10;return
            if data == '可选消费':
                mask = mask + 10;return
            if data == '主要消费':
                mask = mask + 10;return
            if data == '医药卫生':
                mask = mask + 10;return
            if data == '金融地产':
                mask = mask + 10;return
            if data == '信息技术':
                mask = mask + 10;return
            if data == '电信业务':
                mask = mask + 10;return
            if data == '公用事业':
                mask = mask + 10;return

            if CsiJsonPayload.type == 'zz2' and mask==20:
                DataBuf[0]=data;mask=0;return
            if CsiJsonPayload.type == 'zz3' and mask==20:
                DataBuf[1]=data;mask=0;return
            if CsiJsonPayload.type == 'zz2' and mask==21:
                DataBuf[2]=data;mask=0;return
            if CsiJsonPayload.type == 'zz3' and mask==21:
                DataBuf[3]=data;mask=0;return
            if CsiJsonPayload.type == 'zz2' and mask==22:
                DataBuf[4]=data;mask=0;return
            if CsiJsonPayload.type == 'zz3' and mask==22:
                DataBuf[5]=data;mask=0;return
            if CsiJsonPayload.type == 'zz2' and mask==23:
                DataBuf[6]=data;mask=0;return
            if CsiJsonPayload.type == 'zz3' and mask==23:
                DataBuf[7]=data;mask=0;return
            if CsiJsonPayload.type == 'zz2' and mask==24:
                DataBuf[8]=data;mask=0;return
            if CsiJsonPayload.type == 'zz3' and mask==24:
                DataBuf[9]=data;mask=0;return
            if CsiJsonPayload.type == 'zz2' and mask==25:
                DataBuf[10]=data;mask=0;return
            if CsiJsonPayload.type == 'zz3' and mask==25:
                DataBuf[11]=data;mask=0;return
            if CsiJsonPayload.type == 'zz2' and mask==26:
                DataBuf[12]=data;mask=0;return
            if CsiJsonPayload.type == 'zz3' and mask==26:
                DataBuf[13]=data;mask=0;return
            if CsiJsonPayload.type == 'zz2' and mask==27:
                DataBuf[14]=data;mask=0;return
            if CsiJsonPayload.type == 'zz3' and mask==27:
                DataBuf[15]=data;mask=0;return
            if CsiJsonPayload.type == 'zz2' and mask==28:
                DataBuf[16]=data;mask=0;return
            if CsiJsonPayload.type == 'zz3' and mask==28:
                DataBuf[17]=data;mask=0;return
            if CsiJsonPayload.type == 'zz2' and mask==29:
                DataBuf[18]=data;mask=0;return
            if CsiJsonPayload.type == 'zz3' and mask==29:
                DataBuf[19] = data;mask=0;return

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
    ExcelFileName = u'MainTypesData.xls'
    #If the Excel file is not exist, create a new one
    try:
        ReadExcelFileHandle = xlrd.open_workbook(ExcelFileName)
    except:
        WriteExcelFileHandle = xlwt.Workbook()
        WorkSheet = WriteExcelFileHandle.add_sheet(u'CSIOrinalData', cell_overwrite_ok=True)
        SpreadTitle = [ u'日期',
                        u'能源PE', u'能源PB',
                        u'原材料PE', u'原材料PB',
                        u'工业PE', u'工业PB',
                        u'可选消费PE', u'可选消费PB',
                        u'主要消费PE', u'主要消费PB',
                        u'医药卫生PE', u'医药卫生PB',
                        u'金融地产PE', u'金融地产PB',
                        u'信息技术PE', u'信息技术PB',
                        u'电信业务PE', u'电信业务PB',
                        u'公用事业PE', u'公用事业PB',]
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
            CsiJsonPayload.type = 'zz2' #zy1 is Dynamic PE of CSI mainboard
            url = CsiAddrHeader + 'type=' + CsiJsonPayload.type + '&' + 'date=' + CsiJsonPayload.date
            while ErrFlag == 0:
                try:
                    CsiWebHtml = requests.get(url)
                    ErrFlag = 1
                except:
                    pass
            ErrFlag = 0
            parser.feed(CsiWebHtml.text)
            CsiJsonPayload.type = 'zz3' #zz3 is PB of CSI mainboard
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
                for i in range(len(DataBuf)):
                    WriteSheet.write(NewRowIndex, i+1, DataBuf[i])

                #update the new row index
                NewRowIndex = NewRowIndex + 1

            # init the databuffer to next operation
            DataBuf = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
            WriteExcelFileHandle.save(ExcelFileName)

            if WorkDay.weekday() == 4:
                WorkDay = WorkDay + StepDate + StepDate + StepDate  # Jump to next Monday
            else:
                WorkDay = WorkDay + StepDate

#Boot Segment
if __name__=='__main__':
    main()