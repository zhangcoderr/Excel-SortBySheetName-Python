import openpyxl
import os
import datetime
import re
import xlrd
from xlutils.copy import copy
#from openpyxl import load_workbook
from win32com.client import Dispatch


class SheetData:
    def __init__(self,sheetFullName,sheetKey,number,ws):
        self.sheetFullName=sheetFullName#101.表2-4措施项目工料机分析表
        self.sheetKey=sheetKey#2-4
        self.number=number#101
        self.ws_data=ws

def replace_xlsx(sheetname,sheetvalue):
    table = wb.sheet_by_name(sheetname)
    #ws2=wb2.create_sheet(sheetname)

    # for i in range(table.nrows):
    #     for j in range(table.ncols):
    #         ws2.cell(row=i + 1, column=j + 1, value=table.cell(i,j).value)

def move_sheet(move_sheet,before_sheet_index):
    # excel = Dispatch("Excel.Application")
    # excel.Visible = True
    # book = excel.Workbooks.Open(r'C:\Users\123\Desktop\ExcelSort\111.xls', False, True)
    # sheet = book.Worksheets('Sheet3')
    # print(sheet.Name)
    # sheet.Move(Before=book.Worksheets('Sheet1'))
    # book.Worksheets('Sheet2').Move(Before=book.Worksheets('Sheet3'))
    move=book.Worksheets(move_sheet)
    before=book.Worksheets(before_sheet_index)
    move.Move(Before=before)


def sort_sheet():
    #name_dic={}
    #sheets=[SheetData(1,1,1,1)]
    sheets=[]
    for sheetname in sheetnames:
        numbers = re.findall('^(.*?)\.', sheetname)
        names=re.findall('表(.*?)[\u4e00-\u9fa5]',sheetname)

        ws=wb.sheet_by_name(sheetname)
        sheet=None
        if( not names):
            sheet=SheetData(sheetname,None,numbers[0],ws)
        else:

            sheet=SheetData(sheetname,names[0],numbers[0],ws)


        sheets.append(sheet)
    for i in range(len(sheets)):
        for j in range(len(sheets)):
            if(i==j): continue
            sheetData1 = sheets[i]
            sheetData2 = sheets[j]
            if (sheetData1 != sheetData2):
                key = sheetData1.sheetKey
                keyNext = sheetData2.sheetKey
                if (not key or not keyNext): continue
                for index, rule in enumerate(sortRule):
                    if (rule == key):
                        index_1 = index
                        break
                for index, rule in enumerate(sortRule):
                    if (rule == keyNext):
                        index_2 = index
                        break
                if (keyNext > key):
                    temp = sheetData1
                    sheets[i] = sheets[j]
                    sheets[j] = temp
    #for s in sheets:
    #     print(s.number)
    #     print(s.sheetKey)
    #     print('--------')
    for i in range(len(sheets)):
        for j in range(len(sheets)):

            if(i>=j or sheets[i].sheetKey!=sheets[j].sheetKey):
                # print(sheetData1.number)
                # print(sheetData2.number)
                # print(sheetData1.sheetKey)
                # print(sheetData2.sheetKey)
                continue
            #print('i' + str(i) + 'j' + str(j))
            sheetData1 = sheets[i]
            sheetData2 = sheets[j]
            if(int(sheetData1.number)>int(sheetData2.number)):
                temp = sheetData1
                sheets[i] = sheets[j]
                sheets[j] = temp
    # for s in sheets:
    #     print(s.sheetFullName)

    return sheets
if __name__ == "__main__":
    os.chdir(r"C:\Users\123\Desktop\ExcelSort")

    # ----------
    excel= Dispatch("Excel.Application")
    excel.Visible = True
    book= excel.Workbooks.Open(r'C:\Users\123\Desktop\ExcelSort\all.xls',False,True)
    # sheet=book.Worksheets('Sheet1')
    # print(sheet.Name)
    # sheet.Move(Before=book.Worksheets('Sheet1'))
    # book.Worksheets('Sheet2').Move(Before=book.Worksheets('Sheet3'))
    #
    # ----------

    sortRule=['1-1', '1-1-1', '1-1-2', '1-2','1-3-A', '1-3-B','1-3-C',  '1-3-A-1', '1-4', '1-4-1', '1-4-2', '1-5', '1-6', '1-7',  '2-1','2-2', '2-3', '2-4','2-5', '2-7',]

    filename = 'all.xls'
    print('loading')

    wb=xlrd.open_workbook(filename)
    print('load done')
    output=copy(wb)


    sheetnames = wb.sheet_names()
    result= sort_sheet()

    #move_sheet(result[2].sheetFullName,1)
    index=1
    for i in range(len(result)):
        smallest=result[i]
        if(i==len(result)):
            break
        else:
            move_sheet(smallest.sheetFullName,index)
            index=index+1
    print('done!!!')


    # for sheetname in result:
    #     print(sheetname)
    #     replace_xlsx(sheetname,result[sheetname])
    # for sheetname in sheetnames:
    #     replace_xlsx(sheetname)
    #
    #
    # d= datetime.datetime.now().strftime('%d-%M-%S')
    # filename2='testResult'+d+'.xls'
    # output.save(filename2)
