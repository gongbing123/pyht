#!/usr/bin/python3
# _-*- coding:utf-8 -*-
from decimal import Decimal
import openpyxl,os,re
import win32com.client as win32

def save_as_xlsx(fname):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()

def zhuanhuan(sheet):
    data=[]
    for row in sheet.values:
        if row[1] is None or row[1]=="姓名" or row[1]=="户名" or len(row)<3 or row[0]=="序号":
            continue
        data.append(list(row)[0:5])
    l=[]
    for i in range(len(data)):
        for j in range(len(data[i])):
            if len(str(data[i][j])) == 18 and re.match(r'\d{17}.', data[i][j]):
                l.append(j)
        break
    for i in range(len(data)):
        if l!=[]:
            data[i].pop(l[0])
        if len(data[i])<6:
            for n in range(6-len(data[i])):
                data[i].append(",")
        else:
            for n in range(5,len(data[i])):
                data[i][-1].pop()
        for j in range(len(data[i])):
            if data[i][j] is None:
                data[i][j]=","
            data[i][j]=str(data[i][j]).replace(" ","").replace('\n',"")
            if j<6 and len(data[i][j])==18 and re.match(r'\d{17}.',data[i][j]):
                data[i][j]=","
                data[i][5]=data[i][j]
                for x in range(j,5):
                    data[i][x]=data[i][x+1]
            if j==0:
                data[i][j]=str(i+1).zfill(8)
            if j==3:
                data[i][3]=str(Decimal(data[i][3]).quantize(Decimal('0.00')))
            if j==4:
                data[i][4]=str(101)
            if j==6:
                data[i][6]=","
    return data

def fenli(file_path):
    (filepath, tempfilename) = os.path.split(file_path)
    (filename, extension) = os.path.splitext(tempfilename)
    return filename

def xieru(sheet,path):
    data=zhuanhuan(sheet)
    fn=fenli(path).strip()
    filename=sheet.title+".data"
    fi=os.path.join(os.path.abspath('.')+"\\df",fn)
    if not os.path.exists(fi):
        os.makedirs(fi)
    with open(os.path.join(fi,filename),'w') as file:
        for i in data:
            file.write(','.join(i))
            file.write('\n')

def folder(path,ex):
    xlfs = [x for x in os.listdir(path) if os.path.isfile(os.path.join(path, x))
            and os.path.splitext(os.path.join(path, x))[1] == ex]
    return xlfs

dfyuanbiao=os.path.abspath('.')+"\\dfyuanbiao"
if not os.path.exists(dfyuanbiao):
    os.mkdir(dfyuanbiao)
xls=folder(dfyuanbiao,'.xls')
for i in range(len(xls)):
    x=os.path.join(dfyuanbiao,xls[i])
    save_as_xlsx(x)
    os.remove(x)
xlfs =folder(dfyuanbiao,'.xlsx')
for n in range(len(xlfs)):
    xl=os.path.join(dfyuanbiao,xlfs[n])
    table=openpyxl.load_workbook(filename=xl,data_only=True)
    for sh in table:
        xieru(sh,xl)
