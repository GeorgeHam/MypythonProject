#!usr/bin/python3
# -I Excel文件路径或文件夹
# -C 删除第C列
# -R 删除第R行
import os
from openpyxl import *
import sys
import argparse

def DeleteExcelItems(filename,col=0,row=0):
    if not(os.path.exists(filename)):
         return
    wb = load_workbook(filename)
    ws = wb.active
    if(col>0):
        ws.delete_cols(col) #删除第 col 列数据
        print("Delete col %d successful"%col)       
    if(row>0):
        ws.delete_rows(row) #删除第 row 行数据
        print("Delete row %d successful"%row)
    wb.save(filename)

def GetFiles():
    parse=argparse.ArgumentParser()
    parse.add_argument("-I","--items",type=str,help="请输入excel文件路径或文件夹")
    parse.add_argument("-C","--column",type=int,help="请输入要删除的列数")
    parse.add_argument("-R","--row",type=int,help="请输入要删除的行数")
    args=parse.parse_args()
    input=args.items
    col=args.column
    row=args.row
    print("删除 %s 的第 %d 列、第 %d 行"%(input,col,row))
    excelsfiles=[]
    if(os.path.isdir(input)):
        for root, dirs, files in os.walk(input):
            for f in files:
                if os.path.splitext(f)[1] == '.xlsx':
                    print("待处理文件 %s"%f)
                    excelsfiles.append(os.path.join(root, f))
    else:
        if(input.endswith('.xlsx')):
            excelsfiles.append(input)
            print("待处理文件 %s"%input)
    return excelsfiles,col,row



excelsfiles,col,row=GetFiles()
if(len(excelsfiles)==0):
    exit
for r in excelsfiles:
    print(r)
    DeleteExcelItems(r,col,row)
