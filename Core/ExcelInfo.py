# -*- coding: utf-8 -*-
# 这段代码主要的功能是把excel表格转换成python可用的列表和字典数据
import sys
import xlrd #http://pypi.python.org/pypi/xlrd
import json
import time
from datetime import datetime
from xlrd import xldate_as_tuple

class NameFlag:
    master='~'
    main='*'
    Type='#'
    error='!'

class SplitFlag:
    flag1='|'
    flag2=':'
class SheetType: 

    # 普通表
    # 输出JSON ARRAY

    NORMAL = 0

    # 有主外键关系的主表
    # 输出JSON MAP

    MASTER = 1

    # 有主外键关系的附表
    # 输出JSON MAP

    SLAVE = 2

# 支持的数据类型
class DataType:
    Int= 'int'
    Float='float'
    STRING= 'str'
    BOOL= 'bool'
    DATE= 'date'
    Obj='obj'
    ARRAY= '[]'
    DIC= '{}'
    UNKNOWN= 'unknown'

class HeadSetting:
    name=''
    Type=DataType.UNKNOWN
    index=0
    param=None
    def __init__(self,name,Type,index):
        self.name=name
        self.Type=Type
        self.index=index

class SheetInfo:
    name=''
    sheetName=''
    Type = SheetType.NORMAL
    dataType=DataType.ARRAY
    idHead=HeadSetting('',DataType.UNKNOWN,0)
    masterHead=HeadSetting('',DataType.UNKNOWN,0)
    slaves= []
    head= []
    sheet={}
    table=[]
    masterCols=[]

    def __init__(self,name):
        self.name=name
        self.sheetName=name
        self.Type=SheetType.NORMAL
        self.dataType=DataType.ARRAY
        self.idHead=None
        self.masterHead=None
        self.slaves=[]
        self.head=[]
        self.sheet={}
        self.table=[]
        self.masterCols=[]



class ExcelInfo:
    sheetInfos={'':SheetInfo('')}
    headRow=0
    Round=True
    ignoreEmpty=True
    def __init__(self,excelName,hr,r,i):
        self.sheetInfos={}
        self.headRow=hr
        self.Round=r
        self.ignoreEmpty=i
        self.setupSheetInfos(xlrd.open_workbook(excelName))
        self.parseSheetInfos()
        pass

    #获取最终所有可用表单
    def FinalTable(self):
        table={}
        for key in self.sheetInfos:
            sheetInfo=self.sheetInfos[key]
            if sheetInfo.Type==SheetType.SLAVE:
                continue
            if sheetInfo.idHead.Type==DataType.DIC:
                table[sheetInfo.name]=sheetInfo.sheet
            else:
                table[sheetInfo.name]=sheetInfo.table
        return table
    #处理表单父子关系
    def parseSheetInfos(self):
        if self.sheetInfos==None:
            pass
        for sheetInfo in self.sheetInfos.values():
            if sheetInfo.masterHead==None or sheetInfo.masterHead.name not in self.sheetInfos:
                if sheetInfo.slaves!=None and len(sheetInfo.slaves)>0:
                    sheetInfo.Type=SheetType.MASTER
                continue

            if sheetInfo.dataType==DataType.DIC:
                for r in range(0,len(sheetInfo.table)):
                    masterRow=self.sheetInfos[sheetInfo.masterHead.name].sheet[sheetInfo.masterCols[r]]
                    if sheetInfo.name not in masterRow:
                        masterRow[sheetInfo.name]={}
                    idHead=sheetInfo.table[r][sheetInfo.idHead.name]
                    masterRow[sheetInfo.name][idHead]=sheetInfo.table[r]
            elif sheetInfo.dataType==DataType.Obj:
                for r in range(0,len(sheetInfo.table)):
                    idHead=sheetInfo.table[r][sheetInfo.idHead.name]
                    self.sheetInfos[sheetInfo.masterHead.name].sheet[sheetInfo.masterCols[r]][idHead]=sheetInfo.table[r]
            else:
                for r in range(0,len(sheetInfo.table)):
                    masterRow=self.sheetInfos[sheetInfo.masterHead.name].sheet[sheetInfo.masterCols[r]]
                    if sheetInfo.name not in masterRow:
                        masterRow[sheetInfo.name]=[]
                    masterRow[sheetInfo.name].append(sheetInfo.table[r])

            


    # 预处理表单数据
    def setupSheetInfos(self, workbook):
        self.sheetInfos = {}

        sheetNames=workbook.sheet_names()
        for index in range(len(sheetNames)):
            sheet_name=sheetNames[index]
            if sheet_name[0]==NameFlag.error:
                continue

            sheetInfo=SheetInfo(sheet_name)
            sheetInfo.sheetName=sheet_name


            if sheet_name.count(NameFlag.master) > 0:            
                pair=sheet_name.split(NameFlag.master)
                sheetInfo.name=pair[0].strip()
                sheetInfo.Type=SheetType.SLAVE
                sheetInfo.masterHead=HeadSetting(pair[1].strip(),DataType.UNKNOWN,0)
            if sheetInfo.name.count('#') > 0:
                pair=sheetInfo.name.split('#')
                sheetInfo.dataType=pair[1]
                sheetInfo.name=pair[0]
            self.sheetInfos[sheetInfo.name]=sheetInfo
        row=self.headRow - 1
        #初始化表单数据
        for sheetInfo in self.sheetInfos.values():

            if sheetInfo.Type==SheetType.SLAVE:
                self.sheetInfos[sheetInfo.masterHead.name].slaves.append(sheetInfo.name)
            sheet=workbook.sheet_by_name(sheetInfo.sheetName)
            cols = sheet.ncols

            
            for col in range(0,cols):
                cell=sheet.cell_value(row,col)
                ctype=sheet.cell(row,col).ctype
                if ctype==0 or cell[0]==NameFlag.error:
                    continue
                head_setting=HeadSetting(cell,DataType.UNKNOWN,col)
                if cell[0]==NameFlag.master:
                    if sheetInfo.masterHead!=None:
                        sheetInfo.masterHead.Type=cell[1:len(cell)]
                        sheetInfo.masterHead.index=col
                        continue

                elif cell[0]==NameFlag.main:
                    sheetInfo.idHead=head_setting
                    cell=cell[1:len(cell)]
                    head_setting.name=cell
                if cell.count(NameFlag.Type) != 0:

                    pair=cell.split(NameFlag.Type)

                    head_setting.name=pair[0].strip()
                    head_setting.Type=pair[1].strip()
                    if len(pair)>2:
                        head_setting.param=pair[2].strip()
                
                sheetInfo.head.append(head_setting)

            if sheetInfo.idHead==None and sheetInfo.head!=None and len(sheetInfo.head)>0:
                sheetInfo.idHead=sheetInfo.head[0]
            sheetInfo.table = []
            sheetInfo.sheet = {}
            for i_row in range(self.headRow,sheet.nrows):
                self.parseRow(sheet, i_row, sheetInfo)

        for sheetInfo in self.sheetInfos.values():
            if sheetInfo.Type!=SheetType.SLAVE and len(sheetInfo.slaves)>0:
                sheetInfo.Type=SheetType.MASTER

    def Int(self,num):
        if self.Round:
            num=float(num)
            if num>0:
                return int(num+0.5)
            else:
                return int(num-0.5)
        return int(float(num))
    def unknownValue(self,num,ctype):
        if ctype == 2 and num % 1 == 0:  # 如果是整形
                return self.Int(num)
        elif ctype == 3:
            # 转成datetime对象
            date = datetime(*xldate_as_tuple(num, 0))
            return date.strftime('%Y/%d/%m %H:%M:%S')
        elif ctype == 4:
            return False if num == 0 else True
        else:
            return num
    # 解析一行
    def parseRow(self,sheet,rowIndex, sheetInfo) :
        if sheetInfo.head==None or len(sheetInfo.head)==0:
            return
        result = {}
        if sheetInfo.masterHead!=None:
            cell =sheet.cell_value(rowIndex,sheetInfo.masterHead.index)
            #cell = row[sheetInfo.masterHead.index]
            sheetInfo.masterCols.append(cell)
        headIndex=0
        if sheetInfo.idHead!=None:
            headIndex=sheetInfo.idHead.index
        for i in range(0,len(sheetInfo.head)):
            name = sheetInfo.head[i].name
            if name[0]=='!':
                continue

            
            Type = sheetInfo.head[i].Type
            param=sheetInfo.head[i].param
            index= sheetInfo.head[i].index
            cell =sheet.cell_value(rowIndex,index)
            #cell = row[index]
            if cell == None:
                if self.ignoreEmpty==False:
                    result[name] = None
                continue
            #ctype： 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
            ctype=sheet.cell(rowIndex,index).ctype

            if Type==DataType.UNKNOWN: # number string boolean
                cell=self.unknownValue(cell,ctype)
                if cell=='' and self.ignoreEmpty:
                    continue
            elif  Type==DataType.DATE:
                # 转成datetime对象
                date = datetime(*xldate_as_tuple(cell, 0))
                cell = date.strftime('%Y/%d/%m %H:%M:%S')
            elif  Type==DataType.Int:
                if ctype == 2 : 
                    cell = self.Int(cell)
                else:
                    Warning("type error at [%s,%s] ,%s is not a Int"%(rowIndex,index,cell))
            elif  Type==DataType.Float:
                if ctype == 2:  # 如果是数值
                    if param!=None:
                        r=self.Int(param)
                        r=pow(10,r)
                        cell=cell*r
                        cell=self.Int(cell)
                        cell=cell/r                      
                else:
                    Warning("type error at [%s,%s] ,%s is not a Float"%(rowIndex,index,cell))
            elif Type==DataType.STRING:
                cell=str(cell)
                if cell=='':
                    continue
            elif  Type==DataType.BOOL:
                cell = False if cell == 0 else True
            elif Type==DataType.Obj:
                if index==headIndex:
                     result[name]=str(cell)
                     continue
                if ctype==1:
                    temp=cell.split(SplitFlag.flag1)
                    if len(temp)>0:
                        for value in temp:
                            tp=value.split(SplitFlag.flag2)
                            ##if len(tp)==0:
                            ##    result["数据格式不对"]=str(cell)
                            if len(tp)==1:
                                result[tp[0]]=""
                            elif len(tp)>1:
                                result[tp[0]]=cell[len(tp[0])+1:len(value)]
                continue
            elif  Type==DataType.ARRAY:
                if ctype==1:
                    if cell=='':
                        continue
                    temp=cell.split(SplitFlag.flag1)
                    if len(temp)>0:
                        cell=temp
                    else:
                        temp.append(cell)

                        cell=temp
                else:
                    temp=self.unknownValue(cell,ctype)
                    if temp=='':
                        continue
                    cell=[temp]                
            elif  Type==DataType.DIC:

                if index==headIndex:
                     result[name]=str(cell)
                     continue
                cl={}
                if ctype==1:
                    temp=cell.split(SplitFlag.flag1)
                    if len(temp)>0:
                        for value in temp:
                            tp=value.split(SplitFlag.flag2)
                            if len(tp)==1:
                                cl[tp[0]]=""
                            elif len(tp)>1:
                                cl[tp[0]]=tp[1]
                if len(cl)>0:
                    cell=cl
                ##else:
                ##    cell={"数据格式不对":cell}
            else :
                print("无法识别的类型:[%s,%s],%s,%s"%(rowIndex,index,cell,type(cell)))
            result[name] = cell

        sheetInfo.table.append(result)
        if sheetInfo.idHead!=None:
            if sheetInfo.idHead.name in result:
                sheetInfo.sheet[result[sheetInfo.idHead.name]]=result
            elif len(result)>0:
                sheetInfo.sheet[result[list(result)[0]]]=result



