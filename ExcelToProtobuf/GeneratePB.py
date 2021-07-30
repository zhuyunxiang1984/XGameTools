#!/usr/bin/python
# -*- coding: UTF-8 -*-

import openpyxl
import os
import sys
import shutil

"""
1.读取表格内容
2.解析字段名称
3.解析字段类型
4.解析导出模式
5.生成.proto文件
6.protoc生成应用类(python,C#)
7.import python应用类,导入表格内容
8.生成protobuf二进制文件
"""

ROW_PROP_NAMES = 0 #字段名称
ROW_PROP_TYPES = 1 #字段类型
ROW_PROP_MODES = 2 #导出到服务器,客户端,还是全都要
ROW_PROP_START = 3 #表内容开始

TAB_BLANK = " " * 4

OUTPUT_CSHARP_CODE = "./CodeCSharp"
OUTPUT_PYTHON_CODE = "./CodePython"
OUTPUT_PROTO_FILE = "./ProtoFile"
OUTPUT_PROTO_DATA = "./ProtoData"

if not os.path.exists(OUTPUT_CSHARP_CODE):
    os.makedirs(OUTPUT_CSHARP_CODE)
if not os.path.exists(OUTPUT_PYTHON_CODE):
    os.makedirs(OUTPUT_PYTHON_CODE)
if not os.path.exists(OUTPUT_PROTO_FILE):
    os.makedirs(OUTPUT_PROTO_FILE)
if not os.path.exists(OUTPUT_PROTO_DATA):
    os.makedirs(OUTPUT_PROTO_DATA)


_FrontColorCode = {
    "black" : 30,
    "red" : 31,
    "green" : 32,
    "yellow" : 33,
    "blue" : 34,
    "white" : 37,
}
_BackGColorCode = {
    "black" : 40,
    "red" : 41,
    "green" : 42,
    "yellow" : 43,
    "blue" : 44,
    "white" : 47,
}

def LogError(str):
    print("\033[{1};{2}m {0} \033[0m".format(str, 31, 40))

def LogPrint(str, front="white", backg="black"):
    color1 = _FrontColorCode[front]
    color2 = _BackGColorCode[backg]
    print("\033[{1};{2}m {0} \033[0m".format(str, color1, color2))

def GetPBName(name):
    return "PBConfig{0}".format(name)

#表格字段类
class Field:
    def __init__(self, name, type, mode):
        self.SetData(name, type, mode)
    def SetData(self, name, type, mode):
        self.Name = name
        self.Repeated = False
        self.Type = ''
        self.Client = False
        self.Server = False
        if not type is None:
            temps = type.split(" ")
            if len(temps) == 1:
                self.Type = temps[0]
            else:
                self.Repeated = temps[0] == "repeated"
                self.Type = temps[1]

        if not mode is None:
            self.Client = mode.find('c') != -1
            self.Server = mode.find('s') != -1

    def IsValid(self):
        return len(self.Type) > 0

    def ToString(self, blank=""):
        text = blank
        text = text + "{0} {1} ".format(self.Name, self.Type)
        if self.Client:
            text = text + "c"
        if self.Server:
            text = text + "s"
        if self.Repeated:
            text = text + " (repeated)"
            '''
        if len(self.ChildList) > 0:
            blank = blank + TAB_BLANK
            text = text + "\n"
            for child in self.ChildList:
                text = text + child.ToString(blank)'''
        return text


#生成protobuf
class PBExporter:
    def __init__(self):
        pass

    #导出Excel中指定的Sheet
    def Export(self, fileName, sheet):
        LogPrint("正在导出 ==={0} {1}=== ......".format(fileName, sheet.title), "yellow")
        #解析字段
        rowTuples = tuple(sheet.rows)
        fields = self.GenerateFieldData(rowTuples)
        #print(fields)
        if len(fields) < 1:
            LogError("导出失败! {0} {1}".format(fileName, sheet.title))
            return False
        self.GeneratePBDescFile(fileName, sheet, fields)
        self.GeneratePBCodeFile(fileName, sheet)
        self.GeneratePBDataFile(fileName, sheet, fields, rowTuples)
        LogPrint("导出完成!", "green")
        return True

    #生成字段数据
    def GenerateFieldData(self, rowTuples):
        fields = []
        if len(rowTuples) < ROW_PROP_START + 1:
            LogError("表格字段描述信息不足!")
            return fields
        try:
            rowNames = rowTuples[ROW_PROP_NAMES]
            rowTypes = rowTuples[ROW_PROP_TYPES]
            rowModes = rowTuples[ROW_PROP_MODES]
            for cellIndex in range(len(rowNames)):
                #print("{0}-{1}-{2}".format(rowNames[cellIndex].value, rowTypes[cellIndex].value, rowModes[cellIndex].value))
                name = rowNames[cellIndex].value
                type = rowTypes[cellIndex].value
                mode = rowModes[cellIndex].value
                fields.append(Field(name, type, mode))

        except BaseException as e:
            raise
        return fields

    #生成PB描述文件
    def GeneratePBDescFile(self, workbook, sheet, fields):
        content = []
        content.append("syntax = \"proto3\";\n\n")
        content.append("message {0}ConfigItem{{\n".format(GetPBName(sheet.title)))
        index = 1
        for field in fields:
            if field.IsValid():
                if field.Repeated:
                    content.append(TAB_BLANK + "repeated {0} {1} = {2};\n".format(field.Type, field.Name, index))
                else:
                    content.append(TAB_BLANK + "{0} {1} = {2};\n".format(field.Type, field.Name, index))
                index = index + 1
        content.append("}\n\n")

        content.append("message {0}Config{{\n".format(GetPBName(sheet.title)))
        content.append(TAB_BLANK + "map<int32, {0}ConfigItem> Items = 1;\n".format(GetPBName(sheet.title)))
        content.append("}")

        pb_file = open("{0}/{1}.proto".format(OUTPUT_PROTO_FILE, GetPBName(sheet.title)), "w+")
        pb_file.writelines(content)
        pb_file.close()

    #生成PB解析文件
    def GeneratePBCodeFile(self, workbook, sheet):
        # 将PB转换成py格式
        try :
            command = "protoc -I={1} --csharp_out={0} {1}/{2}.proto".format(OUTPUT_CSHARP_CODE, OUTPUT_PROTO_FILE, GetPBName(sheet.title))
            os.system(command)
            command = "protoc -I={1} --python_out={0} {1}/{2}.proto".format(OUTPUT_PYTHON_CODE, OUTPUT_PROTO_FILE, GetPBName(sheet.title))
            os.system(command)
        except BaseException as e:
            print("GeneratePBCodeFile failed! sheet:{0}".format(sheet.title))
            raise

    #生成PB数据文件
    def GeneratePBDataFile(self, fileName, sheet, fields, rowTuples):

        sys.path.append(os.getcwd() + "/" + OUTPUT_PYTHON_CODE)

        try :
            moduleName = "{0}_pb2".format(GetPBName(sheet.title))
            exec("from {0} import *".format(moduleName))
            module = sys.modules[moduleName]
            #print(module)

            confClass = getattr(module, GetPBName(sheet.title)+"Config")

            config = confClass()
            #print(dir(conf.Items))
            #print(dir(item))

            rowLen = len(rowTuples)
            for i in range(ROW_PROP_START, rowLen):
                cols = rowTuples[i]
                id = cols[0].value
                if id is None:
                    continue
                item = config.Items.get_or_create(id)
                setattr(item, "Id", id)
                for cellIndex in range(1, len(cols)):
                    if fields[cellIndex].IsValid():
                        if cols[cellIndex] is not None:
                            cellValue = cols[cellIndex].value
                        else:
                            cellValue = ''
                        self._WriteToItem(item, fields[cellIndex], cellValue)

            file = open("{0}/{1}.bytes".format(OUTPUT_PROTO_DATA, GetPBName(sheet.title)), "wb+")
            file.write(config.SerializeToString())
            file.close()
        except BaseException as e:
            print("GeneratePBDataFile failed! {0} {1}".format(fileName, sheet.title))
            raise

    #写入表格数据
    def _WriteToItem(self, item, field, cellValue):
        #print("field:{0} cell:{1}".format(field.Name, cellValue))
        cellValue = str(cellValue)
        if field.Repeated:
            splitStrs = cellValue.split(',')
            #print(dir(getattr(item,field.Name)))
            #print(strs)
            for splitStr in splitStrs:
                getattr(item, field.Name).append(self._ConvertValue(field, splitStr))
                #child = item.add()
                #setattr(child, field.Name, cellValue)
        else:
            setattr(item, field.Name, self._ConvertValue(field, cellValue))

    def _ConvertValue(self, field, value):
        if field.Type == "int32":
            return int(value)
        if field.Type == "float":
            return float(value)
        if field.Type == "string":
            return str(value)
        return value

def LoadConfig(fullPath):
    result = {}
    try:
        file = open(fullPath, "r")
        lines = file.readlines()

        for line in lines:
            temps = line.split(' ')
            excelName = temps[0].replace("\n","")
            result[excelName] = []
            for i in range(1, len(temps)):
                temp = temps[i].replace("\n","")
                if len(temp) > 0:
                    result[excelName].append(temp)
        file.close()
    except BaseException as e:
        pass

    return result

def Main():
    os.system("") #这句调用没有意义,是为了print染色正常显示

    #LogPrint("运行参数:" + sys.argv.__str__(), "blue", "white")

    argLen = len(sys.argv)
    outputCodeDir = ""
    outputDataDir = ""
    excelDir = ""
    configFilePath = ""

    if argLen > 1:
        outputCodeDir = os.path.abspath(sys.argv[1])
    if argLen > 2:
        outputDataDir = os.path.abspath(sys.argv[2])
    if argLen > 3:
        excelDir = os.path.abspath(sys.argv[3])
    if argLen > 4:
        configFilePath = os.path.abspath(sys.argv[4])

    LogPrint("代码输出路径:{0}".format(outputCodeDir), "blue", "white")
    LogPrint("数据输出路径:{0}".format(outputDataDir), "blue", "white")
    LogPrint("表格路径:{0}".format(excelDir), "blue", "white")
    LogPrint("配置文件:{0}".format(configFilePath), "blue", "white")

    #读取配置信息
    config = LoadConfig(configFilePath)
    needCheckConfig = len(config) > 0
    print(config)

    #导出PB描述文件,数据文件,代码文件
    sheetNames = []
    pbExporter = PBExporter()
    files = os.listdir(excelDir)
    for file in files:
        if not file.endswith(".xlsx") or file.startswith("~$"):
            continue
        fileFullPath = excelDir + "/" + file
        if needCheckConfig:
            if file in config:
                workbook = openpyxl.load_workbook(fileFullPath)
                needCheckSheet = len(config[file]) > 0
                for sheet in workbook.worksheets:
                    if not needCheckSheet or config[file].__contains__(sheet.title):
                        if pbExporter.Export(file, sheet):
                            sheetNames.append(sheet.title)
        else:
            workbook = openpyxl.load_workbook(fileFullPath)
            for sheet in workbook.worksheets:
                if pbExporter.Export(file, sheet):
                    sheetNames.append(sheet.title)

    #复制代码
    if not os.path.exists(outputCodeDir):
        os.makedirs(outputCodeDir)
    for sheetName in sheetNames:
        shutil.copyfile("{0}/{1}.cs".format(OUTPUT_CSHARP_CODE, GetPBName(sheetName)), "{0}/{1}.cs".format(outputCodeDir, GetPBName(sheetName)))

    #复制数据
    if not os.path.exists(outputDataDir):
        os.makedirs(outputDataDir)
    for sheetName in sheetNames:
        shutil.copyfile("{0}/{1}.bytes".format(OUTPUT_PROTO_DATA, GetPBName(sheetName)), "{0}/{1}.bytes".format(outputDataDir, GetPBName(sheetName)))

#入口
if __name__ == '__main__':
    Main()
