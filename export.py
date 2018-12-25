#encoding=utf-8

'''
Copyright (c) 2018 hanxi(hanxi.info@gmail.com)

https://github.com/hanxi/export.py

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'''

import sys
if sys.version_info < (3, 0):
    print('python version need more than 3.x')
    sys.exit(1)

import os
import getopt
import xlrd

KIND_NORMAL = "normal"
KIND_GLOBAL = "global"

TARGET_TYPE_PYTHON = ".py"
TARGET_TYPE_JSON   = ".json"
TARGET_TYPE_LUA    = ".lua"
TARGET_TYPE_CSHARP = ".cs"

def ToPy(pyDict):
    return "ToPy"

def ToJson(pyDict):
    return "ToJson"

def ToLua(pyDict):
    return "ToLua"

def ToCSharp(pyDict):
    return "ToCSharp"

SUPPORT_TARGET_TYPE = {
    TARGET_TYPE_PYTHON: ToPy,
    TARGET_TYPE_JSON: ToJson,
    TARGET_TYPE_LUA: ToLua,
    TARGET_TYPE_CSHARP: ToCSharp,
}

class Exporter:
    def __init__(self, context):
        self.pyDict = {}
        self.outStr = ""
        self.context = context
        # read xlsx
        self.ReadExcel()
        pass

    def ReadExcel(self):
        data = xlrd.open_workbook(self.context.excelFile)
        # get sheet by name?
        print(data)

    def Export(self):
        transFunc = SUPPORT_TARGET_TYPE[self.context.targetType]
        self.outStr = transFunc(self.pyDict)
        # to str
        # save to file
        print("export...", self.outStr)
        pass

class NormalExport(Exporter):
    def __init__(self, context):
        self.context = context
        # to pyDict
        print("NormalExport")

class GlobalExport(Exporter):
    def __init__(self, context):
        self.context = context
        # to pyDict
        print("GlobalExport")

def Export(context):
    Exporter(context).Export()
    print("export finsish successful!!!")

class Context:
    def __init__(self):
        self.excelFile = ""
        self.sheetName = ""
        self.targetFile = ""
        self.targetType = TARGET_TYPE_PYTHON
        self.kind = KIND_GLOBAL

    def Usage(self):
        print('''Usage python3 export.py [OPTION...] -f excelfile -s sheetname -t targetfile
    Arguments
    -f      : input excel file
    -s      : the excel file's sheet name
    -t      : out file, support suffix: py,json,lua,cs

    Options
    -h      : print this help message and exit
    -k      : kind of excel sheet, support: normal[default],global

    Example
    python3 export.py -f hero.xlsx -s heroAttr -t hero.py

    https://github.com/hanxi/export.py''')

    def SetTargetType(self):
        filename, extension = os.path.splitext(self.targetFile)
        self.targetType = extension

if __name__ == '__main__':
    context = Context()

    try:
        opst, args = getopt.getopt(sys.argv[1:], 'f:s:t:hk:')
    except:
        context.Usage()
        sys.exit(1)

    for op,v in opst:
        if op == '-f':
            context.excelFile = v
        elif op == '-s':
            context.sheetName = v
        elif op == '-t':
            context.targetFile = v
            context.SetTargetType()
        elif op == '-h':
            context.Usage()
            sys.exit(0)

    if context.targetType not in SUPPORT_TARGET_TYPE:
        print("Not support export target file type, only support: py,json,lua,cs")
        context.Usage()
        sys.exit(2)

    Export(context)

