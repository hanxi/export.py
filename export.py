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
import json
import pprint

KIND_NORMAL = "normal"
KIND_GLOBAL = "global"

TARGET_TYPE_PYTHON = ".py"
TARGET_TYPE_JSON   = ".json"
TARGET_TYPE_LUA    = ".lua"

def ToPy(pyDict):
    pyDictStr = pprint.pformat(pyDict, indent = 4)
    outStr = '''#encoding=utf-8
def GetData():
    return %s''' % pyDictStr
    return outStr

def ToJson(pyDict):
    return json.dumps(pyDict, sort_keys = True, indent = 4, ensure_ascii = False)

def _NewLine(count):
  return '\n' + '    ' * count

def _ToLua(out, obj, indent = 1):
    if isinstance(obj, int) or isinstance(obj, float) or isinstance(obj, str):
        out.append(json.dumps(obj, ensure_ascii = False))
    else:
        isList = isinstance(obj, list)
        out.append('{')
        isFirst = True
        for i in obj:
            if isFirst:
                isFirst = False
            else:
                out.append(',')
            out.append(_NewLine(indent))
            if not isList:
                k = i
                i = obj[k]
                out.append('[')
                if isinstance(k, int) or isinstance(k, float):
                    out.append(str(k))
                else:
                    out.append('"')
                    out.append(str(k))
                    out.append('"')
                out.append(']')
                out.append(' = ')
            _ToLua(out, i, indent + 1)
        out.append(_NewLine(indent - 1))
        out.append('}')

def ToLua(pyDict):
    out = []
    _ToLua(out, pyDict)
    luaStr = "".join(out)
    outStr = 'return %s' % luaStr
    return outStr


SUPPORT_TARGET_TYPE = {
    TARGET_TYPE_PYTHON: ToPy,
    TARGET_TYPE_JSON: ToJson,
    TARGET_TYPE_LUA: ToLua,
}


FORMAT_FUNC = {
    "str": lambda v: v,
    "int": lambda v: int(float(v)),
    "arrstr": lambda v: [ i.strip() for i in v.split(',') ],
    "arrint": lambda v: [ int(float(i.strip())) for i in v.split(',') ],
}

class Exporter:
    def __init__(self, context):
        self.pyDict = {}
        self.context = context
        # read xlsx
        self.ReadExcel()
        pass

    def ReadExcel(self):
        book = xlrd.open_workbook(self.context.excelFile)
        sheet = None
        try:
            sheet = book.sheet_by_name(self.context.sheetName)
        except:
            print("has no sheet named: %s" % self.context.sheetName)
            sys.exit(3)
        self.sheet = sheet

    def Export(self):
        transFunc = SUPPORT_TARGET_TYPE[self.context.targetType]
        # to str
        outStr = transFunc(self.pyDict)
        # save to file
        with open(self.context.targetFile, 'w') as f:
            f.write(outStr + "\n")
        pass

    def FormatValue(self, value, valueType, name):
        if valueType not in FORMAT_FUNC:
            print("In %s, Unknow valueType:%s,name:%s,value:%s" % (self.context.sheetName, valueType, name, value))
            sys.exit(5)

        formatFunc = FORMAT_FUNC[valueType]
        return formatFunc(value)

    def CheckValueType(self, valueType, name):
        if valueType in FORMAT_FUNC:
            return False
        return True

def IsOutTypeMatch(outTypeList, outType):
    if not outType:
        return True
    if outTypeList is None:
        return True
    if outType in outTypeList.split('/'):
        return True
    return False

class NormalExporter(Exporter):
    def __init__(self, context):
        super().__init__(context)
        self.context = context
        # to pyDict
        self.ToPyDict()

    def ToPyDict(self):
        self.pyDict = {}
        self.ReadHead()
        self.ReadBody()
        pass

    def ReadHead(self):
        # first 4 row is head
        if self.sheet.nrows < 4:
            print("Maybe not right sheet:%s" % self.context.sheetName)
            sys.exit(6)

        descriptions = self.sheet.row_values(0)
        valueTypes = self.sheet.row_values(1)
        names = self.sheet.row_values(2)
        outTypes = self.sheet.row_values(3)

        self.headInfos = []
        try:
            for colIndex in range(self.sheet.ncols):
                valueType = str(valueTypes[colIndex]).strip()
                name = str(names[colIndex]).strip()
                outTypeMatch = IsOutTypeMatch(str(outTypes[colIndex]).strip(), self.context.outType.strip())

                if self.CheckValueType(valueType, name):
                    print("In %s, Ignore unknow valueType:%s,name:%s" % (self.context.sheetName, valueType, name))
                    outTypeMatch = false

                self.headInfos.append((valueType, name, outTypeMatch))

        except Exception as e:
            e.args += ('%s has a title error, %s at %d column in %s' % (self.sheet.name, (valueType, name), colIndex + 1, self.context.excelFile) , '')
            raise e
        pass

    def GetSheetValue(self, row, colIndex):
        outTypeMatch = self.headInfos[colIndex][2]
        if not outTypeMatch:
            return None

        valueType = self.headInfos[colIndex][0]
        name = self.headInfos[colIndex][1]
        value = str(row[colIndex]).strip()

        if valueType and name and value:
            return self.FormatValue(value, valueType, name)
        return None

    def ReadBody(self):
        if not len(self.headInfos):
            return
        for i in self.headInfos:
            if not i[0] or not i[1]:
                return
            pass

        # body start row 5
        for rowIndex in range(4, self.sheet.nrows):
            outTypeMatch = self.headInfos[0][2]
            if not outTypeMatch:
                print("Don't export this sheet:%s for %s" % (self.context.sheetName, self.context.outType))
                sys.exit(0)

            row = self.sheet.row_values(rowIndex)
            firstText = str(row[0]).strip()
            if not firstText:
                # skip row
                print("%s skip line: %d" % (self.sheet.name, rowIndex + 1))
                continue

            key = self.GetSheetValue(row, 0)
            for colIndex in range(1, self.sheet.ncols):
                value = self.GetSheetValue(row, colIndex)
                if value is None:
                    continue

                if key not in self.pyDict:
                    self.pyDict[key] = {}

                name = self.headInfos[colIndex][1]
                self.pyDict[key][name] = value

# head macro for global exporter
K_NAME = "name"
K_VALUE = "value"
K_TYPE = "type"
K_OUT = "out"
K_DESCRIPTION = "description"
GLOBAL_HEAD_KEY = [
    K_NAME,
    K_VALUE,
    K_TYPE,
    K_OUT,
    K_DESCRIPTION
]

class GlobalExporter(Exporter):
    def __init__(self, context):
        super().__init__(context)
        self.context = context
        # to pyDict
        self.ToPyDict()

    def ToPyDict(self):
        self.pyDict = {}
        self.ReadHead()
        self.ReadBody()
        pass

    def ReadHead(self):
        # first 1 row is head
        if self.sheet.nrows < 1:
            print("Maybe not right sheet:%s" % self.context.sheetName)
            sys.exit(6)

        self.headInfo = {}
        headRow = self.sheet.row_values(0)
        for colIndex in range(self.sheet.ncols):
            value = str(headRow[colIndex])
            if value in GLOBAL_HEAD_KEY:
                self.headInfo[value] = colIndex
        pass

    def ReadBody(self):
        for key in GLOBAL_HEAD_KEY:
            if key not in self.headInfo:
                print("global exporter head not right in %s" % self.context.sheetName)
                sys.exit(8)
        pass

        # body start row 2
        for rowIndex in range(1, self.sheet.nrows):
            row = self.sheet.row_values(rowIndex)
            rowValue = {}
            for headKey in GLOBAL_HEAD_KEY:
                headIndex = self.headInfo[headKey]
                rowValue[headKey] = str(row[headIndex]).strip()
            if not IsOutTypeMatch(rowValue[K_OUT], self.context.outType):
                continue

            name = rowValue[K_NAME]
            if not name:
                continue

            valueType = rowValue[K_TYPE]
            if not valueType:
                continue

            value = self.FormatValue(rowValue[K_VALUE], valueType, name)
            self.pyDict[name] = value
        pass

def Export(context):
    if context.kind == KIND_GLOBAL:
        GlobalExporter(context).Export()
    elif context.kind == KIND_NORMAL:
        NormalExporter(context).Export()
    else:
        print("Not support kind: %s" % context.kind)
        sys.exit(4)
    print("%s => %s Export successful!" % (context.excelFile, context.targetFile))

class Context:
    def __init__(self):
        self.excelFile = ""
        self.sheetName = ""
        self.targetFile = ""
        self.targetType = TARGET_TYPE_PYTHON
        self.kind = KIND_NORMAL
        self.outType = ""

    def Usage(self):
        print('''Usage python3 export.py [OPTION...] -f excelfile -s sheetname -t targetfile
    Arguments
    -f      : input excel file.
    -s      : the excel file's sheet name.
    -t      : out file, support suffix: py,json,lua.

    Options
    -h      : print this help message and exit.
    -k      : kind of excel sheet, support: normal[default],global.
    -o      : export out type, support: server,client, default out all.

    Example
    python3 export.py -f hero.xlsx -s heroAttr -t hero.py

    https://github.com/hanxi/export.py''')

    def SetTargetType(self):
        filename, extension = os.path.splitext(self.targetFile)
        self.targetType = extension

if __name__ == '__main__':
    context = Context()

    try:
        opst, args = getopt.getopt(sys.argv[1:], 'f:s:t:hk:o:')
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
        elif op == '-k':
            context.kind = v
        elif op == '-o':
            context.outType = v

    if not context.excelFile:
        context.Usage()
        sys.exit(1)

    if context.targetType not in SUPPORT_TARGET_TYPE:
        print("Not support export target file type:'%s', only support: py,json,lua" % context.targetType)
        sys.exit(2)

    Export(context)

