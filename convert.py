# -*- coding:utf-8 -*-

"""excel格式转换"""

# 配置文件和程序文件(或脚本文件)同名即可,扩展名为`.ini`
# 支持参数格式`-e excel路径 -x xml路径 -f fmt路径 -s`,`--excel=excel路径 --xml=xml路径 --fmt=fmt路径 --enableskip`
# 参数会覆盖配置文件的设置
# 如果在意换行符的问题,可以按以下代码修改
# f = open(outPath, 'wb') # 以'w'方式写文件,python会自动在换行符结尾按照系统默认换行符替换'\n',使用'b'二进制方式,则不会做任何替换
# f.write("字符串数据\n".encode("utf8")) # 将字符串转为二进制字符数组
# f.close() # 关闭写文件
# 导出exe命令`pyinstaller --clean -F -p xlrd convert.py`

import sys
import getopt
import os
import xlrd
import configparser
import datetime

if __name__ == "__main__":
    # 运行时间统计
    timeStart = datetime.datetime.now()

    # 参数解析
    try:
        options, args = getopt.getopt(sys.argv[1:], "e:x:f:s", ["excel=", "xml=", "fmt=", "enableskip"])
    except getopt.GetoptError:
        print("invalid options")
        exit(1)

    # 配置读取
    fileConfig = os.path.splitext(os.path.basename(sys.argv[0]))[0] + ".ini"
    pathConfig = os.path.join(os.path.dirname(sys.argv[0]), fileConfig)
    enableSkip = False
    if os.path.exists(pathConfig):
        config = configparser.ConfigParser()
        config.read(pathConfig)
        pathExcel = config.get("path", "excel")
        pathXML = config.get("path", "XML")
        pathFMT = config.get("path", "FMT")
        enableSkip = config.get("option", "enableSkip") == "1"
    for option in options:
        if (option[0] == "-e") or (option[0] == "--excel"):
            pathExcel = option[1]
        elif (option[0] == "-x") or (option[0] == "--xml"):
            pathXML = option[1]
        elif (option[0] == "-f") or (option[0] == "--fmt"):
            pathFMT = option[1]
        elif (option[0] == "-s") or (option[0] == "--enableskip"):
            enableSkip = True

    # 参数整理
    enableFMT = len(pathFMT) > 0
    if enableFMT:
        pathDir = os.path.dirname(pathFMT)
        if not os.path.exists(pathDir):
            os.mkdir(os.path.dirname(pathDir))
        # 删除原格式文件
        if os.path.exists(pathFMT):
            os.remove(pathFMT)

    # 功能输出
    print("功能:")
    print("从EXCEL目录:%s" % pathExcel)
    print("导出XML到目录:%s" % pathXML)
    if enableFMT:
        print("在目录[%s]生成格式文件" % pathFMT)
    else:
        print("不生成格式文件")
    if enableSkip:
        print("开启根据日期跳过文件功能")
    else:
        print("禁用根据日期跳过文件功能")

    # 获取到全部的excel文件
    allExcels=[]
    for maindir, subdir, fileNameList in os.walk(pathExcel):
        # print(maindir, subdir, fileNameList)
        for fileName in fileNameList:
            if (fileName[0] == "~") or (os.path.splitext(fileName)[1] != ".xlsx"):
                continue
            allExcels.append(os.path.join(maindir, fileName))

    # 解析excel
    fileCounter = 0
    fileExportCounter = 0
    for excelPath in allExcels:
        # 基本参数
        # print(excelPath)
        excel = xlrd.open_workbook(excelPath)
        sheet = excel.sheet_by_index(0)
        numRow = sheet.nrows
        if numRow < 5:
            continue

        # 数据头处理
        # 服务器有用数据整理
        enableColumns = {}
        enableTypeRow = sheet.row(3)
        for index, enableType in enumerate(enableTypeRow):
            if (enableType.value != "Server") and (enableType.value != "Both"):
                continue
            enableColumns[index] = sheet.row(4)[index].value
        if len(enableColumns) == 0:
            continue

        # 目录整理
        relPath = os.path.relpath(excelPath, pathExcel)
        relPathFull = os.path.splitext(relPath)[0]+".xml"
        outPath = os.path.join(pathXML, relPathFull)
        if not os.path.exists(os.path.dirname(outPath)):
            os.mkdir(os.path.dirname(outPath))

        # 新旧判断
        fileCounter += 1
        needExport = True
        if enableSkip:
            if os.path.exists(outPath):
                xlsxMTime = os.stat(excelPath).st_mtime
                xmlMTime = os.stat(outPath).st_mtime
                if xlsxMTime < xmlMTime:
                    needExport = False
        # print("%s\t[%s](%s)" % (fileCounter, needExport and "+" or "=", excelPath))

        if needExport:
            # 日志
            print("%s\t%s" % (fileCounter, excelPath))

            # 导出文件计数
            fileExportCounter += 1

            # 写数据文件
            # f = open(outPath, 'w', encoding='utf8')
            f = open(outPath, 'wb')

            # 第一行标头
            f.write('<?xml version="1.0" encoding="UTF-8"?>\n'.encode("utf8"))

            # 第二行注释
            f.write("<!-- ".encode("utf8"))
            for index, enableType in enumerate(enableTypeRow):
                if index not in enableColumns:
                    continue
                f.write(('%s=%s ' % (sheet.row(4)[index].value, sheet.row(2)[index].value)).encode("utf8"))
            f.write("-->\n".encode("utf8"))

            # 写数据
            f.write("<root>\n".encode("utf8"))
            for i in range(5, numRow):
                row = sheet.row(i)
                isEmptyRow = True
                for index, cell in enumerate(row):
                    if index not in enableColumns:
                        continue
                    if (cell.ctype != xlrd.XL_CELL_EMPTY) and (cell.value != "") and (cell.value != 0):
                        isEmptyRow = False
                        break
                if isEmptyRow:
                    continue
                f.write("\t<data ".encode("utf8"))
                for index, cell in enumerate(row):
                    key = enableColumns.get(index)
                    if not key:
                        continue
                    value = cell.value
                    if cell.ctype == xlrd.XL_CELL_NUMBER: # 数字的特殊处理(excel中没有int,只有float)
                        if value == int(value): # "x.0"的处理
                            value = int(value)
                        else:
                            pt5char = str(value).split(".", 1)[1][:5] # 小数点最近的5个字符如果是0或9,则视为int
                            if pt5char == "00000": # "x.00000"的处理
                                value = int(value)
                            elif pt5char == "99999": # "x.99999"的处理
                                value = int(value) + 1
                    elif cell.ctype == xlrd.XL_CELL_TEXT: # 特殊字符转义
                        value = value.replace("&", "&amp;") # 这个转义要放在前面(因为会将后面转义的&替换为该转义)
                        value = value.replace("<", "&lt;")
                        value = value.replace(">", "&gt;")
                        value = value.replace("'", "&apos;")
                        value = value.replace("\"", "&quot;")
                    f.write(('%s="%s" ' % (key, value)).encode("utf8"))
                f.write("/>\n".encode("utf8"))
            f.write("</root>\n".encode("utf8"))

            # 关闭数据文件
            f.close()

        # 格式文件
        if enableFMT:
            # 写格式文件
            # f = open(pathFMT, 'a', encoding='utf8')
            f = open(pathFMT, 'ab')

            f.write(("%s\n" % relPathFull).encode("utf8"))
            for index, enableType in enumerate(enableTypeRow):
                if index not in enableColumns:
                    continue
                name = sheet.row(4)[index].value
                comment = sheet.row(2)[index].value
                ctype = sheet.row(1)[index].value
                if sheet.row(1)[index].value == "int":
                    ctype = "int64"
                f.write(("\t%s\t%s\t`xml:\"%s,attr\"`\t// %s\n" % (name, ctype, name, comment)).encode("utf8"))
            f.write("\n".encode("utf8"))

            # 关闭格式文件
            f.close()

    timeEnd = datetime.datetime.now()
    print("done, use [%s] seconds. [%s/%s] file converted." % ((timeEnd - timeStart).seconds, fileExportCounter, fileCounter))


# 测试代码
# print(allExcels)
# print("______________")
#
# # 读取excel
# excel = xlrd.open_workbook(os.path.join(pathExcel, "test.xlsx"))
# sheet = excel.sheet_by_index(0)
# numRow = sheet.nrows
# for i in range(1,numRow):
#     row = sheet.row(i)
#     print(row[3])
#     if row[3].ctype == xlrd.XL_CELL_NUMBER:
#         print(111111111, row[3].value == int(row[3].value))
#
# # 写文件
# f = open(os.path.join(pathXML, "a.txt"), 'w')
# f.write("aaa")
# f.write("bbb")
# f.write("ccc")
# f.close()
