# -*- coding:utf-8 -*-

# excel格式转换

import sys
import os
import xlrd
import configparser

fileConfig="convert.ini"
# pathExcel="D:/pan/test_python/xlsx/data/src"
# pathXML="D:/pan/test_python/xlsx/data/tar"
# pathFMT="D:/pan/test_python/xlsx/data/tar/fmt.txt"

if __name__ == "__main__":
	# 配置读取
	pathConfig = os.path.join(os.path.dirname(sys.argv[0]), fileConfig)
	config = configparser.ConfigParser()
	config.read(pathConfig)
	pathExcel = config.get("path", "excel")
	pathXML = config.get("path", "XML")
	pathFMT = config.get("path", "FMT")

	# 参数整理
	enableFMT = len(pathFMT) > 0
	if enableFMT:
		pathDir = os.path.dirname(pathFMT)
		if not os.path.exists(pathDir):
			os.mkdir(os.path.dirname(pathDir))
		# 删除原格式文件
		if os.path.exists(pathFMT):
			os.remove(pathFMT)


	# 获取到全部的excel文件
	allExcels=[]
	for maindir, subdir, fileNameList in os.walk(pathExcel):
		# print(maindir, subdir, fileNameList)
		for fileName in fileNameList:
			if (fileName[0] != "~") or (os.path.splitext(fileName)[1] != ".xlsx"):
				allExcels.append(os.path.join(maindir, fileName))

	# 解析excel
	for excelPath in allExcels:
		print(excelPath)
		excel = xlrd.open_workbook(excelPath)
		sheet = excel.sheet_by_index(0)
		numRow = sheet.nrows
		if numRow < 6:
			continue

		# 数据头处理
		# 服务器有用数据整理
		enableColumns = {}
		enableTypeRow = sheet.row(3)
		for index, enableType in enumerate(enableTypeRow):
			if enableType.value == "Client":
				continue
			enableColumns[index] = sheet.row(4)[index].value
		if len(enableColumns) == 0:
			continue

		# 写数据文件
		relPath = os.path.relpath(excelPath, pathExcel)
		relPathFull = os.path.splitext(relPath)[0]+".xml"
		outPath = os.path.join(pathXML, relPathFull)
		if not os.path.exists(os.path.dirname(outPath)):
			os.mkdir(os.path.dirname(outPath))
		f = open(outPath, 'w', encoding='utf8')

		# 第一行标头
		f.write('<?xml version="1.0" encoding="UTF-8"?>\n')

		# 第二行注释
		f.write("<!-- ")
		for index, enableType in enumerate(enableTypeRow):
			if index not in enableColumns:
				continue
			f.write('%s=%s ' % (sheet.row(4)[index].value, sheet.row(2)[index].value))
		f.write("-->\n")

		# 写数据
		f.write("<root>\n")
		for i in range(5, numRow):
			row = sheet.row(i)
			f.write("\t<data ")
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
				f.write('%s="%s" ' % (key, value))
			f.write("/>\n")
		f.write("</root>\n")

		# 关闭数据文件
		f.close()

		# 格式文件
		if enableFMT:
			# 写格式文件
			f = open(pathFMT, 'a', encoding='utf8')

			f.write("%s\n" % relPathFull)
			for index, enableType in enumerate(enableTypeRow):
				if index not in enableColumns:
					continue
				name = sheet.row(4)[index].value
				comment = sheet.row(2)[index].value
				ctype = sheet.row(1)[index].value
				if sheet.row(1)[index].value == "int":
					ctype = "int64"
				f.write("\t%s\t%s\t`xml:%s,attr`\t// %s\n" % (name, ctype, name, comment))
			f.write("\n")

			# 关闭格式文件
			f.close()

	print("done.")


# 测试代码
# print(allExcels)
# print("______________")
#
# # 读取excel
# excel = xlrd.open_workbook(os.path.join(pathExcel, "test.xlsx"))
# sheet = excel.sheet_by_index(0)
# numRow = sheet.nrows
# for i in range(1,numRow):
# 	row = sheet.row(i)
# 	print(row[3])
# 	if row[3].ctype == xlrd.XL_CELL_NUMBER:
# 		print(111111111, row[3].value == int(row[3].value))
#
# # 写文件
# f = open(os.path.join(pathXML, "a.txt"), 'w')
# f.write("aaa")
# f.write("bbb")
# f.write("ccc")
# f.close()
