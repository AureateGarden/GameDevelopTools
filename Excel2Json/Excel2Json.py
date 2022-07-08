#!/usr/bin/python
# -*- coding: UTF-8 -*-

from ast import Expression
from audioop import reverse
import os, sys, getopt, openpyxl, json, re

Version = \
R'''
+---------------------------------------------------------------+
|   Excel2Json version: 0.0.1                                   | 
+---------------------------------------------------------------+
|   MadeBy: Zeo                                                 |
+---------------------------------------------------------------+
|   If you encounter any problems, please give me feedback.     |
|   github Issue page:                                          |
|   https://github.com/wlz330860661/UnityTransformToos/issues   |
+---------------------------------------------------------------+ 
'''

HelpInfo = \
'''
Use function:
	pls add "[start]" flag in excel file, where you want to transform.
Transformd structer like this below:
{"item_name":{"item_Propertiy":"value", "item_Propertiy2":"value2"}, ...}

item_name is the first col of the excel file at start position's right
item_name's value is a dictionary witch started transformd from the second
col of the excel file at start position's right.

sorry about that I can only use chinglish to write this help infomation.^_^

+-----------------------------------------------------------------------------+
|   Options:                                                                  |
|      -h, --help          Show help info.                                    |
|      -v, --version       Show version.                                      |
|      -i, --input         Set excel file path witch you want to transform.   |
|      -o, --output        Set json file path witch you want to save.         |
|      -s, --sheet         Set excel file's sheet witch to ransform.          |
+-----------------------------------------------------------------------------+
'''

ShortOpts = "-h-v-i:-o:-s:"
LongOpts = ("help", "version", "input=", "output=", "sheet=")

DefualtFileName = "Output.json"
LocalPath = os.path.dirname(os.path.abspath(__file__))
InputPath = ""
OutputPath = ""
JumpFlag = "[ignore]"
TransformSheet = ""
SkipStartRow = 3
SkipStartColumn = 1

Eldata = None

builtInType = ("int", "float", "string", "boolean", "enum")

class Vector:
	"""docstring for Vector"""
	def __init__(self, *args):
		lenth = len(args)
		if lenth == 0:
			self.x = 0
			self.y = 0
		elif lenth == 2:
			self.x = args[0]
			self.y = args[1]

	def __add__(self, rhs):
		tempX = self.x + rhs.x
		tempY = self.y + rhs.y
		return Vector(tempX, tempY)

	def __sub__(self, rhs):
		tempX = rhs.x - self.x
		tempY = rhs.y - self.y
		return Vector(tempX, tempY)

	def __str__(self):
		return "Vector(%d, %d)" % (self.x, self.y)

class ExcelData:
	"""docstring for ExcelData"""
	def __init__(self, inputElData, inputSheet, outputPath):
		if inputElData == None:
			del self
			return
		self.__ElData = inputElData
		self.__CurSheet = inputSheet
		self.m_DicData = []
		if self.__CurSheet == "":
			self.__CurSheet = self.__ElData.sheetnames[0]
			print("* set sheet to %s" % self.__CurSheet)
		elif self.__ElData.sheetnames.count(self.__CurSheet) <= 0:
			self.__ElData._add_sheet(self.__CurSheet)
		self.m_OutputPath = outputPath
		self.__SheetData = self.__ElData[self.__CurSheet]

	def Transform(self):
		for i in range(self.__SheetData.min_row + SkipStartRow, self.__SheetData.max_row + 1):
			itemName = self.__SheetData.cell(row=i, column=self.__SheetData.min_column).value
			if itemName == "" or str(itemName).find(JumpFlag) != -1:
				continue
			tmpData = {}
			for j in range(self.__SheetData.min_column + SkipStartColumn, self.__SheetData.max_column + 1):
				value = self.__SheetData.cell(row=i, column=j).value
				columnHead = self.__SheetData.cell(row=self.__SheetData.min_row, column=j).value
				columnType = self.__SheetData.cell(row=self.__SheetData.min_row + 1, column=j).value
				if builtInType.count(columnType) <= 0:
					try:
						obj = json.loads(value)
						tmpData[columnHead] = obj
					except Exception as e:
						print("[EORROR]: {0}".format(e))
						tmpData[columnHead] = value
				else:
					tmpData[columnHead] = value
			self.m_DicData.append(tmpData)

def GetHelp():
	print(HelpInfo)
	return

def GetVersion():
	print(Version)
	return

def OpenExcel(Path):
	data = openpyxl.load_workbook(Path)
	return data

def GetInput(inputPath):
	if os.path.exists(inputPath) and os.path.isfile(inputPath):
		Edata = OpenExcel(inputPath)
	else:
		path = os.path.join(LocalPath, inputPath)
		if os.path.exists(path) and	os.path.isfile(path):
			Edata = OpenExcel(path)
	return Edata

def GetOutput(outputPath):
	if os.path.isfile(outputPath):
		print("*Output path found:\n* " + outputPath)
		return outputPath
	elif os.path.exists(outputPath):
		path = os.path.join(outputPath, DefualtFileName)
		print("*Create file with default name in path:\n* " + path)
		return path
	else:
		path = os.path.join(LocalPath, DefualtFileName)
		print("* Not defined output path, set it with default path:\n* " + path)
		return path

		
		

if __name__ == '__main__':
	opts, args = getopt.getopt(sys.argv[1:], ShortOpts, LongOpts)
	if len(opts) == 0 and len(args) == 0:
		GetHelp()
		exit()
	for opt_name, opt_value in opts:
		if opt_name in ('-h','--help') or opt_value == "help":
			GetHelp()
			exit()
		elif opt_name in ('-v', '--version'):
			GetVersion()
			exit()
		elif opt_name in ('-i', '--input'):
			InputPath = str(opt_value)
		elif opt_name in ('-o', '--output'):
			OutputPath = str(opt_value)
		elif opt_name in ('-s', '--sheet'):
			TransformSheet = opt_value

	OutputPath = GetOutput(OutputPath)
	Eldata = GetInput(InputPath)
	if Eldata != None:
		trans = ExcelData(Eldata, TransformSheet, OutputPath)
		trans.Transform()

		try:
			writer = open(trans.m_OutputPath, "w")
			jsonstr = json.dumps(trans.m_DicData, indent=4, separators=(',', ' : '))
			writer.write(jsonstr)
		except Exception as e:
			print(e)
		else:
			print("Transform successful!")
			print("File is saved in path: " + trans.m_OutputPath)
		finally:
			writer.close()