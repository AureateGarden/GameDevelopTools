#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os, sys, getopt, xlrd, json, re

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
LongOpts = ["help", "version", "input=", "output=", "sheet="]

DefualtFileName = "Output.json"
LocalPath = os.path.dirname(os.path.abspath(__file__))
InputPath = ""
OutputPath = ""
StartFlag = "[start]"
JumpFlag = "[ignore]"
TransformSheet = -1;

Eldata = None;

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
		self.m_DicData = {}
		if self.__CurSheet < 0:
			self.__CurSheet = 0
			print("* set sheet to %d" % self.__CurSheet)
		if self.__ElData.nsheets < int(self.__CurSheet) + 1:
			self.__CurSheet = self.__ElData.nsheets - 1
			print("* Input sheet is too big, changed it to biggest sheet: %d" % self.__CurSheet)
		self.m_OutputPath = outputPath
		self.__SheetData = self.__ElData.sheet_by_index(self.__CurSheet)
		findFlag = False
		for i in range(self.__SheetData.nrows - 1):
			for j in range(len(self.__SheetData.row(i)) - 1):
				if self.__SheetData.row(i)[j].value.find(StartFlag) != -1:
					self.__StartPosi = Vector(j, i)
					findFlag = True
					break
				if findFlag:
					break
		del findFlag
		self.__DataRowLen = self.__SheetData.nrows - (self.__StartPosi.x + 1)
		self.__DataColLen = self.__SheetData.ncols - (self.__StartPosi.y + 1)

	def Transform(self):
		for i in range(self.__StartPosi.x + 1, self.__StartPosi.x + self.__DataRowLen + 1):
			itemName = self.__SheetData.row(i)[self.__StartPosi.x].value;
			if itemName == "" or itemName.find(JumpFlag) != -1:
				continue
			Tempdic = {}
			for j in range(self.__StartPosi.y + 1, self.__StartPosi.y + self.__DataColLen +	 1):
				if j != self.__StartPosi.y + 1:
					value = self.__SheetData.row(i)[j].value
					if type(value) == float:
						value = int(value)
					Tempdic[self.__SheetData.row(self.__StartPosi.x)[j].value] = value
			self.m_DicData[self.__SheetData.row(i)[self.__StartPosi.x + 1].value] = Tempdic

def GetHelp():
	print(HelpInfo)
	return

def GetVersion():
	print(Version)
	return

def OpenExcel(Path):
	data = xlrd.open_workbook(Path)
	return data

def GetInput(inputPath):
	if os.path.exists(inputPath) and os.path.isfile(inputPath):
		Edata = OpenExcel(inputPath);
	else:
		path = os.path.join(LocalPath, inputPath)
		if os.path.exists(path) and	os.path.isfile(path):
			Edata = OpenExcel(path)
	return Edata

def GetOutput(outputPath):
	if os.path.exists(outputPath):
		if os.path.isfile(outputPath):
			print("*Output path found:\n*-- " + outputPath)
			return outputPath
		else:
			path = os.path.join(outputPath, DefualtFileName)
			print("*Create file with default name in path:\n*-- " + path)
			return path
	else:
		tempPath = os.path.join(LocalPath, outputPath)
		if os.path.exists(tempPath):
			if os.path.isfile(tempPath):
				print("* Find file in default path:\n*-- " + tempPath)
				return tempPath
		else:
			if re.match("\w+.json", outputPath) != None:		#here!!!!!!!!!!!!!!!!!!!!!!!!!
				return tempPath
			else:
				print("file should named with \".json\"")
				path = os.path.join(LocalPath, DefualtFileName)
				print("* Not defined output path, set it with default path:\n*-- " + path)
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
				TransformSheet = int(opt_value)

	OutputPath = GetOutput(OutputPath)
	Eldata = GetInput(InputPath)
	if Eldata != None:
		trans = ExcelData(Eldata, TransformSheet, OutputPath)
		trans.Transform();
		try:
			jsonstr = json.dumps(trans.m_DicData)
			writer = open(trans.m_OutputPath, "w")
			writer.write(jsonstr)
		except Exception as e:
			print(e)
		else:
			print("Transform successful!")
			print("File is saved in path: " + trans.m_OutputPath)
		finally:
			writer.close()