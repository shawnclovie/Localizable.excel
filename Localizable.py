# -*- coding:utf-8 -*-

import codecs
import pyExcelerator
import os, sys

import Excel


class Localizable:
	def __init__(self, projDir, filename, baseLanguage):
		self.projectDir = projDir
		self.filename = filename
		self.languageStrings = {}
		self.baseLanguage = baseLanguage

	def excelPath(self):
		return self.projectDir + '/' + self.filename + '.xls'

	def mkdirs(self, path, mode = 0777):
		if os.path.isdir(path):
			return
		dir = os.path.dirname(path)
		self.mkdirs(dir, mode)
		os.mkdir(path, mode)

	def invoke(self):
		print()


class LocalizableExcelBuilder(Localizable):
	def listFilesAtPath(self, path, result):
		for file in os.listdir(path):
			subpath = os.path.join(path, file)
			if os.path.isdir(subpath):
				self.listFilesAtPath(subpath, result)
			else:
				result.append(subpath)
		return result

	# Load Localizable.strings from path
	# @return string contents: [[key,text], ...other keys]
	# Or nil if file not found.
	def parseStrings(self, path):
		fullpath = path + "/" + self.filename + ".strings"
		if not os.path.exists(fullpath):
			return None
		file = codecs.open(fullpath, 'r', 'utf-8')
		string = file.read()
		file.close()

		string = string.replace('/* No comment provided by engineer. */', '').replace('\n', '')
		list = [x.split(' = ') for x in string.split(';')]
		for rowIndex in range(len(list)):
			for valueIndex in range(len(list[rowIndex])):
				list[rowIndex][valueIndex] = list[rowIndex][valueIndex].strip("\"")
		return list

	def keysForLanguage(self, lang):
		keys = []
		for pair in self.languageStrings[lang]:
			if len(pair[0]) == 0:
				continue
			keys.append(pair[0])
		return keys

	def invoke(self):
		files = os.listdir(self.projectDir)
		for file in files:
			if not file.endswith(".lproj"):
				continue
			language = os.path.splitext(file)[0]
			path = self.projectDir + "/" + file
			strings = self.parseStrings(path)
			if strings is not None:
				self.languageStrings[language] = strings
		if not self.baseLanguage in self.languageStrings:
			strings = self.parseStrings(self.projectDir)
			if strings is not None:
				self.languageStrings[self.baseLanguage] = strings
		# print self.languageStrings
		self.keys = self.keysForLanguage(self.baseLanguage)

		workbook = pyExcelerator.Workbook()
		self.worksheet = workbook.add_sheet(self.filename)
		column = 0
		self.writeKeys()
		for (lang, strings) in self.languageStrings.iteritems():
			column += 1
			stringsDict = {}
			for pair in strings:
				if len(pair) < 2:
					continue
				stringsDict[pair[0]] = pair[1]
			self.writeStringsForLanguageOnColumn(lang, stringsDict, column)
		workbook.save(self.excelPath())


class LocalizableStringsExporter(Localizable):
	def __init__(self):
		super(Localizable).__init__()

	def invoke(self):
		sheets = pyExcelerator.parse_xls(self.excelPath())
		for (sheet_name, values) in sheets:
			names = sheet_name.split("@")
			filename = names[0] + ".strings"
			subdir = len(names) > 1 and names[1] or values[(0,0)]

			cellKeys = values.keys()

			languageNames = []
			colIndex = 1
			while (0, colIndex) in cellKeys:
				name = values[(0, colIndex)]
				languageNames.append(name)
				colIndex += 1

			languageKeys = []
			rowIndex = 0
			while (rowIndex, 0) in cellKeys:
				languageKeys.append(values[(rowIndex, 0)])
				rowIndex += 1

			languageStrings = {}
			for lang in languageNames:
				languageStrings[lang] = {}

			for rowIndex, colIndex in cellKeys:
				if rowIndex == 0 or colIndex == 0:
					continue
				index = (rowIndex, colIndex)
				v = values[index]
				lang = languageNames[colIndex - 1]
				key = languageKeys[rowIndex]
				languageStrings[lang][key] = v

			for lang in languageNames:
				path = self.projectDir + "/" + subdir + "/" + lang + ".lproj/" + filename
				self.saveStringsForPath(languageStrings[lang], path)

	def saveStringsForPath(self, strings, path):
		# print len(strings.keys()) + " keys writing to " + path
		dir = os.path.dirname(path)
		if not os.path.isdir(dir):
			os.makedirs(dir)
		file = open(path, 'wb')
		for (key, text) in strings.items():
			line = '"' + key + '" = "' + text + '";\n'
			file.write(line.encode("utf-8"))
		file.close()


args = sys.argv[1:]
if len(args) < 2:
	print "usage: " + sys.argv[0] + " <to_excel|export> project_directory [strings_file_name] [base_language_name]"
	exit()

action = args[0]
projDir = args[1]
filename = len(args) < 3 and "Localizable" or args[2]
baseLanguage = len(args) < 4 and "Base" or args[3]
if not os.path.isdir(projDir):
	print projDir + " is not a directory."
	exit()

if action == "to_excel":
	inst = LocalizableExcelBuilder(projDir, filename, baseLanguage)
	inst.invoke()
elif action == "export":
	LocalizableSheets(projDir + "/" + filename + ".xlsx")
	inst = LocalizableStringsExporter(projDir, filename, baseLanguage)
	inst.invoke()
