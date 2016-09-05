# -*- coding:utf-8 -*-

import pyExcelerator
import os, sys

style = pyExcelerator.XFStyle()
style.font.name = "Consolas"


# Multiple language Localizable.strings
class Document:

	def __init__(self, name, path, file):
		self.name = name
		self.path = path
		self.file = file
		# {language_name:{key:text, ...}, ...}
		self.__language_strings = {}
		self.language_names = []
		self.key_names = []

	def non_text_keys_for_language(self, language):
		if not self.__language_strings.has_key(language):
			return self.key_names
		strings = self.__language_strings[language]
		if len(strings) == 0:
			return self.key_names
		keys = []
		for key in self.key_names:
			if not strings.has_key(key):
				keys.append(key)
		return keys

	def strings_for_language(self, language):
		return self.__language_strings[language]

	def string_for_key_and_language(self, key, language):
		if not self.__language_strings.has_key(language) or not self.__language_strings[language].has_key(key):
			return None
		return self.__language_strings[language][key]

	def set_string_with_key_for_language(self, text, key, language):
		if key not in self.key_names:
			self.key_names.append(key)
		if not self.__language_strings.has_key(language):
			self.__language_strings[language] = {key:text}
		else:
			self.__language_strings[language][key] = text

	def save_strings_for_language_at_path(self, language, path):
		if not self.__language_strings.has_key(language):
			return False
		print "writing document {0} ({1} keys) to\n{2}".format(self.name, len(self.key_names), path)
		dir = os.path.dirname(path)
		if not os.path.isdir(dir):
			os.makedirs(dir)
		file = open(path, 'wb')
		strings = self.__language_strings[language]
		for key in self.key_names:
			text = strings[key]
			line = '"' + key + '" = "' + text.replace('"', '\\"') + '";\n'
			file.write(line.encode("utf-8"))
		file.close()
		return True

	def save_on_sheet(self, sheet):
		print "saving sheet " + self.name
		sheet.write(0, 0, self.path + "\n" + self.file, style=style)
		for (keyIndex, key) in enumerate(self.key_names):
			# print "saving key {0} on row {1}".format(key, rowIndex + 1)
			sheet.write(keyIndex + 1, 0, key, style=style)
		for (langIndex, lang) in enumerate(self.language_names):
			colIndex = langIndex + 1
			sheet.write(0, langIndex + 1, lang, style=style)
			for (keyIndex, key) in enumerate(self.key_names):
				rowIndex = keyIndex + 1
				value = self.string_for_key_and_language(key, lang)
				if value is not None:
					sheet.write(rowIndex, colIndex, value, style=style)


class Sheets:
	def __init__(self, excel_path):
		# [LocalizableDocument, ...]
		self.documents = []
		self.excel_path = excel_path
		if len(excel_path) > 0:
			self.load_excel_with_path(excel_path)

	def load_excel_with_path(self, path):
		sheets = pyExcelerator.parse_xls(path)
		for (sheet_name, values) in sheets:
			names = values[(0,0)].split("\n")
			doc = Document(sheet_name, names[0], names[1])

			cell_keys = values.keys()
			language_names = []
			col_index = 1
			while (0, col_index) in cell_keys:
				name = values[(0, col_index)]
				language_names.append(name)
				col_index += 1
			doc.language_names = language_names

			language_keys = []
			row_index = 1
			while (row_index, 0) in cell_keys:
				language_keys.append(values[(row_index, 0)])
				row_index += 1
			doc.key_names = language_keys

			for row_index, col_index in cell_keys:
				if row_index == 0 or col_index == 0:
					continue
				index = (row_index, col_index)
				v = values[index]
				lang = language_names[col_index - 1]
				key = language_keys[row_index - 1]
				doc.set_string_with_key_for_language(v, key, lang)
			print "loaded document '{0}' with {1} keys.".format(doc.name, len(doc.key_names))
			self.documents.append(doc)

	def save_excel(self, path = None):
		if path is None or len(path) == 0:
			path = self.excel_path
		workbook = pyExcelerator.Workbook()
		for doc in self.documents:
			worksheet = workbook.add_sheet(doc.name)
			doc.save_on_sheet(worksheet)
		tmp_path = os.path.dirname(path) + "/_" + os.path.basename(path)
		workbook.save(tmp_path)
		os.unlink(path)
		os.rename(tmp_path, path)

	def save_strings(self, basepath):
		for doc in self.documents:
			for lang in doc.language_names:
				path = basepath + "/" + doc.path + "/" + lang + ".lproj/" + doc.file + ".strings"
				doc.save_strings_for_language_at_path(lang, path)
