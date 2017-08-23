# -*- coding:utf-8 -*-

import Excel
import AzureTranslate
import sys, os

if __name__ == '__main__':
	reload(sys)
	sys.setdefaultencoding('utf-8')

args = sys.argv[1:]
if len(args) < 2:
	print "usage: " + sys.argv[0] + " project_directory <to_excel|export|translate> [strings_file_name]"
	exit()

projDir = args[0]
action = args[1]

filename = len(args) < 3 and "Localizable" or args[2]
if not os.path.isdir(projDir):
	print projDir + " is not a directory."
	exit()

excel_path = projDir + "/" + filename + ".xls"
if not os.path.exists(excel_path):
	excel_path += "x"
sheets = Excel.Sheets(excel_path)
if action == "to_excel":
	sheets.save_excel()
elif action == "export":
	sheets.save_strings(projDir)
elif action == "translate":
	authorizer = AzureTranslate.Authorizer()
	for doc in sheets.documents:
		if len(doc.language_names) <= 1:
			print doc.name + " should have at least 2 languages to translate."
			continue
		print "Translating " + doc.name + "..."
		primary_lang = doc.language_names[0]
		for lang in doc.language_names:
			if lang == primary_lang:
				continue
			keys = doc.non_text_keys_for_language(lang)
			if len(keys) == 0:
				continue
			strings_of_primary = []
			for key in keys:
				text = doc.string_for_key_and_language(key, primary_lang)
				if isinstance(text, basestring):
					strings_of_primary.append(text)
			if not strings_of_primary:
				continue
			tr = AzureTranslate.Translator(authorizer)
			strings_transed = tr.translate(strings_of_primary, primary_lang, lang)
			# TODO: if any primary text lost, the index would be wrong after the lost key.
			for index, str in enumerate(strings_transed):
				doc.set_string_with_key_for_language(str, keys[index], lang)
			sheets.save_excel()
