# -*- coding:utf-8 -*-

import Excel
import AzureTranslate
import sys, os

if __name__ == '__main__':
	reload(sys)
	sys.setdefaultencoding('utf-8')

args = sys.argv[1:]
if len(args) < 2:
	print "usage: " + sys.argv[0] + " <to_excel|export|translate> project_directory [strings_file_name] [base_language_name]"
	exit()

action = args[0]
projDir = args[1]
filename = len(args) < 3 and "Localizable" or args[2]
baseLanguage = len(args) < 4 and "Base" or args[3]
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
		primary_lang = doc.language_names[0]
		for lang in doc.language_names:
			if lang == primary_lang:
				continue
			keys = doc.non_text_keys_for_language(lang)
			if len(keys) == 0:
				continue
			strings_of_primary = []
			for key in keys:
				strings_of_primary.append(doc.string_for_key_and_language(key, primary_lang))
			tr = AzureTranslate.Translator(authorizer)
			strings_transed = tr.translate(strings_of_primary, primary_lang, lang)
			for index, str in enumerate(strings_transed):
				doc.set_string_with_key_for_language(str, keys[index], lang)
			sheets.save_excel()
