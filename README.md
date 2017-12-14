# Localizable.excel
Easy translate and export Localizable strings from Excel.

The command line tool can work based on an excel file.

It may easy export translated strings as Localizable.strings with difference file name and location.

And it may translate all empty cell with Bing Translate service. You may register free, and translate 200000 characters per month as free, it usually enough for simple app.

## Structure of Excel
Each work sheet is a Localizabe.strings group:

* Cell[0,0] should be "PathToSaving\nLocalizableFileName".
* Other cells on row[0] should be language code, e.g. en, zh-Hans, it.
* Other cells on column[0] should be strings key, what you typed in code - NSLocalizableString("StringKey", comment: nil).
* Cells on column[1] is primary language, it should be fully filled.

		YourApp
		Localizable   en           ja               de
		home.title    Text Editor  テキストエディタ
		home.footer   Copyright

## Install python library
The tool would use pyExcelerator and pycurl, download and install them please.

	https://sourceforge.net/projects/pyexcelerator/
	http://pycurl.io/

## Translate
You should register on Bing Translation, save client ID and client secret into azure_client.json:

	{"client_id":"YourClientID","client_secret":"YourClientSecret"}
Call the command on Command Line:

	python main.py translate path/to/project/contains/excel
The tool would translate all empty cell with primary language, and then save back.

## Export strings
Call the command on Command Line:

	python main.py export path/to/project/contains/excel
All strings defined work sheet would be export, with difference language sub directory. e.g.

	path/to/project/contains/excel/YourApp/en.lproj/Localizable.string
	path/to/project/contains/excel/YourApp/ja.lproj/Localizable.string
	...

## Tips
* Currently the tool cannot save style on excel but font name, you can change it on head of Excel.py.
* After translate and export, the excel and strings would be overwrite, any changes in strings and style in excel would lost.

## License
MIT
