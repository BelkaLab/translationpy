try:
	import openpyxl
except ImportError:
	print ("openpyxl must be installed to run the script. [ pip install openpyxl ]")
	exit(-1)

from openpyxl import load_workbook
from openpyxl import workbook
import sys
from os import path, makedirs
import shutil
import codecs
import argparse

directory = 'export/'
#encoding = codecs.BOM_UTF16_BE

def createfile(column, keys, translations):
	langCode = column[0].value
	langName = column[1].value
	if langName==None or langCode==None:
		return

	langCode = langCode.strip()
	langName = langName.strip().title()

	translations.append((langName,langCode))

	print('  Creating translation file for {:10s}\t({:s}.json)..'.format(langName, langCode) )
	f = open( directory + '/' + langCode + '.json', 'w')
	#f.write(encoding)

	f.write('{\n')

	for index, cell in enumerate(column):
		if index>1 and index<keys.__len__()-1:
			if cell.value is None:
				cell.value = ''
			f.write( ('\t"' + keys[index].strip() + '": "' + cell.value.replace( '"', '\\"' ) + '"').encode('utf-8') )
			if index<column.__len__()-1:
				f.write( ',\n' )
			else:
				f.write( '\n' )

	f.write( '}\n' )
	f.close()

	if langCode == 'en':
		shutil.copyfile( directory + '/en.json',  directory + '/lang.json' )

def processSheet(sheet):
	keys = []
	translations = []
	for column in sheet.columns:
		# generate the keys list
		if keys.__len__() == 0:
			for cell in column:
				if cell.value is not None:
					keys.append(cell.value)
		else:
			createfile(column, keys, translations)

	print( 'Writing ' + directory + '/languages.json.. ' )
	f = open( directory + '/languages.json', 'w' )
	#f.write(encoding)
	f.write( 'loadTranslationsList({\n' )
	for index, (langName,langCode) in enumerate(translations):
		f.write( '\t"' + langName + '": "' + langCode + '"' )
		if index<translations.__len__()-1:
			f.write( ',\n' )
		else:
			f.write( '\n' )
	f.write( '});' )

if __name__ == '__main__':
	parser = argparse.ArgumentParser(description='Transform an excel file containing translations in the separate JSON files required by polyglot.js')
	parser.add_argument('input', help='the input excel file')
	parser.add_argument('--out-dir', help='specify the output directory (default: export/)', default='export')
	parser.add_argument('--sheets', help='process only the specified sheets (default: active sheet, * for all)')

	args = parser.parse_args()

	wb = load_workbook( args.input )

	if args.sheets:
		# process all sheets matching given list, or all sheets if '*' was given
		for sheetName in wb.get_sheet_names():
			if args.sheets == '*' or sheetName in args.sheets:
				print( '\nProcessing worksheet "%s"' % sheetName )

				directory = args.out_dir + '/' + sheetName
				if not path.exists(directory):
					makedirs(directory)

				processSheet(wb.get_sheet_by_name(sheetName))
	else:
		 # process only active worksheet
		directory = args.out_dir
		if not path.exists(directory):
			makedirs(directory)

		print('\nProcessing active worksheet "%s"' % wb.active.title)
		processSheet(wb.active)

	print('\nEverything done.\n')
