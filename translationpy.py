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

keys = []
translations = []
directory = 'export/'
filename = ''
#encoding = codecs.BOM_UTF16_BE

def createfile(column):
	langCode = column[0].value
	langName = column[1].value
	if langName==None or langCode==None:
		return

	langCode = langCode.strip()
	langName = langName.strip().title()

	translations.append((langName,langCode))

	print('Creating translation file for ' + langName + '\t(' + langCode + '.json)..')
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

if __name__ == '__main__':
	if len(sys.argv) > 1:
		filename = sys.argv[1]
	else:
		print ("Usage: python excel_to_json.py filename.xlsx [export_directory]\n")
		exit(-1)
	if len(sys.argv) > 2:
		directory = sys.argv[2]

	if not path.exists(directory):
		makedirs(directory)

	wb = load_workbook( filename )
	ws = wb.active #get active worksheet

	for column in ws.columns:
		# generate the keys list
		if keys.__len__() == 0:
			for cell in column:
				if cell.value is not None:
					keys.append(cell.value)
		else:
			createfile(column)

	print( 'Writing languages.json.. ' )
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

	print('Everything done.\n')
