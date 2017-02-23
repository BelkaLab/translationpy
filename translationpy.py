try:
	from openpyxl import load_workbook
except ImportError:
	print('openpyxl must be installed to run the script. [ pip install openpyxl ]')
	exit(-1)

from os import path, makedirs
import argparse
import json
import io

# nested dict allow us to create keys more easily, using nesting syntax
# see http://stackoverflow.com/a/16724937/2304450
import collections
nested_dict = lambda: collections.defaultdict(nested_dict)

directory = 'export/'
prologue = None
epilogue = None
file_format = 'json'
force_insertion = False
flatten_keys = False

# two super useful utils to get/set a key recursively
# http://stackoverflow.com/a/14692747/2304450
from functools import reduce
import operator

def getFromDict(dataDict, mapList):
	try:
		return reduce(operator.getitem, mapList, dataDict)
	except KeyError:
		return None
	except TypeError:
		return None

def setInDict(dataDict, mapList, value):
	getFromDict(dataDict, mapList[:-1])[mapList[-1]] = value

def createNestedKeysFile(column, keys, jsonSchema):
	langCode = column[0].value
	langName = column[1].value
	if langName is None or langCode is None:
		return

	langCode = langCode.strip()
	langName = langName.strip().title()

	print('  Processing ' + langName)

	translatedDict = nested_dict()

	for index, cell in enumerate(column):
		if index > 1 and index < keys.__len__():
			if cell.value is None:
				cell.value = ''

			key = keys[index].strip()
			value = cell.value.replace('"', '\\"')

			if jsonSchema:
				if getFromDict(jsonSchema, key.split('.')) is not None:
					setInDict(translatedDict, key.split('.'), value)
				else:
					if force_insertion:
						setInDict(translatedDict, key.split('.'), value)
						print('    WARN: key not found in given schema: ' + key + ' [forcing insert]')
					else:
						print('    WARN: key not found in given schema: ' + key + ' [skipping]')
			elif flatten_keys:
				translatedDict[key] = value
			else:
				setInDict(translatedDict, key.split('.'), value)

	filename = directory + '/' + langCode + '.' + file_format
	print('    Saving file ' + filename)

	with io.open(filename, 'w', encoding='utf-8') as outfile:
		if prologue:
			outfile.write(unicode(prologue))
		outfile.write(json.dumps(translatedDict, ensure_ascii=False, indent=2, sort_keys=flatten_keys))
		if epilogue:
			outfile.write(unicode(epilogue))

def process(sheet, schema):
	keys = []

	jsonSchema = None
	if schema is not None:
		with open(schema, 'r') as schema_file:
			jsonSchema = json.load(schema_file)

	for column in sheet.columns:
		# generate the keys list
		if keys.__len__() == 0:
			for cell in column:
				if cell.value is not None:
					keys.append(cell.value)
		else:
			createNestedKeysFile(column, keys, jsonSchema)

if __name__ == '__main__':
	parser = argparse.ArgumentParser(
		description='Transform an excel file containing translations in the separate JSON files required by polyglot.js',
		epilog='For more info, see https://github.com/BelkaLab/translationpy'
	)
	parser.add_argument('input', help='the input excel file')
	parser.add_argument('--out-dir', help='specify the output directory (default: export/)', default='export')
	parser.add_argument('--sheets', help='process only the specified sheets (default: active sheet, * for all)')
	parser.add_argument('--prologue', help='a string to prepend to the dumped JSON variable')
	parser.add_argument('--epilogue', help='a string to append to the dumped JSON variable')
	parser.add_argument('--schema', help='honour the given schema in output files')
	parser.add_argument('--format', help='output files format', choices={'js', 'json', 'jsonp'}, default='json')
	parser.add_argument('--force', help='force insertion of keys not found in given schema', action='store_true')
	parser.add_argument('--flatten-keys', help='make the output a flat dictionary', action='store_true')

	args = parser.parse_args()

	prologue = args.prologue
	epilogue = args.epilogue
	file_format = args.format
	force_insertion = args.force
	flatten_keys = args.flatten_keys

	workbook = load_workbook(args.input)

	if args.sheets:
		# process all sheets matching given list, or all sheets if '*' was given
		for sheetName in workbook.get_sheet_names():
			if args.sheets == '*' or sheetName in args.sheets:
				print('\nProcessing worksheet "%s"' % sheetName)

				directory = args.out_dir + '/' + sheetName
				if not path.exists(directory):
					makedirs(directory)

				process(workbook.get_sheet_by_name(sheetName), args.schema)
	else:
		 # process only active worksheet
		directory = args.out_dir
		if not path.exists(directory):
			makedirs(directory)

		print('\nProcessing active worksheet "%s"' % workbook.active.title)
		process(workbook.active, args.schema)

	print('\nEverything done.\n')
