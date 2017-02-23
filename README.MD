# translationpy
**Translationpy** is a Python script to automatically extract tranlsations from an excel file and create the needed JSON files to be used by [polyglot.js](http://airbnb.io/polyglot.js/)

## Install
The script usage is really simple, take a cup of coffee and install the following Python library:

```bash 
$ pip install openpyxl
```  
then, execute the script

```bash 
$ python translationpy.py translations.xlsx
```  
Done :)

## How to use it 
### Input File
The script input file must be an Excel file (.xlsx) formatted like the screenshot below.

- Define the first column "A" as the label column.
- Define the second column "B" and following columns as the translation columns.

![Alt Text](https://s29.postimg.org/wdq8t426f/Schermata+2016-03-30+alle+18.05.32.png)

Translationpy will extract one JSON file per column, using the first row as file names. Each file will be in the `key: value` form, where key is always the value coming from the first column.

### Options

|Flag|Option|Example|Default|
|----|------|-------|-------|
|`--out-dir`|specify the output directory|`--out-dir=translations`|`export/`|
|`--sheets`|process only the specified sheets|`--sheets=Login,Container`, `--sheets=*`|only the active sheet|
|`--prologue`|a string to prepend to the dumped JSON variable|`--prologue="const translations = "`|-|
|`--epilogue`|a string to append to the dumped JSON variable|`--epilogue=";\n\nexport default translations;\n"`|-|
|`--format`|output files format {js, json, jsonp}|`--format=js`|`json`|
|`--schema`|honour the given schema in output files|`--schema=schema.json`|-|
|`--force`|force insertion of keys not found in given schema|`--schema=schema.json --force`|`false`|
|`--flatten-keys`|make the output a flat dictionary|`--flatten-keys`|`false`|

### Using a pre-existing schema
If an existing JSON file is given through the `--schema` flag, the script will load it and use it to verify and keep the outputted keys ordered.

Otherwise, the dictionary structure will be automatically generated; in this case, key ordering will be given by the first occourrence of any new dictionary sub-key.

Additionally, you can print a flattened, alphabetically-sorted dictionary (ie: having depth 1) using the `--flatten-keys` option.

## Caveats and pitfalls
 - Rows #1 and #2 **must** contain the language tag and the language name.
 - Always ensure that the first column has no empty rows between keys, otherwise the keys order will be broken.
 - Columns missing a value in the first or second row (ie: language tag or language name) will be skipped.

## License
translationpy is Copyright (c) 2017 Belka, srl. It is free software, and may be redistributed under the terms specified in the LICENSE file.  

## About Belka
![Alt text](http://s2.postimg.org/rcjk3hf5x/logo_rosso.jpg)

[Belka](http://belka.us/en) is a Digital Agency specialized in design, mobile applications development and custom solutions.
We love open source software! You can [see our projects](http://belka.us/en/portfolio/) or look at our case studies.

Interested? [Hire us](http://belka.us/en/contacts/) to help build your next amazing project.

[www.belka.us](http://belka.us/en)
