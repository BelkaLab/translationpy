# translationpy
Translationpy is a simple Python script to automatically extract and create translation files.

## Install
The script usage is really simple, take a cup of coffee and install the following Python library:

```bash 
$ pip install openpyxl
```  
then, execute the script

```bash 
$ python translationpy.py input.excel output-folder
```  
Done :)

## How to use it 
### Input File
The script input file must be an Excel file (.xlsx) formatted like the screenshot below.
- Define the first column "A" as the label column.
- Define the second column "B" and following columns as the translation columns.

![Alt Text](https://s29.postimg.org/wdq8t426f/Schermata+2016-03-30+alle+18.05.32.png)

Translationpy will extract one JSON file per column, using the first row as file names. Each file will be in the key: value form, where key is always the value coming from the first column.

### Output folder
You should specify the output folder where the script will save the translation files after the elaboration process (defaults to output/). 

## License
translationpy is Copyright (c) 2016 Belka, srl. It is free software, and may be redistributed under the terms specified in the LICENSE file.  

## About Belka
![Alt text](http://s2.postimg.org/rcjk3hf5x/logo_rosso.jpg)

[Belka](http://belka.us/en) is a Digital Agency specialized in design, mobile applications development and custom solutions.
We love open source software! You can [see our projects](http://belka.us/en/portfolio/) or look at our case studies.

Interested? [Hire us](http://belka.us/en/contacts/) to help build your next amazing project.

[www.belka.us](http://belka.us/en)