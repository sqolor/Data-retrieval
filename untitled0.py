import zipfile
from bs4 import BeautifulSoup

paths = []
mySheet = 'sheet'
filename = 'data.xlsx'
file = zipfile.ZipFile(filename, "r")

for name in file.namelist():
    if name == 'xl/workbook.xml':
        data = BeautifulSoup(file.read(name), 'html.parser')
        sheets = data.find_all('sheet')
        for sheet in sheets:
            paths.append([sheet.get('name'), 'xl/worksheets/sheet' + str(sheet.get('sheetid')) + '.xml'])

for path in paths:
    if path[0] == mySheet:
        with file.open(path[1]) as reader:
            for row in reader:
                print(row)  ## do what ever you want with your data
        reader.close()