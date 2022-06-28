import json
import re
import pandas as pd



filePath = str(input("Enter the path of the file: "))
excelFile = pd.ExcelFile(filePath)
spreadsheets = excelFile.sheet_names
print('spreadsheet loaded successfully')


def toCamelCase(spreadsheet):
    if "-" in spreadsheet:
        return ''.join(word.title() for word in spreadsheet.split('-'))
    return spreadsheet

def stringToJson(dictionary, keys, value):
    if "." in keys:
        key, rest = keys.split(".",1);
        if key not in dictionary:
            dictionary[key] = {}
        stringToJson(dictionary[key], rest, value)
    else:
        dictionary[keys] = value if isinstance(value, str) else "";
    return dictionary
def saveFile(dictionary, fileName):
    jsonStr = json.dumps(dictionary, indent=4, ensure_ascii=False);
    removeQuotes = re.sub("\"([^\"]+)\":", r"\1:", jsonStr);
    fileNameCleaned = fileName.split(" ")[0]
        
    with open(fileNameCleaned + ".ts", "w",encoding='utf_8') as outfile:
        print("saving spreadsheet to file: " + fileNameCleaned + ".ts")
        outfile.write("export const " + toCamelCase(fileNameCleaned) + " = " + removeQuotes + ";")
def parseTranslations(excelFile, spreadsheet, keyName, valueName):
    parseXlsx = pd.read_excel(excelFile, spreadsheet);
    keysDict = {}
    for key, value in zip(parseXlsx[keyName], parseXlsx[valueName]):
        stringToJson(keysDict, key, value)
    saveFile(keysDict,spreadsheet)
  


keyName = str(input("Enter the name of the column with the keys (e.g. 'Key'): "))
valueName = str(input("Enter the name of the column with the values (e.g. 'en'): "))

for spreadsheet in spreadsheets:
    print("parsing spreadsheet: " + spreadsheet)
    parseTranslations(excelFile, spreadsheet, keyName, valueName)
 
      
