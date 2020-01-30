from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located
from selenium.webdriver.common.by import By



import win32com.client as word
import pandas as pd
import re
import numpy
import os
import sys
import time
import requests
import progressbar

global intag
global intag_withp
global nonclosed
global closed
global findrows
global findcell
global data_in_brackets



intag = re.compile(r'(.*?(<(?P<tag>[^p].*[ ]?).*?>))(?P<data>.*?)(</(?P=tag)>)')
intag_withp = re.compile(r'(.*?(<(?P<tag>.*[ ]?).*?>))(?P<data>.*?)(</(?P=tag)>)')
nonclosed = re.compile(r'(<p.*?>)(.*)')
closed = re.compile(r'<p.*?>(.*?)</p>')
findrows = re.compile(r'<tr.*?</tr>')
findcell = re.compile(r'<td.*?</td>')
data_in_brackets = re.compile(r'>.*<')

libDir = 'C:/SpecificationsDump/library/'
dataDir = r'C:/SpecificationsDump/data/'
dataType = r'PK'
libType = r'PK'
workingDir = 'C:\SpecificationsDump'
slash = '\ '
workingDir = 'C:\SpecificationsDump'+slash[0]+dataType+'Specs'+slash[0]

sys.stdout = open(dataDir + 'log' + dataType+ '.txt', 'w') #LOGGER

def countQSweep(prices):
    listLen = prices['Кол-во'].sum()
    midValue = (prices['Цена за ед.']*prices['Кол-во']).sum() / (listLen)
    result = (((prices['Цена за ед.']-midValue) ** 2)*prices['Кол-во']).sum()
    result = 3*numpy.sqrt(result / (listLen))
    return result+2000

def onlyNums(nList):
    temp = []
    toDot = lambda x: re.sub(',', '.', x)
    isNumber = re.compile(r'\A[1-9].*')
    for num in nList:
        num = toDot(num)
        if isNumber.match(num) is not None:
            temp.append(isNumber.match(num).group())
    return temp

def parseFile(data):

    data = re.sub(r'\n', '', data)
    data = re.sub(r'&nbsp;', ' ', data)
    data = re.sub(r'&quot;', '\"', data)

    rows = []
    for i in findrows.finditer(data):
        rows.append(i[0])

    for row in rows:
        rows[rows.index(row)] = re.sub(r'<.*?>', '', row)
    for row in rows:
        rows[rows.index(row)] = re.sub(r'(\d+)? (\d+)? (\d+)? (\d+[,.]\d+)', r'\1\2\3\4', row)
    for row in rows:
        rows[rows.index(row)] = re.sub(r'(\d+),(\d+)', r'\1.\2', row)
    for row in rows:
        rows[rows.index(row)] = row.lower()


    return rows

def toCountable(list, type='float'):
    result = []
    for num in range(len(list)):
        try:
            if type == 'int':
                result.append(numpy.int(list[num]))
            else:
                result.append(numpy.float(list[num]))
        except OverflowError:
            num.pop()
    return result


def convertdTypes(dFrame):
    chFormatColumns = ['Номер записи', 'Цена за ед.', 'Кол-во', 'Сумма']
    for col in chFormatColumns:  # ' 123 123,00 '
        dFrame[col] = dFrame[col].astype(str).apply(lambda x: re.sub(r'\s(.*)\s', r'\1', x))  # '123 123,00'
        dFrame[col] = dFrame[col].astype(str).apply(lambda x: re.sub(',', '.', x))  # '123 123.00'
        dFrame[col] = dFrame[col].astype(str).apply(lambda x: re.sub(r'(\d)\s(\d)', r'\1\2', x))  # '123123.00'
    dFrame['Цена за ед.'] = pd.to_numeric(dFrame['Цена за ед.'], errors='coerce')
    dFrame['Кол-во'] = pd.to_numeric(dFrame['Кол-во'], errors='coerce')
    dFrame['Сумма'] = pd.to_numeric(dFrame['Сумма'], errors='coerce')
    dFrame = dFrame.replace(numpy.nan, 0, regex=True)
    return dFrame

def changeFormat(badFile):
    msWord = word.Dispatch("Word.Application")
    doc = msWord.Documents.Open(badFile)
    docxFormatNumber = 10#10-html#7=text#6-RTF #5 - DOStxt\w\line\breaks #16 - код docx по WdSaveFormat (Office VBA Reference)
    if re.match(r'\..{3}', badFile) is not None:
        doc.SaveAs2(FileName=workingDir+libType+'txt\\'+re.findall(r'\d.*', badFile)[0][:-4] + ".txt", FileFormat=docxFormatNumber) #обрезаем старый формат и ставим новый
        msWord.Quit(False)
        return workingDir+libType+'txt\\'+re.findall(r'\d.*', badFile)[0][:-4] + ".txt"
    else:
        doc.SaveAs2(FileName=workingDir+libType+'txt\\'+re.findall(r'\d.*', badFile)[0][:-5] + ".txt", FileFormat=docxFormatNumber)  # обрезаем старый формат и ставим новый
        msWord.Quit(False)
        return  (workingDir+libType+'txt\\'+re.findall(r'\d.*', badFile)[0][:-5] + ".txt")

def findLinkTo(source):
    try:
        bad = False
        fileLink = re.compile(
            r'http://zakupki.gov.ru/44fz/filestore/public/1.0/download/.{1,100}"\s*title=".{1,100}".{1,640}онтракт.{1,640}№.{1,640}</a>', re.DOTALL | re.IGNORECASE) #{1}\s*?</td>{1}(?!.*<)
        fileLink = fileLink.findall(source)
        fileLink = fileLink[0]
        flink = re.findall(r'http://zakupki.gov.ru/44fz/filestore/public/1.0/download/.*?"', fileLink)
        flink = flink[0]
        flink = flink[:-1]
        fFormat = re.findall(r'(?<=title=").*?"', fileLink)
        fFormat = fFormat[0]
        fFormat = re.findall(r'\.[A-Za-z]*?\s', fFormat)
        fFormat = fFormat[0]
        fFormat = fFormat[:-1]
        return flink, fFormat, bad
    except IndexError:
        bad = True
        return '', '', bad

def isNAN(x):
    if [x] == [numpy.nan]:
        return '0'
    else:
        return x

# def makeItGood(txtFile):
#     ints, floats, txtFile = parseFile(txtFile)
#     ints, floats = onlyNums(ints), onlyNums(floats)
#     ints, floats = toCountable(ints, 'int'), toCountable(floats, 'float')
#     ints, floats = list(set(ints)), list(set(floats))
#     return ints, floats, txtFile


#настройка chrome
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {"download.default_directory": workingDir, "download_restrictions": 0}) #Место загрузки по умолчанию
#driver = webdriver.Chrome('C:/chromedriver.exe', chrome_options=options)

fileEVH = pd.read_csv(dataDir+dataType+r'.csv', ";", encoding="ANSI")
fileEVH = convertdTypes(fileEVH)
fileEVH[['Модель','Пр-ль']] = fileEVH[['Модель','Пр-ль']].apply(lambda x: x.str.lower())
modelPrices = pd.read_csv(libDir+libType+r'_modelPrices.csv', ";", encoding="ANSI") #columns={'model': str, 'midprice': float, 'entrys': int}
lib = pd.read_csv(libDir+libType+r'LIB.csv', ";", encoding="ANSI")
nan = numpy.nan
fileEVH['Модель'] = fileEVH['Модель'].apply(lambda x: isNAN(x))
fileEVH['Модель'] = fileEVH['Модель'].apply(lambda x: re.sub('не определен', '0', x))
isModel = pd.Series((fileEVH['Модель'].str.contains('0', na=False)).values, name='Модель')

for model in modelPrices['model']:
    i = 0
    for filledM in fileEVH['Модель']:
        if [filledM] != [nan]:
            try:
                if re.search(model, filledM) is not None:
                    isModel.loc[i]=False
            except TypeError:
                print('NAN in models prices list. Check ', libDir + libType + '_modelPrices.csv')
        i = i + 1

searchColumns = ['Модель', 'Пр-ль', 'Номер записи', 'Цена за ед.', 'Кол-во', 'Сумма']
emptyEVH = fileEVH[searchColumns][isModel].copy()
for f in os.listdir(workingDir+libType+'txt'):
    os.remove(workingDir+libType+'txt\\'+f)
# for f in os.listdir(workingDir):
#     os.rename(f, re.sub(r'used', '', f))
for contract in pd.unique(emptyEVH['Номер записи']):
    lilEmptyFilter = emptyEVH['Номер записи'].isin([contract])
    lilEmpty = emptyEVH[lilEmptyFilter]
    pricesList = lilEmpty['Цена за ед.']
    contractFiles = os.listdir(workingDir)
    for fs in contractFiles:
        if len(fs)>21:
            filename = fs
            if re.match(contract, filename) is not None:
                canOpen = True
                if (re.match(r'(.*doc.*$)|(.*rtf.*$)', filename, re.IGNORECASE) is not None) and canOpen:
                    fileLink = changeFormat(workingDir+filename)
                    txtFileObj = open(fileLink)

                    try:
                        txtFile = txtFileObj.read()
                        rows = parseFile(txtFile)
                    except UnicodeDecodeError:
                        print(fileLink+' SHITTED')

                    for priceidx in pricesList.index:
                        modelFound = False
                        price = pricesList.loc[priceidx]
                        allowedModels = pd.DataFrame()
                        for idx in modelPrices.index:
                            max = modelPrices.loc[idx, 'midprice'] + (modelPrices.loc[idx, 'qsweep'])
                            min = modelPrices.loc[idx, 'midprice'] - (modelPrices.loc[idx, 'qsweep'])
                            if min < price and price < max:
                                allowedModels = pd.concat([allowedModels, pd.DataFrame(modelPrices.iloc[idx]).transpose()])
                        if allowedModels.size != 0:
                            for row in rows:
                                for model in allowedModels['model']:
                                    if (re.search(str(price), row) is not None) and (re.search(model.lower(), row)):
                                        thisisit = model
                                        modelFound = True
                            if modelFound == False:
                                rowsdump = ''
                                for row in rows:
                                    rowsdump = rowsdump + ' ' + row
                                if (re.search(str(price), rowsdump) is not None) and (re.search(model.lower(), rowsdump)):
                                    thisisit = model
                                    modelFound = True
                        if modelFound:
                            fileEVH.loc[priceidx, 'Модель'] = thisisit
                        else:
                            print('For' + fileEVH.loc[priceidx].to_string() + ' no models found')
                else:
                    canOpen = False
                    print('Can`t open file for', contract, ' ' ,workingDir+filename)

fileEVH.to_csv(dataDir+dataType+'_filledfromContracts.csv', mode='a', encoding='ANSI', sep=';', index=False)





