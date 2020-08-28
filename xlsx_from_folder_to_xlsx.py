import os
import pandas as pd
import numpy as np

from pprint import pprint
from tkinter import filedialog
from tkinter import *

def walklevel(some_dir, level=1):
    some_dir = some_dir.rstrip(os.path.sep)
    assert os.path.isdir(some_dir)
    num_sep = some_dir.count(os.path.sep)
    for root, dirs, files in os.walk(some_dir):
        yield root, dirs, files
        num_sep_this = root.count(os.path.sep)
        if num_sep + level <= num_sep_this:
            del dirs[:]
resultfilename = '_result.xlsx'
root = Tk()
root.withdraw()
folder_selected = filedialog.askdirectory()

filelist = []
for _, _, files in walklevel(folder_selected, 0):  
    for filename in files:
        if 'сборка' not in filename and 'result' not in filename:
            filelist.append(filename)

filnumdict = {0: 'Белгородский',
 1: 'БО',
 2: 'ВИ',
 3: 'Воронежский',
 4: 'Калужский',
 5: 'Курский',
 6: 'Липецкий',
 7: 'Смоленский',
 8: 'Тамбовский',
 9: 'Тверской',
 10: 'TР',
 11: 'ЯК'}

filnumb = {0: ['елг'],
 1: ['БО', 'рянск', 'рел', 'рл'],
 2: ['ВИ', 'ладим', 'ван'],
 3: ['орон'],
 4: ['алу'],
 5: ['урск'],
 6: ['ипец'],
 7: ['ТР', 'язан', 'ул'],
 8: ['мол'],
 9: ['амбов'],
 10: ['вер'],
 11: ['ЯК', 'росл', 'остр']}

numfil = {'елг': 0,
 'БО': 1,
 'рянск': 1,
 'рел': 1,
 'рло': 1,
 'ВИ': 2,
 'ладим': 2,
 'ван': 2,
 'орон': 3,
 'алу': 4,
 'урск': 5,
 'ипец': 6,
 'ТР': 7,
 'язан': 7,
 'ул': 7,
 'мол': 8,
 'амбов': 9,
 'вер': 10,
 'ЯК': 11,
 'росл': 11,
 'остр': 11}



def getsheetname(_fullfilename):
    with pd.ExcelFile(_fullfilename) as xl:
        for _sheet_name in xl.sheet_names:
            if 'нараст' in _sheet_name:
                return _sheet_name


filedict = {}
for file in filelist:
    isnumber = False
    number = None
    for key in numfil.keys():
        if key in file:
            isnumber = True
            number = numfil[key]
    if isnumber:
        filedict.update({number:file})

dfprev = False
for key in filedict.keys():
    file = filedict[key]
    
    fullfilename = folder_selected + '/' + file
    sheetname = getsheetname(fullfilename)
    searchfor = '|'.join(filnumb[key])
    def makedf():
        _header = 2
        _result = False
        while not _result:
            if _header < 5:
                _df = pd.read_excel(fullfilename, sheet_name=sheetname, header=_header, use_inf_as_na=False)
                try:
                    _ = df[df['РФ'].str.contains(searchfor, na=False)]
                    _result = True
                except KeyError:
                    print('Error')
                    _header += 1
                    _result = False
                return _df
            else:
                print('!!! ERROR !!! \n не найден заголовок!')
    df = makedf()
        
    df.replace(np.nan, 0, inplace=True)
    df.fillna(0, inplace=True)
    
    
    print(searchfor)
    if type(dfprev) == bool:
        dfprev = df[df['РФ'].str.contains(searchfor, na=False)]
    else:
        dfnow = df[df['РФ'].str.contains(searchfor, na=False)]
        dfprev = pd.concat([dfprev, dfnow], sort=False)

dfprev['Код проекта'] = dfprev['Код проекта'].apply(str)
dfprev['Код проекта'] = "\'" + dfprev['Код проекта'].str.zfill(14)

out_file_name = folder_selected + '/' + resultfilename
try:
    xlWriter = pd.ExcelWriter(out_file_name, engine='xlsxwriter')
    workbook = xlWriter.book
    dfprev.to_excel(xlWriter, encoding='utf-8', index=False, sheet_name='сборка')
    xlWriter.save()
    xlWriter.close()
    print('Файл ', out_file_name, ' сохранен')
except IOError:
    print("Could not open file! Please close Excel!")
