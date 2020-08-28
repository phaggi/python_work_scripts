import pandas as pd
from path import Path
from pprint import pprint
import xlsxwriter

def writexlsx(_outpath, _df):
    writer = pd.ExcelWriter(_outpath, engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    _df.to_excel(writer, sheet_name='Sheet1')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

mypath = '/Users/phaggi/PythonNotes/belov/'
infile = 'intable.xlsx'
inpath = Path(mypath + infile)
outfile = 'outtable.xlsx'
outpath = Path(mypath + outfile)

df = []
xl = pd.ExcelFile(inpath)


df = []
for sheet in xl.sheet_names:
    df.append(xl.parse(sheet_name=sheet, index_col='ID'))

tele2 = df[0]
zakl = df[1]

tele2_filtered = tele2[tele2['№ БС'].isin(zakl['№ БС'])]

tele2path = Path(mypath + 'tele2_filtered.xlsx')
writexlsx(tele2path, tele2_filtered)

tele2_punkts = tele2_filtered['Пункты договора']

punkts = tele2_punkts.to_list()

splitpunkts = [punkt.split('\n') for punkt in punkts]

finsplitpunkts = []
for punkt in splitpunkts:
    for subpunkt in punkt:
        if subpunkt != '':
            finsplitpunkts.append(subpunkt)

fin_punkts = dict()
for i in range(len(finsplitpunkts)):
    key, value = finsplitpunkts[i].split('(')
    key = key.strip()
    value = float(value.split(')')[0])
    if key not in fin_punkts.keys():
        counter = 1
        newvalue = value
    else:
        counter = fin_punkts[key]['counter'] + 1
        newvalue = fin_punkts[key]['summa'] + value
    newvalue = round(newvalue, 2)
    fin_punkts.update({key: {'counter': counter, 'summa': newvalue}})

punkt_count = pd.DataFrame(fin_punkts).transpose()
punkt_count = punkt_count.sort_index()

writexlsx(outpath, punkt_count)
