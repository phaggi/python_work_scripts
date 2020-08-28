import pandas as pd
import numpy as np
from copy import copy
from pprint import pprint
import datetime

in_file_RF = 'Владимир_8_2.xlsx'
in_name_sheet_RF = 'Таблица для заполнения '
in_file_AUK = 'c://in_test/kal.xlsx'
in_name_sheet_AUK = 'прогноз_август'
in_file_god = 'god.xlsx'
in_name_sheet_god = 'года'
col_from_np = 'Naselennyj_punkt'



load_list = {'RF': [in_file_RF, in_name_sheet_RF], 'AUK': [in_file_AUK, in_name_sheet_AUK], 'god': [in_file_god, in_name_sheet_god]}


def name_for_read(income_file, searched_word):
    for sheet_name in income_file.sheet_names:
        if searched_word in sheet_name:
            return sheet_name


def load_table(load_punkt, load_list = load_list):
    in_file = load_list[load_punkt][0] 
    searched_word = load_list[load_punkt][1]
    income_file = pd.ExcelFile(in_file)
    sheet_name = name_for_read(income_file, searched_word)
    return income_file.parse(sheet_name), sheet_name


df_RF, sheetname_RF = load_table('RF')

df_AUK, sheetname_AUK = load_table('AUK')

df_god, sheetname_god = load_table('god')
df_god.drop(df_god.columns[3:],axis='columns', inplace=True)
df_god.columns = df_god.iloc[0]
df_god.drop(0, axis='rows', inplace=True)
df_god.reset_index(inplace=True)

def first_income(lst = ['f',2,3,'f',5,1], tgt = 'f'):
    if type(lst) is not type([]):
        return False
    found=[]
    for index, suspect in enumerate(lst):
        if(tgt==suspect):
            found.append(index)
    if found == []:
        return False
    else:
        return min(found)
    

def is_first_income(lst = ['f',2,3,'f',5,1], num_line = 0):
    if (type(lst) is not type([])) & (type(num_line) is not type(0)):
        return None
    f_income = first_income(lst, lst[num_line])
    for index, suspect in enumerate(lst):
        if(index == num_line == f_income):
            return True
    return False





def colnum_transliterator(df):
    # transliteration column names: text.translate(tr)
    symbols = (u"абвгдеёжзийклмнопрстуфхцчшщъыьэюяАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ /№(),*.-",
               u"abvgdeejzijklmnoprstufhzcss_yjeuaABVGDEEJZIJKLMNOPRSTUFHZCSS_YJEUA__N______")

    tr = {ord(a):ord(b) for a, b in zip(*symbols)}

    column_list = {}
    column_rus_list = {}
    for col in df.columns:
        col_rus = copy(col)
        col_lat = copy(str(col).translate(tr))
        column_list.update({col_rus: col_lat})
        column_rus_list.update({col_lat:col_rus})
    return df.rename(columns=column_list), column_rus_list, column_list
    # end transliteration modul


def make_column_cyph_list(df):
    column_cyph_list = {}
    for col in df.columns:
        col_num = df[col][1]
        try:
            col_num = str(int(col_num))
        except:
            col_num = str(col_num)
        column_cyph_list.update({col_num:col})
    return column_cyph_list

df_RF, column_rus_list_RF, column_list_RF = colnum_transliterator(df_RF)

df_AUK, column_rus_list_AUK, column_list_AUK = colnum_transliterator(df_AUK)
df_god, column_rus_list_god, column_list_god = colnum_transliterator(df_god)


df_RF = df_RF.drop(0)  # drop None str
#df_RF = df_RF.astype(object).replace(np.nan, 'None')
# column_numbers to column_cyph_list and delete column_numbers from df_RF
column_cyph_list = make_column_cyph_list(df_RF)

df_RF = df_RF.drop(1)
df_RF = df_RF.reset_index()

print(len(df_RF) == len(df_AUK))

df_AUK['Tocka_A_MCC__ukazyvaetsa_dla_pervogo_SZO_v_dannom_NP_'] = df_RF[column_cyph_list['38']]

col_to = 'Trebuemaa_protajennostj_MSS_VOLS__km__ukazyvaetsa_dla_pervogo_SZO_v_dannom_NP_'
col_from = column_cyph_list['201']
col_from_2 = column_cyph_list['180']
df_AUK[col_to] = df_RF[col_from].astype('float') + df_RF[col_from_2].astype('float')

col_to = 'Trebuemaa_protajennostj_VOLS_seti_dostupa__km'
col_from = column_cyph_list['202']
df_AUK[col_to] = df_RF[col_from].astype('float')

col_to = 'Planiruemyj_tip_podklucenia_SZO__1___po_VOLS__2___sputnikovoj_liniej__3___podklucaet_aljternativnyj_operator_po_poruceniu_Rostelekomoma_'
word_to = 1
df_AUK[col_to] = word_to

col_to = 'Podtverjdenie_praviljnosti_adresa_SZO'
col_from = column_cyph_list['9']
df_AUK[col_to] = df_RF[col_from]

col_to = 'Pervonacaljnyj_UIN'
col_from = 'ID_ob_ekta'
df_AUK[col_to] = df_RF[col_from]

col_to = 'Priznak_NP__1___tupikovyj__0___tranzitnyj_'
word_to = 1
tips_for_fill_1 = ['нет', 'нет данных', 'Нет', 'НЕТ', 'NaN', 'None', np.nan, None]
tips_for_not_fill = [1, 0, '1', '0']
print('\nВ колонке', col_to, 'следующие данные:')
print(df_AUK[col_to].head(), '\n')
if df_AUK[col_to][0] in tips_for_fill_1:
    df_AUK[col_to] = word_to
    print('Заполняем колонку Признак НП единичками')
elif df_AUK[col_to][0] not in tips_for_not_fill:
    print('\n!!!Проверь колонку про Признак НП тупиковой!!!\n')
else:
    print('Оставляем колонку про Признак НП как есть.')

col_to_old = 'Tip_uzla_konzentrazii__1___susestvuusij__pomesenie__kontejner__klimaticeskij_skaf___2___novyj__pomesenie__kontejner__klimaticeskij_skaf____ukazyvaetsa_dla_pervogo_SZO_v_dannom_NP_'
col_to_new = 'Tip_uzla_konzentrazii__1___susestvuusij__pomesenie__kontejner__klimaticeskij_skaf___2___novyj__pomesenie__kontejner__klimaticeskij_skaf___3___kombouzel___ukazyvaetsa_dla_pervogo_SZO_v_dannom_NP_'
col_from = 'Novyh_STK'
col_from_new = 'Nujna_stojka_v_SZO'
df_from = df_RF

if col_to_old in df_AUK.columns:
    print('\nСТАРЫЙ ТИП УЗЛА КОНЦЕНТРАЦИИ - проверь\n')
    col_to = col_to_old
elif col_to_new in df_AUK.columns:
    col_to = col_to_new
else:
    print('\nПроверь колонку ТИП УЗЛА КОНЦЕНТРАЦИИ\n')
    
def make_TIP_lst(df_from = df_from, col_from = col_from, col_from_np = col_from_np, 
                 df_from_out_date = df_from, col_from_out_date = col_from, default_result = 0):
    col_from_lst = list(df_from[col_from])
    col_from_np_lst = list(df_from[col_from_np])
    col_from_out_date_lst = list(df_from_out_date[col_from_out_date])
    col_to_lst = []
    for i in range(len(col_from_lst)):
        if col_from_lst[i] != 'None':
            if is_first_income(col_from_np_lst, i):
                result = col_from_out_date_lst[i]
            else:
                result = default_result
        else:
            result = None
        col_to_lst.append(result)
    return col_to_lst


tip_list1 = make_TIP_lst(df_from, col_from, col_from_np, df_from, col_from, 0)
tip_list2 = make_TIP_lst(df_from, col_from_new, col_from_np, df_from, col_from_new, 0)
df_AUK[col_to] = pd.Series(tip_list1) + pd.Series(tip_list2) * 2 + 1

col_to = 'Trebuemoe_kolicestvo_kommutatorov_konzentrazii_v_sostave_uzla_konzentrazii__ed_'
col_from = 'Itogo_po_KK'
df_AUK[col_to] = df_RF[col_from]

col_to = 'Dopolniteljnye_zatraty__rub___slojnye_i_bolsie_bolee_500_m_GNB_perehody__OOPT_i_t_p____'
col_from = column_cyph_list['210']
df_AUK[col_to] = df_RF[col_from].astype('float')

col_to = 'Vozmojnostj_okazania_uslug_soglasno_TZ_na_susestvuusej_mednoj_linii__pri_uslovii_zameny_oborudovania__1__da__0___net'
word_to = 'нет'
df_AUK[col_to] = word_to

col_to = 'Primecanie'
word_to = np.nan
df_AUK[col_to] = word_to

col_to = 'Dejstvuusee_vklucenie___Da__Net_'
col_from = column_cyph_list['10']
df_AUK[col_to] = df_RF[col_from]

col_to = 'Tehonologia_dejstvuusego_vklucenia__Medj__VOLS__RRL__WiMax__Procee_utocnitj___'
col_from = column_cyph_list['10']
df_from = df_RF
df_to = df_AUK


def make_fill_col_list(df_from = df_from, col_from = col_from):
    result = []
    true_lines = {'ДА', 'да', 'Да', 'дА'}
    for i in range(len(df_from[col_from])):
        line = df_from.loc[i, col_from] 
        if line in true_lines  :
            out = 'ВОЛС'
        else:
            out = np.nan
        result.append(out)
    return(result)


df_to[col_to] = make_fill_col_list()


col_to = 'Privazka_uzla_konzentrazii__ukazatj_UIN_SZO__v_kotorom_uctena_tocka_A_MSS_'
col_by = 'ID'
col_from = 'Finaljnyj_UIN'
col_for = 'Privazka_k_ID_dla_KK'
df_from = df_RF
df_to = df_AUK

def make_fill_privjaz_list(df_to = df_to, df_from = df_from, col_from = col_from, col_by = col_by, col_for = col_for):
    result = []
    from_dict = {}
    for i in range(len(df_to[col_from])):
        key_dict = df_from.loc[i, col_by]
        value_dict = df_to.loc[i, col_from]
        from_dict.update({key_dict:value_dict})
    for i in df_from[col_for]:
        if (i == 0) | (i == 'None') | (i == 'NaN'):
            out = np.nan
        else:
            out = from_dict[i]
        result.append(out)
    return(result)

df_to[col_to] = make_fill_privjaz_list()


col_to_chng = 'Tocka_A_MCC__ukazyvaetsa_dla_pervogo_SZO_v_dannom_NP_'
col_from = column_cyph_list['38']
col_from_np = 'Naselennyj_punkt'
df_from = df_RF
df_from_out_date = df_AUK
col_from_out_date = 'Finaljnyj_UIN'


def make_MSS_lst(df_from = df_from, col_from = col_from, col_from_np = col_from_np, 
                 df_from_out_date = df_from_out_date, col_from_out_date = col_from_out_date):
    col_from_lst = list(df_from[col_from])
    col_from_np_lst = list(df_from[col_from_np])
    col_from_out_date_lst = list(df_from_out_date[col_from_out_date])
    col_to_lst = []
    for i in range(len(col_from_lst)):
        # print(': ', col_from_lst[i], col_from_np_lst[i], ' :\t', end='')
        if col_from_lst[i] != 'None':
            if is_first_income(col_from_np_lst, i):
                result = col_from_out_date_lst[i]
            else:
                result = None
        else:
            result = None
        col_to_lst.append(result)
    return col_to_lst



df_AUK['Privazka_MSS__ukazatj_UIN_SZO__v_kotorom_uctena_tocka_A_MSS_'] = pd.Series(make_MSS_lst())



col_to = 'Planiruemyj_god_podklucenia'

df_from = df_god
def take_god(df_to = df_AUK, df_from = df_god, col_to = col_to):
    for i in df_to.index:
        key = df_to.loc[i, 'Sub_ekt_Rossijskoj_Federazii'] + str(df_to.loc[i,'N_p_p'])
        df_to.loc[i,col_to] = str(df_from.god[df_from.ID == key]).split()[1]
        
    # df_AUK.loc[i,col_to] = 1

take_god()





print('Done')
save_pin = 1

if save_pin == 1:
    
    def make_date_time():
        from datetime import datetime
        dt = str(datetime.today())
        date = dt.split()[0]
        time_lst = dt.split()[1].split(':')[0:2]
        time = ''
        for i in range(2):
            time += '_' + time_lst[i]

        return(date + time)

    datetime = make_date_time()

    df_AUK.replace('None', np.nan, inplace=True)
    df_AUK.rename(columns=column_rus_list_AUK, inplace=True)
    
    def make_name_xlsx(sheetname_AUK = sheetname_AUK, date = datetime):
        first_name = ''
        sheetname_lst = sheetname_AUK.split()
        for i in range(2):
            first_name += sheetname_lst[i] + '_'
        return(first_name + date + '.xlsx')
    
    name_file = make_name_xlsx()
    xlWriter = pd.ExcelWriter(name_file, engine='xlsxwriter')
    workbook = xlWriter.book
    df_AUK.to_excel(xlWriter, encoding='utf-8', index=False, sheet_name='Sheet{}'.format(1))

    xlWriter.save()
    xlWriter.close()
    print(name_file, 'saved.')
else:
    print('AUK_not_saved_')
