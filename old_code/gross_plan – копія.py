from glob import glob
import os
import pandas as pd
import numpy as np
from datetime import datetime

from pivot_subtotals import pivot_w_subtot3

# Вхідні параметри процедури побудови Загального Фінплану
start_month = (2024, 1)
cur_sel = 'EUR'
print(f'GROSS FINPLAN\nПочаток: {start_month}, валюта розрахунку: {cur_sel}')

# Роб теки
finplans_dir = r"../../240305 GROSSPLAN/Finplans"
reports_dir = r"../../240305 GROSSPLAN/Reports"

# Формуємо ім'я результуючого звіту
now = datetime.now()
rep_suffix = datetime.strftime(now, '%y%m%d-%H%M%S')
start_label = f'{start_month[0]-2000}{start_month[1]:02d}'
report_file = f'gross_finplan-{start_label}-{rep_suffix}.xlsx'
report_file = os.path.join(reports_dir, report_file)

# Обрані стовпчики у вх таблицях (фінплани окремих проєктів)
cols = ['prjm', 'y', 'm', 'q', 'cat', 'user', 'pers_day', 
        'cache', 'prj_name', 'currency']
cols_simple = ['y', 'm', 'cat', 'user', 'pers_day', 
        'cache', 'prj_name', 'currency']

print('== ЗАГАЛЬНА ІНФО ПРО ПРОЄКТИ, ЛЮДЕЙ, КУРС ВАЛЮТ ==')
prj_usr_cur_path = r'D:/boa_uniteam/UNITEAM/_DOCS/FIN/230708 GROSSBUCH/Загальне'
prj_usr_cur_path = os.path.join(prj_usr_cur_path, 'Проєкти-Люди-Курс.xlsx')

prj_usr_cur = pd.read_excel(prj_usr_cur_path, sheet_name=None)

prj = prj_usr_cur['Проєкти']
usr = prj_usr_cur['Люди'][['nick', 'pib']]
cur = prj_usr_cur['Курс']

def check_currency():
    ''' Перевірка валют. Нормалізація курса валют відносно обраної 'cur_sel'. '''
    # Перевірка валют
    cur['UAH'] = 1
    cur_list = cur.columns[2:].to_list()
    assert cur_sel in cur_list, f'Валюта {cur_sel} відсутня серед валют проєктів'
    # Нормалізація курса валют
    start_month_idx = start_month[0] * 100 + start_month[1]
    cur.index = cur.year * 100 + cur.month
    if cur.index.max() >= start_month_idx:
        cur_tbl = cur[cur_list].loc[cur.index >= start_month_idx, :]
    else:
        cur_tbl = cur[cur_list].loc[cur.index.max(), :].to_frame().T
    cur_tbl = cur_tbl.div(cur_tbl.loc[:, cur_sel], axis=0) # Нормалізація
    return cur_list, cur_tbl

def check_tbl():
    ''' Фільтрувати рядки в 'tbl', де місяць дорівнює або пізнше 'start_month'.
    Валідація назви проєкту. Валідація коротких імен учасників. Валідація назви 
    валюти проєкту. Вставляємо до таблиці 'tbl' курс валюти на кожний місяць 
    відносно 'cur_sel'. '''
    # Перевірити набір стовпчиків
    if set(cols).issubset(set(tbl.columns)):
        print(f'Повний фінплан {file}')
    elif set(cols_simple) == set(tbl.columns):
        tbl['prjm'] = None
        tbl['q'] = (tbl.m - 1) // 3
        print(f'Спрощений фінплан {file}')
    else:
        f'У таблиці "{file}" має бути набір стовчиків: {cols} або {cols_simple}'
    # Фільтрувати рядки в 'tbl', де місяць дорівнює або пізнше 'start_month'
    ym_idx = (start_month[0] * 100 + start_month[1]) > (tbl.y * 100 + tbl.m)
    tbl.drop(ym_idx[ym_idx].index, inplace=True)
    if len(tbl) == 0: 
        print(f'-- {os.path.split(file)[1]} -- пропущено, завершено раніше {start_month}')
        return False
    # Валідація назви проєкту
    prj_name = tbl.prj_name.unique()
    assert len(prj_name) == 1 and any(prj_name[0].upper() == prj.acronym), \
        f'У таблиці "{file}" проєкт {prj_name} відсутній у списку проєктів організації'
    # Валідація коротких імен учасників
    nicks_prj = set(tbl.user[~pd.isnull(tbl.user)].unique())
    nicks_org = set(usr.nick)
    assert nicks_prj.issubset(nicks_org), \
        f'У таблиці "{file}" люди {nicks_prj-nicks_org} відсутні у списку організації'
    # Валідація назви валюти проєкту
    cur_name = tbl.currency.unique()
    assert len(cur_name) == 1 and cur_list, \
        f'У таблиці "{file}" валюта {cur_name} відсутня серед валют проєктів'
    # Вставляємо до таблиці 'tbl' курс валюти на кожний місяць відносно 'cur_sel'
    tbl.index = tbl.y * 100 + tbl.m
    tbl['cur_rate'] = None
    ym_common = list(set(tbl.index) & set(cur_tbl.index))
    if len(ym_common) >= 1:
        tbl.loc[ym_common, 'cur_rate'] = cur_tbl.loc[ym_common, cur_sel]
        tbl.loc[pd.isnull(tbl['cur_rate']), 'cur_rate'] = cur_tbl.loc[max(ym_common), cur_sel]
    else:
        tbl['cur_rate'] = cur_tbl.loc[max(cur_tbl.index), cur_sel]
    
    return True

cur_list, cur_tbl = check_currency()

print('== ДОДАЄМО ФІНПЛАНИ ОКРЕМИХ ПРОЄКТІВ, ПРИВОДИМО ПЛАТЕЖІ ДО {cur_sel} ==')
finplans_list = glob(os.path.join(finplans_dir, '*_finplan.xlsx'))
cols_ = cols + ['cur_rate',]
gross_tbl = pd.DataFrame(columns=cols_)
for i, file in enumerate(finplans_list):
    tbl = pd.read_excel(file, sheet_name='Long table')
    if check_tbl():
        gross_tbl = tbl[cols_] if i == 0 else pd.concat([gross_tbl, tbl[cols_]])
        print(f'-- {os.path.split(file)[1]}')

# Приводимо всі платежі до обраної валюти 'cur_sel'
gross_tbl.cache *= gross_tbl.cur_rate 

print('== ОБРАХУВАННЯ ПІДСУМКІВ ==')
print('-- щомісячні виплати')
user_month_pay = pivot_w_subtot3(df=gross_tbl.loc[~pd.isnull(gross_tbl.user), :], 
                                 values='cache', 
                                 #indices=['user', 'prj_name'], 
                                 indices=['user'], 
                                 columns=['y', 'q', 'm'], 
                                 aggfunc=np.nansum)
print('-- щомісячні трудовитрати')
user_person_month = pivot_w_subtot3(df=gross_tbl, 
                                 values='pers_day', 
                                 #indices=['user', 'prj_name'], 
                                 indices=['user'], 
                                 columns=['y', 'q', 'm'], 
                                 aggfunc=np.nansum)
print('-- щомісячні постатейні витрати')
cat_month_pay = pivot_w_subtot3(df=gross_tbl, 
                                 values='cache', 
                                 indices=['cat', 'prj_name'], 
                                 columns=['y', 'q', 'm'], 
                                 aggfunc=np.nansum)

''' Зберегти всі таблиці в окремих аркушах ексель-файлу '''

with pd.ExcelWriter(report_file, engine='xlsxwriter') as writer:
    user_month_pay.to_excel(writer, sheet_name=f'Monthly payment, {cur_sel}')
    user_person_month.to_excel(writer, sheet_name='Person-months, pers-day')
    cat_month_pay.to_excel(writer, sheet_name=f'All monthly payments, {cur_sel}')

print(f'Всі дані збережено у файлі `{report_file}`!')
print('== DONE! ==')
