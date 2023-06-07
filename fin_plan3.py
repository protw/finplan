print(''' 
FIN_PLAN.PY v. 0.2.1 -- Побудова фінансового плану грантового проєкту 
    Author: Olegh Bondarenko, https://protw.github.io/oleghbond
    Date: March 18, 2023
    CC BY-NC 3.0
    ''')

print(''' Необхідні бібліотеки ''')

import pandas as pd
import numpy as np
import os, sys

from fin_plan_util import (read_input_data, validate_input_data,
                           create_two_long_tables, gen_pers_days_payment,
                           join_two_long_tables)
from fin_plan_time_util import set_proj_months
from pivot_subtotals import pivot_w_subtot2
from proj_calendar import make_calendar

print(''' == БЛОК ПОЧАТКОВИХ ПАРАМЕТРІВ ПРОЄКТУ == ''')

debug = True #False

def get_file_names():
    if not debug:
        if len(sys.argv) != 2:
            print('Введіть правильну команду:')
            print('> python fin_plan.py <data_file.xlsx>')
            sys.exit('Спробуй ще.')
    
        data_file = sys.argv[1]
    else:
        data_file = r'data/Task schedule (template).xlsx' # тестовий приклад-шаблон
        data_file = r'D:/boa_uniteam/UNITEAM/_PROJ/2_NOW/210703 EFACA/РОБОЧІ ПЛАНИ/'
        data_file += '230605 EFACA Task schedule - CEPA (4).xlsx'
    
    file_name, file_ext = os.path.splitext(data_file)
    result_file = file_name + '_finplan' + file_ext
    
    return data_file, result_file

data_file, result_file = get_file_names()

print(''' == ЗЧИТУЄМО ДАНІ СТОСОВНО ПРОЄКТУ == ''')

tsk_short_names = {'Tsk idx': 'id', 'WP': 'wp', 'Cat': 'cat', 'Level': 'lev', 
                   'Task/Deliv #': 'no','DateStart': 'ds', 'DateEnd': 'de', 
                   'Duration': 'dur', 'Cache': 'cache', 
                   'Description of task': 'descr'}

tsk, team, tbudget, tsk_type, pars = read_input_data(data_file, tsk_short_names)

print(''' == ПІДГОТУВАТИ ДОПОМІЖНІ ДАНІ == ''')

users = list(team.index)
tot_salary = tbudget.loc['Personnel costs']['Total']

''' Project overhead, start and end dates -- overhead, dps, dpe '''

pars = pars.set_index('Parameter')['Value'] # convert from dataframe to series
dps = pars['Start date']  # project start date
dpe = pars['End date']  # project end date
overhead = pars['Overhead']  # нарахування накладних витрат
prj_name = pars['Project acronym']
''' Таблиця проєктних місяців містить на кожний місяць проєкту значення
    {'y': years, 'm': months, 'q': quarters, 'prjm': prj_months, 
     'ld': last_day_in_month, 'wd': working_days_in_month}
    prjm встановлено в якості індекса'''
prj_months = set_proj_months(dps, dpe).set_index('prjm')

print(f'Фінплан проєкту `{prj_name}`')

print(''' == ВАЛІДАЦІЯ ВХІДНИХ ДАНИХ == ''')

validate_input_data(tsk, tbudget, tsk_short_names, dps, dpe)

# Викреслюємо повністю порожні рядки у стовпчиках 'cache' та `users`
tsk = tsk.loc[tsk[['cache',] + users].any(axis=1)]

print(''' == СТВОРЮЄМО ДВІ ДОВГИХ ТАБЛИЦІ ДЛЯ ПРОПОРЦІЙНИХ І КЕШ-ЗАВДАНЬ == ''')

tsk_prop, tsk_cache = create_two_long_tables(tsk, prj_months, users, tsk_type)

print(''' == ЗАГАЛЬНІ ТРУДОВИТРАТИ І ВИПЛАТИ == ''')

team = gen_pers_days_payment(team, tsk_prop, tot_salary)

print(''' == ЩОМІСЯЧНІ ТРУДОВИТРАТИ І ВИПЛАТИ == ''')

user_month_pay = pivot_w_subtot2(df=tsk_prop, 
                                 values='pers_pay', 
                                 indices=['user'], 
                                 columns=['y', 'q', 'prjm'], 
                                 aggfunc=np.nansum)
user_month_pay.rename(index={'':'TOTAL'}, columns={'':'TOTAL'}, inplace=True)

user_month_pers_day = pivot_w_subtot2(df=tsk_prop, 
                                 values='pers_day', 
                                 indices=['user'], 
                                 columns=['y', 'q', 'prjm'], 
                                 aggfunc=np.nansum)
user_month_pers_day.rename(index={'':'TOTAL'}, columns={'':'TOTAL'}, inplace=True)

print(''' == ЗАГАЛЬНІ ЩОМІСЯЧНІ ПОСТАТЕЙНІ ВИТРАТИ == ''')

tsk_cache_upd = join_two_long_tables(tsk_prop, tsk_cache, tsk_type)

# Підрахуємо накладні на кожний проєктний місяць
overhead_month = tsk_cache_upd.groupby('prjm').aggregate({'cache':np.sum}) * overhead
overhead_month = pd.concat([prj_months[['y', 'q']], overhead_month], axis=1)
overhead_month['cat'] = 'Overheads'
overhead_month.reset_index(inplace=True)

# Включимо накладні до загальних витрат
tsk_cache_upd = pd.concat([tsk_cache_upd, overhead_month])

general_month_pay = pivot_w_subtot2(df=tsk_cache_upd, 
                                    values='cache', 
                                    indices=['cat'], 
                                    columns=['y', 'q', 'prjm'], 
                                    aggfunc=np.nansum)
general_month_pay.rename(index={'':'TOTAL'}, columns={'':'TOTAL'}, inplace=True)

print(''' == Створення таблиці дат подій для розміщення у Гугл Календарі == ''')

tsk.rename(columns={'no': 'Task/Deliv #', 'ds': 'DateStart', 'de': 'DateEnd',
                    'descr': 'Description of task'}, inplace=True)
calendar = make_calendar(tsk, users, prj_name)

''' Зберегти всі таблиці в окремих аркушах ексель-файлу '''

with pd.ExcelWriter(result_file, engine='xlsxwriter') as writer:
    team.to_excel(writer, sheet_name='Personal contribution')
    user_month_pay.to_excel(writer, sheet_name='Monthly pers payment, currency')
    user_month_pers_day.to_excel(writer, sheet_name='Tabel, pers-day')
    general_month_pay.to_excel(writer, sheet_name='All monthly payments')
    calendar.to_excel(writer, sheet_name='Google calendar', index=False)

print(f'Всі дані збережено у файлі `{result_file}`!')
print('DONE!')


