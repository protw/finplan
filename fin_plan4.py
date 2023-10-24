descr = ''' 
FIN_PLAN.PY v. 0.2.2 -- Побудова фінансового плану грантового проєкту 
    Author: Olegh Bondarenko, https://protw.github.io/oleghbond
    Date: March 18, 2023
    CC BY-NC 3.0
    '''
print(descr)

print(''' Необхідні бібліотеки ''')

import pandas as pd
import numpy as np
import os, sys, glob

from fin_plan_util import (read_input_data, validate_input_data,
                           create_two_long_tables, gen_pers_days_payment,
                           join_two_long_tables)
from fin_plan_time_util import set_proj_months
from pivot_subtotals import pivot_w_subtot2
from proj_calendar import make_calendar


def get_file_names(data_dir:str=''):
    ''' Коли функція знаходить CLI аргумент, вона бере його як ім'я файлу даних.
        Коли CLI аргумент відсутній, функція зчитує свій вх аргумент як ім'я
        теки, де шукає один .xlsx-файл вх даних. '''
    if len(data_dir) == 0 & len(sys.argv) >= 2:
        data_file = sys.argv[1]
    elif len(data_dir) > 0:
        if not os.path.isdir(data_dir):
            sys.exit(f'Немає такої теки: `{data_dir}`')
        file_list = glob.glob(os.path.join(data_dir, '*.xlsx'))
        if len(file_list) != 1:
            sys.exit(f'Має бути лише 1 XLSX-файл у теці: `{data_dir}`')
        data_file = file_list[0]
        
    print(' == ', 'data_file: ', data_file)
    
    file_name, file_ext = os.path.splitext(data_file)
    if file_ext.lower() != '.xlsx':
        sys.exit(f'Файл має мати розширення ".XLSX": `{data_file}`')
    result_file = file_name + '_finplan' + file_ext
    
    return data_file, result_file

def finplan_main(data_dir:str=''):
    print(''' == БЛОК ПОЧАТКОВИХ ПАРАМЕТРІВ ПРОЄКТУ == ''')
    
    data_file, result_file = get_file_names(data_dir)
    
    print(''' == ЗЧИТУЄМО ДАНІ СТОСОВНО ПРОЄКТУ == ''')
    
    tsk_short_names = {'Tsk idx': 'id', 'WP': 'wp', 'Cat': 'cat', 'Level': 'lev', 
                    'Task/Deliv #': 'no','DateStart': 'ds', 'DateEnd': 'de', 
                    'Duration': 'dur', 'Cache': 'cache', 
                    'Description of task': 'descr'}
    
    tsk, team, tbudget, tsk_type, pars, wppm = read_input_data(data_file, tsk_short_names)
    
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
    
    team, tsk_prop = gen_pers_days_payment(team, tsk_prop, tot_salary, wppm)
    
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
    
    wp_month_pers_day = pivot_w_subtot2(df=tsk_prop,
                                        values='pers_day',
                                        indices=['wp'],
                                        columns=['y', 'q', 'prjm'],
                                        aggfunc=np.nansum)
    wp_month_pers_day.rename(index={'':'TOTAL'}, columns={'':'TOTAL'}, inplace=True)
    
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
        wp_month_pers_day.to_excel(writer, sheet_name='WP labour efforts, pers-day')
        general_month_pay.to_excel(writer, sheet_name='All monthly payments')
        calendar.to_excel(writer, sheet_name='Google calendar', index=False)
    
    print(f'Всі дані збережено у файлі `{result_file}`!')
    print('DONE!')

if __name__ == '__main__':
    data_dir = 'data'
    finplan_main(data_dir)
