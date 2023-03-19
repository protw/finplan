print(''' 
FIN_PLAN.PY v. 0.1.2 -- Побудова фінансового плану грантового проєкту 
    Author: Olegh Bondarenko, https://protw.github.io/oleghbond
    Date: March 18, 2023
    CC BY-NC 3.0
    ''')

''' Необхідні бібліотеки '''

import pandas as pd
import numpy as np
import os, sys

from wdaysm import working_time_plan, working_days_num, purchase_time_plan
from proj_calendar import make_calendar

""" Suppress SettingWithCopyWarning """

pd.set_option('mode.chained_assignment', None)

''' Блок початкових параметрів проєкту '''

if len(sys.argv) != 2:
    print('Введіть правильну команду:')
    print('> python fin_plan.py <data_file.xlsx>')
    sys.exit('Спробуй ще.')

data_file = sys.argv[1]
'''
data_file = r'data/Task schedule (template).xlsx' # тестовий приклад-шаблон
'''
file_name, file_ext = os.path.splitext(data_file)
result_file = file_name + '_finplan' + file_ext

''' Зчитуємо дані стосовно проєкту '''

skiprows = 1
sheet_name = 'Task schedule'
tsk = pd.read_excel(data_file, sheet_name=sheet_name, skiprows=skiprows)
tsk.dropna(subset='WP', inplace=True)

skiprows = 3
sheet_name = 'Team'
team = pd.read_excel(data_file, sheet_name=sheet_name, skiprows=skiprows)
team.dropna(subset='Name', inplace=True)
team.set_index('Nick',inplace=True)

skiprows = 0
sheet_name = 'Budget'
tbudget = pd.read_excel(data_file, sheet_name=sheet_name, skiprows=skiprows)

skiprows = 0
sheet_name = 'Task types'
tsk_type = pd.read_excel(data_file, sheet_name=sheet_name, skiprows=skiprows)

skiprows = 0
sheet_name = 'Parameters'
pars = pd.read_excel(data_file, sheet_name=sheet_name, skiprows=skiprows)

''' Підготувати допоміжні дані '''

users = team.index
tot_salary = tbudget[tbudget['Budget item'] == 'Personnel costs']['Total'][0]
wdays = pd.Series([ # серія (tasks x 1) тривалості завдання у роб днях
    working_days_num(r['DateStart'], r['DateEnd']) + 1 for i, r in tsk.iterrows()
    ])

''' Project overhead, start and end dates -- overhead, dps, dpe '''

pars = pars.set_index('Parameter')['Value'] # convert from dataframe to series
dps = pars['Start date']  # project start date
dpe = pars['End date']  # project end date
overhead = pars['Overhead']  # нарахування накладних витрат
prj_name = pars['Project acronym']

''' Обчислення нормовочного коефіцієнта '''

def get_norm_coef():
    tbl_contrib = tsk[users] # таблиця (tasks x users) % роб часу user на завдання
    rel_contrib = tbl_contrib.sum(axis=1) # серія (tasks) сума % роб часу всіх users по завданню
    pers_days_task = wdays*rel_contrib/100 # серія (tasks) люд*днів на завдання
    pers_day_task_user = pd.DataFrame() # таблиця (tasks x users) люд*днів на завдання кожного user
    for user in users:
        pers_day_task_user[user] = \
            tbl_contrib[user]*pers_days_task/rel_contrib

    pers_day_user = pers_day_task_user.sum() # серія (users) люд*днів на проєкт кожного user

    wages = team['Ставка, Є/день'] # серія (users) денна ставка кожного user
    user_payment_tot = (pers_day_user*wages).sum() # сума виплат всім users за проєкт

    norm_coef = tot_salary / user_payment_tot

    return norm_coef

norm_coef = get_norm_coef()

''' Обчислення кількості роб днів по проєктних місяцях для кожного завдання 
    типу prop '''

prop_types = tsk_type.loc[tsk_type['Type']=='prop','Cat']
tsk_prop = tsk[tsk['Cat'].isin(prop_types)] # вибір завдань типу prop

task_dates = tsk_prop[['DateStart', 'DateEnd']]
task_dates.rename(columns={'DateStart': 'ds', 'DateEnd': 'de'}, inplace=True)

task_dates, drange = working_time_plan(task_dates, dps, dpe)

''' Деталізований роб план у вигляді так званої довгої таблиці '''

def get_workplan_long():
    daily_rate = team['Ставка, Є/день']
    daily_rate.name = 'Rate'
    contrib = tsk[users] / 100
    contribz = [contrib.iloc[i,:].dropna() * norm_coef for i in contrib.index]

    workplan_long = pd.DataFrame()
    for i, r in tsk_prop.iterrows():
        wdays_m = task_dates.loc[i,:]['wdays_m']
        for m, wd in wdays_m.items():
            for user, cntrb in contribz[i].items():
                row = pd.Series(dtype=object)
                row['idx'] = i # внутрішній індекс завдання
                row['prjm'] = m # номер проєктного місяця, від 1
                row['wdays'] = wd # кількість роб днів у міс
                row['tsk_idx'] = r['Tsk idx'] # індекс завдання з вх таблиці
                row['cat'] = r['Cat'] # тип завдання
                row['user'] = user # скорочене ім'я виконавця
                row['cntrb'] = cntrb # частка роб часу виконавця на проєкт
                row['tariff'] = daily_rate[user] # денна ставка виконавця
                    
                workplan_long = pd.concat([workplan_long, 
                                           pd.DataFrame(row).T])
    return workplan_long

def get_user_month_pay(workplan_long):
    ''' План щомісячних виплат кожному user '''
    pay = [r.wdays * r.cntrb * r.tariff for _, r in workplan_long.iterrows()]
    workplan_long['pay'] = pay
    return pd.pivot_table(workplan_long, values='pay', index=['user'],
                          columns =['prjm'], aggfunc = np.sum)

def get_user_month_pers_day(workplan_long):
    ''' Щомісячний план люд*днів кожного user '''
    pers_day = [r.wdays * r.cntrb for _, r in workplan_long.iterrows()]
    workplan_long['pers_day'] = pers_day
    return pd.pivot_table(workplan_long, values='pers_day', index=['user'],
                          columns =['prjm'], aggfunc = np.sum)

def direct_costs_plan():
    ''' Обчислення планів прямих витрат (з/п) і табелів '''
    
    workplan_long = get_workplan_long() # Деталізований роб план у вигляді довгої таблиці

    ''' Отримання результатів (у першому наближенні) '''

    user_month_pay = get_user_month_pay(workplan_long) # План щомісячних виплат кожному user
    user_month_pers_day = get_user_month_pers_day(workplan_long) # Щомісячний план люд*днів кожного user

    ''' Трохи скоригувати отримані результати на norm_coef '''

    user_payment_tot = user_month_pay.sum(axis=1).sum() # Повна сума нарахувань всіх users
    norm_coef = tot_salary / user_payment_tot # поправковий коеф

    user_month_pay *= norm_coef
    user_month_pers_day *= norm_coef

    ''' Підбиття підсумків '''
    
    all_users_month_pay = user_month_pay.sum() # сумарні міс виплати з/п - для загальної таблиці міс витрат
    user_month_pay.loc['TOTALS', :] = all_users_month_pay
    user_month_pay.loc[:, 'TOTALS'] = user_month_pay.sum(axis=1)

    user_month_pers_day.loc['TOTALS',:] = user_month_pers_day.sum()
    user_month_pers_day.loc[:, 'TOTALS'] = user_month_pers_day.sum(axis=1)

    ''' Загальний розподіл трудовитрат і платежів '''

    user_payment = team[['Name', 'Sub team', 'Ставка, Є/день']]
    user_payment.loc['TOTALS',:] = ''
    user_payment.loc[:, 'Pers days'] = user_month_pers_day.loc[:, 'TOTALS']
    user_payment.loc[:, 'Total payment'] = user_month_pay.loc[:, 'TOTALS']
    user_payment.loc[:, 'Contrib'] = 100 * user_payment.loc[:, 'Total payment'] \
                                   / user_payment.loc['TOTALS', 'Total payment']

    all_users_month_pay = pd.DataFrame(all_users_month_pay).T
    all_users_month_pay.rename(index={0: 'Users monthly payment'}, inplace=True)

    return (all_users_month_pay, # сумарні міс виплати з/п - для загальної таблиці міс витрат
            user_payment, # Загальний розподіл трудовитрат і платежів
            user_month_pay, # міс виплати кожного user
            user_month_pers_day) # щомісячні трудові витрати кожного user у люд*днях

(all_users_month_pay, user_payment, 
 user_month_pay, user_month_pers_day) = direct_costs_plan() 

print('Плани прямих витрат і табелі зроблено!')

''' Обчислення кількості днів по проєктних місяцях для кожного завдання 
    типу cache '''

cache_types = tsk_type.loc[tsk_type['Type']=='cache','Cat']
tsk_cache = tsk[tsk['Cat'].isin(cache_types)] # вибір завдань типу cache

task_dates = tsk_cache[['DateStart', 'DateEnd']]
task_dates.rename(columns={'DateStart': 'ds', 'DateEnd': 'de'}, inplace=True)

task_dates, drange = purchase_time_plan(task_dates, dps, dpe)

''' Деталізований план закупівель у вигляді довгої таблиці '''

def get_purchase_plan_long():
    purchase_plan_long = pd.DataFrame()
    for i, r in tsk_cache.iterrows():
        days_m = task_dates.loc[i,:]['days_m']
        months_num = len(days_m)
        for m, d in days_m.items():
            row = pd.Series(dtype=object)
            row['idx'] = i # внутрішній індекс завдання
            row['tsk_idx'] = r['Tsk idx'] # індекс завдання з вх таблиці
            row['prjm'] = m # номер проєктного місяця, від 1
            row['days'] = d # кількість кал днів завдання у міс
            row['dim'] = drange.loc[m,'ld'] # кільк кал днів у проєктному міс 
            row['cache_part'] = r['Cache'] / months_num # частка платежу за покупку
            row['cat'] = r['Cat'] # тип завдання
            
            purchase_plan_long = pd.concat([purchase_plan_long, 
                                            pd.DataFrame(row).T])
    return purchase_plan_long

purchase_plan_long = get_purchase_plan_long()
cache_month_pay = pd.pivot_table(purchase_plan_long, values='cache_part', 
                                 index=['cat'], columns =['prjm'], 
                                 aggfunc = np.sum)

''' General monthly payment table '''

general_month_pay = pd.concat([all_users_month_pay, cache_month_pay])
general_month_pay.replace(np.nan, 0, inplace=True)
general_month_pay.loc['Overhead',:] = general_month_pay.sum() * overhead

def rename_index():
    ''' Перейменувати індекси general_month_pay з коротких назв у повні 
        для завдань категорії типу cache '''
    cache_type = tsk_type.loc[tsk_type['Type'] == 'cache',:]
    for _, r in cache_type.iterrows():
        general_month_pay.rename(index={r['Cat']: r['Item']}, inplace=True)

rename_index()

general_month_pay.loc['TOTALS',:] = general_month_pay.sum()
general_month_pay.loc[:, 'TOTALS'] = general_month_pay.sum(axis=1)

''' Обчислюємо квартальні і річні підсумки та вставляємо їх у відповідні місця 
    до щомісячних трудовитрат (user_month_pers_day), виплат (user_month_pay) і 
    загального фін плану (general_month_pay) '''

def prjm_in_quarters():
    ''' Формуємо допоміжну структуру (словник), що групує проєктні місяці за 
        календарними кварталами і проєктними роками '''
    
    ''' Згрупувати проєктні місяці за календарними кварталами '''
    mbq = [list(range(q*3+1, q*3+4)) for q in range(4)] # months by quarters
    ''' Видобути проєктні роки '''
    drange1 = drange.drop(drange.iloc[-2:,:].index) # викреслити два останніх допоміжних рядки
    drange1.reset_index(inplace=True) # prjm into column
    prj_years = drange1.y.unique()
    prjmiy = {}
    for prjy in prj_years:
        drng_y = drange1.loc[drange1.y == prjy]
        prjmiq = {}
        for qi, mbqi in enumerate(mbq):
            prjm = list(drng_y.loc[drng_y.m.isin(mbqi),:].prjm)
            prjmiq[qi + 1] = prjm
        prjmiy[prjy] = {'prjm': list(drng_y.prjm), 'q': prjmiq}
        
    return prjmiy # допоміжна структура (словник) project months in year - prjmiy

prjmiy = prjm_in_quarters()

def qy_summary(df):
    ''' Власне обчислення та вставка квартальних і річних підсумків у відповідні 
        місця таблиць '''
    def sub_summary(prjm, name):
        cols_all = df.columns
        cols = cols_all[cols_all.isin(prjm)]
        ishift = 1 if 'Q' in name else 2
        icol = pd.Index(cols_all).get_loc(cols.max()) + ishift
        tot = df.loc[:,cols].sum(axis=1)
        tot_name = name
        df.insert(loc=icol, column=tot_name, value=tot)
        
    for y, yv in prjmiy.items():
        for q, qv in yv['q'].items():
            sub_summary(qv, f'Q{q:d}/Y{y}')
        sub_summary(yv['prjm'], f'Y{y}')
        
    return

qy_summary(user_month_pay)
qy_summary(user_month_pers_day)
qy_summary(general_month_pay)

print('Загальні місячні витрати зроблено!')

''' Створення таблиці дат подій для розміщення у Гугл Календарі '''

calendar = make_calendar(tsk, users, prj_name)

print('І ще бонус - календар проєкту зроблено!')

''' Зберегти всі таблиці в окремих аркушах ексель-файлу '''

with pd.ExcelWriter(result_file, engine='xlsxwriter') as writer:
    user_payment.to_excel(writer, sheet_name='Personal contribution')
    user_month_pay.to_excel(writer, sheet_name='Monthly pers payment, Euro')
    user_month_pers_day.to_excel(writer, sheet_name='Tabel, pers-day')
    general_month_pay.to_excel(writer, sheet_name='All monthly payments')
    calendar.to_excel(writer, sheet_name='Google calendar', index=False)

print(f'Всі дані збережено у файлі `{result_file}`!')

