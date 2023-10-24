import pandas as pd

from fin_plan_time_util import (check_dates_consistency, 
                                wdays_in_each_task_month)


def read_input_data(data_file, tsk_short_names):
    skiprows = 1
    sheet_name = 'Task schedule'
    tsk = pd.read_excel(data_file, sheet_name=sheet_name, skiprows=skiprows)
    tsk.rename(columns=tsk_short_names, inplace=True)
    
    skiprows = 3
    sheet_name = 'Team'
    team = pd.read_excel(data_file, sheet_name=sheet_name, skiprows=skiprows)
    team.dropna(subset='Name', inplace=True)
    team.set_index('Nick', inplace=True)
    
    skiprows = 0
    sheet_name = 'Budget'
    tbudget = pd.read_excel(data_file, sheet_name=sheet_name, skiprows=skiprows)
    tbudget.set_index('Budget item', inplace=True)
    
    skiprows = 0
    sheet_name = 'Task types'
    tsk_type = pd.read_excel(data_file, sheet_name=sheet_name, skiprows=skiprows)
    
    skiprows = 0
    sheet_name = 'Parameters'
    pars = pd.read_excel(data_file, sheet_name=sheet_name, skiprows=skiprows)
    
    skiprows = 0
    sheet_name = 'WP-PM'
    wppm = pd.read_excel(data_file, sheet_name=sheet_name, skiprows=skiprows)
    
    return tsk, team, tbudget, tsk_type, pars, wppm

def validate_input_data(tsk, tbudget, tsk_short_names, dps, dpe):
    def get_dict_key_from_value(d, v):
        return list(d.keys())[list(d.values()).index(v)]
    # Перевірка відсутності незаповнених елементів
    cols = ['id', 'wp', 'cat', 'no', 'ds', 'de']
    cols_old = [get_dict_key_from_value(tsk_short_names, c) for c in cols]
    if tsk[cols].isnull().any().any():
        raise Exception(f'Заповніть спочатку порожні клітинки у стовпчиках {cols_old}')
    # Перевірка дат проєкту і завдань
    check_dates_consistency(tsk, dps, dpe)
    # Перевірка бюджету
    budget_items = list(tbudget.index)
    budget_items.remove('TOTAL')
    budget_total = tbudget.loc['TOTAL'].Total # Загальний бюджет з таблиці
    budget_items_sum = tbudget.loc[budget_items].Total.sum() # Сума статей бюджету
    if budget_total != budget_items_sum:
        raise Exception(f'Загальний бюджет {budget_total} не збігається із ' + \
                        f'сумою статей бюджету {budget_items_sum}')
    # Перевірка відсутності дублікатів у стовпчиках 'id' і 'no'
    def check_dupl_in_col(col_name):
        col_old = get_dict_key_from_value(tsk_short_names, col_name)
        tsk_id = tsk[col_name]
        tsk_id_dupl = tsk_id[tsk_id.duplicated()].index
        if len(tsk_id_dupl) > 0:
            raise Exception(f'У стовпчику `{col_old}` є дублікати ' + \
                            f'{tsk_id.iloc[tsk_id_dupl]}')
    check_dupl_in_col('id')
    check_dupl_in_col('no')
    
    return

def create_two_long_tables(tsk, prj_months, users, tsk_type):
    ''' Створюємо дві довгих таблиці для пропорційних і кеш-завдань '''
    # Обчислюємо кількість робочих днів 'twd' для кожного місяця, на який припадає 
    # завдання (додатково до інших параметрів проєктного місяця з `prj_months`)
    tsk_months = wdays_in_each_task_month(tsk, prj_months)
    # Встановлюємо 'id' завдання у якості індекса таблиці
    tsk_months = tsk_months.reset_index().set_index('id')
    # Об'єднуємо завдання і параметри робочих місяців
    cols_selected = ['id', 'wp', 'cat', 'cache'] + users
    tsk_ext = pd.concat([tsk_months, tsk[cols_selected]], axis=1)
    # Створюємо довгу таблицю щомісячних фрагментів завдань
    id_vars = list(tsk_months.columns) + ['id', 'wp', 'cat', 'cache']
    tsk_long = pd.melt(tsk_ext, id_vars=id_vars, value_vars=users, 
                    var_name='user', value_name='contrib')
    # Виокремлюємо довгу таблицю щомісячних фрагментів завдань типу 'prop' (пропорційні) 
    tsk_prop = tsk_long.loc[tsk_long.cache.isnull() & ~tsk_long.contrib.isnull()] 
    tsk_prop = tsk_prop.drop('cache', axis=1)
    tsk_prop_cats = set(tsk_prop.cat.unique()) # Категорії завдань проміж персональних
    tsk_prop_types = set(tsk_type.loc[tsk_type.Type=='prop','Cat']) # Всі категорії проп. завдань
    if not tsk_prop_cats.issubset(tsk_prop_types):
        raise Exception(f'Деякі персональні завдання містять категорії {tsk_prop_cats} ' + \
                        f'поза визначених типів {tsk_prop_types}')
    
    tsk_cache = tsk_long.loc[~tsk_long.cache.isnull() & tsk_long.contrib.isnull()] 
    tsk_cache = tsk_cache.drop(['user', 'contrib'], axis=1)
    tsk_cache_cats = set(tsk_cache.cat.unique()) # Категорії завдань проміж типу 'кеш'
    tsk_cache_types = set(tsk_type.loc[tsk_type.Type=='cache','Cat']) # Всі категорії 'кеш' завдань
    if not tsk_cache_cats.issubset(tsk_cache_types):
        raise Exception(f'Деякі "кеш"-завдання містять категорії {tsk_cache_cats} ' + \
                        f'поза визначених типів {tsk_cache_types}')
    
    return tsk_prop, tsk_cache

def gen_pers_days_payment(team, tsk_prop, tot_salary, wppm):
    denna_stavka = [team.loc[user]['Ставка, Є/день'] for user in tsk_prop.user]
    tsk_prop['d_wages'] = denna_stavka
    wppm = pd.Series(list(wppm.pm), index=wppm.wp)
    '''
    pm_total = sum(wppm)
    tsk_prop['pm'] = [wppm[wp] for wp in tsk_prop.wp]
    
    twd_total = tsk_prop.twd.sum()
    w_days_in_month = 22
    for wp, wp_df in tsk_prop.groupby('wp'):
        twd_wp = wp_df.twd.sum()
        wp_df.twd / twd_wp * wppm[wp] * w_days_in_month
    '''
    
    # Обчислюємо нормувальний коефіцієнт
    norm_coef = tot_salary / (tsk_prop.twd * tsk_prop.contrib * tsk_prop.d_wages).sum()
    # Обчислюємо щомісячні персональні трудовитрати для кожного завдання
    pers_day_month = tsk_prop.twd * tsk_prop.contrib * norm_coef
    tsk_prop['pers_day'] = pers_day_month
    
    w_days_in_month = 20.7
    wppd = wppm * w_days_in_month
    pdm = pd.pivot_table(tsk_prop, values='pers_day',index='wp',aggfunc=sum).pers_day
    coef = wppd / pdm
    tsk_prop['pers_day'] = [row.pers_day * coef[row.wp] for _, row in tsk_prop.iterrows()]

    tsk_prop['pers_pay'] = pers_day_month * tsk_prop.d_wages
    
    pers_day = {pers: df.pers_day.sum() for pers, df in tsk_prop.groupby('user')}
    pers_day = pd.Series(pers_day)
    pers_day.name = 'Люд*дні'
    team = pd.concat([team, pers_day], axis=1)
    
    payment = {pers: df.pers_pay.sum() for pers, df in tsk_prop.groupby('user')}
    payment = pd.Series(payment)
    payment.name = 'Оплата, вал'
    team = pd.concat([team, payment], axis=1)
    
    team.rename(columns={'Ставка, Є/год': 'Ставка, вал/год', 
                         'Ставка, Є/день': 'Ставка, вал/день', 
                         'Ставка, Є/міс': 'Ставка, вал/міс'}, 
                inplace=True)
    
    cols_to_sum = ['Люд*дні', 'Оплата, вал']
    totals = pd.Series(team[cols_to_sum].sum(), index=cols_to_sum)
    totals.name = 'TOTALS'
    team = pd.concat([team, pd.DataFrame(totals).T])
    
    return team, tsk_prop

def join_two_long_tables(tsk_prop, tsk_cache, tsk_type):
    tsk_prop.drop(columns=['user', 'contrib', 'd_wages', 'pers_day'], inplace=True)
    tsk_prop.rename(columns={'pers_pay': 'cache'}, inplace=True)
    tsk_prop.cat = 'Personnel costs'

    tsk_type.set_index('Cat', inplace=True)
    tsk_cache.cat = [tsk_type.loc[cat].Item for cat in tsk_cache.cat]

    tsk_cache_upd = pd.DataFrame()
    for t, df in tsk_cache.groupby('id'):
        df.cache /= len(df) # розподіляємо суму кеш-завдання рівномірно за місяцями
        tsk_cache_upd = pd.concat([tsk_cache_upd, df])

    tsk_cache_upd = pd.concat([tsk_prop, tsk_cache_upd])
    
    return tsk_cache_upd

