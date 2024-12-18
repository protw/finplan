from glob import glob
import os
import pandas as pd
import numpy as np
from datetime import datetime

from pivot_subtotals import pivot_w_subtot3

class GrossFinplan():
    def __init__(self, start_month, cur_sel='EUR', 
                 finplans_dir=r"../../240305 GROSSPLAN/Finplans", 
                 reports_dir=r"../../240305 GROSSPLAN/Reports",
                 prj_usr_cur_path=r'../../230708 GROSSBUCH/Загальне/Проєкти-Люди-Курс.xlsx'):
        print(f'GROSS FINPLAN\nПочаток: {start_month}, валюта розрахунку: {cur_sel}')
        self.start_month = start_month
        self.cur_sel = cur_sel
        self.finplans_dir = finplans_dir
        self.reports_dir = reports_dir

        # Формуємо ім'я результуючого звіту
        now = datetime.now()
        rep_suffix = datetime.strftime(now, '%y%m%d-%H%M%S')
        start_label = f'{start_month[0]-2000}{start_month[1]:02d}'
        report_file = f'gross_finplan-{start_label}-{rep_suffix}.xlsx'
        self.report_file = os.path.join(reports_dir, report_file)

        # Обрані стовпчики у вх таблицях (фінплани окремих проєктів)
        self.cols = ['prjm', 'y', 'm', 'q', 'cat', 'user', 'pers_day', 
                     'cache', 'prj_name', 'currency']
        self.cols_simple = ['y', 'm', 'cat', 'user', 'pers_day', 
                            'cache', 'prj_name', 'currency']

        print('== ЗАГАЛЬНА ІНФО ПРО ПРОЄКТИ, ЛЮДЕЙ, КУРС ВАЛЮТ ==')
        prj_usr_cur = pd.read_excel(prj_usr_cur_path, sheet_name=None)

        self.prj = prj_usr_cur['Проєкти']
        self.usr = prj_usr_cur['Люди'][['nick', 'pib']]
        self.cur = prj_usr_cur['Курс']
        
        return

    def check_currency(self):
        ''' Перевірка валют. Нормалізація курса валют відносно обраної 
        'cur_sel'. '''
        # Перевірка валют
        self.cur['UAH'] = 1
        cur_list = self.cur.columns[2:].to_list()
        assert self.cur_sel in cur_list, \
            f'Валюта {self.cur_sel} відсутня серед ''валют проєктів'
        # Нормалізація курса валют
        start_month_idx = self.start_month[0] * 100 + self.start_month[1]
        self.cur.index = self.cur.year * 100 + self.cur.month
        if self.cur.index.max() >= start_month_idx:
            cur_tbl = self.cur[cur_list].loc[self.cur.index >= start_month_idx, :]
        else:
            cur_tbl = self.cur[cur_list].loc[self.cur.index.max(), :].to_frame().T
        cur_tbl = cur_tbl.div(cur_tbl.loc[:, self.cur_sel], axis=0) # Нормалізація
        self.cur_list, self.cur_tbl = cur_list, cur_tbl
        return 
    
    def check_tbl(self, tbl, file):
        ''' Фільтрувати рядки в 'tbl', де місяць дорівнює або пізнше 'start_month'.
        Валідація назви проєкту. Валідація коротких імен учасників. Валідація 
        назви валюти проєкту. Вставляємо до таблиці 'tbl' курс валюти на кожний 
        місяць відносно 'cur_sel'. '''
        # Перевірити набір стовпчиків
        if set(self.cols).issubset(set(tbl.columns)):
            print(f'Повний фінплан {file}')
        elif set(self.cols_simple) == set(tbl.columns):
            tbl['prjm'] = None
            tbl['q'] = (tbl.m - 1) // 3
            print(f'Спрощений фінплан {file}')
        else:
            print(f'У таблиці "{file}" має бути набір стовчиків: '
                  f'{self.cols} або {self.cols_simple}')
        # Фільтрувати рядки в 'tbl', де місяць дорівнює або пізнше 'start_month'
        ym_idx = (self.start_month[0] * 100 + self.start_month[1]) > (tbl.y * 100 + tbl.m)
        tbl.drop(ym_idx[ym_idx].index, inplace=True)
        if len(tbl) == 0: 
            print(f'-- {os.path.split(file)[1]} -- пропущено, '
                  f'завершено раніше {self.start_month}')
            return False
        # Валідація назви проєкту
        prj_name = tbl.prj_name.unique()
        assert len(prj_name) == 1 and any(prj_name[0].upper() == self.prj.acronym), \
            f'У таблиці "{file}" є неіснуючий проєкт {prj_name}'
        # Валідація коротких імен учасників
        nicks_prj = set(tbl.user[~pd.isnull(tbl.user)].unique())
        nicks_org = set(self.usr.nick)
        assert nicks_prj.issubset(nicks_org), \
            f'У таблиці "{file}" є неіснуючі люди {nicks_prj-nicks_org}'
        # Валідація назви валюти проєкту
        cur_name = tbl.currency.unique()
        assert len(cur_name) == 1 and self.cur_list, \
            f'У таблиці "{file}" є неіснуюча валюта {cur_name}'
        # Вставляємо до таблиці 'tbl' курс валюти на кожний місяць відносно 'cur_sel'
        tbl.index = tbl.y * 100 + tbl.m
        tbl['cur_rate'] = None
        ym_common = list(set(tbl.index) & set(self.cur_tbl.index))
        if len(ym_common) >= 1:
            tbl.loc[ym_common, 'cur_rate'] = self.cur_tbl.loc[ym_common, self.cur_sel]
            tbl.loc[pd.isnull(tbl['cur_rate']), 'cur_rate'] = \
                self.cur_tbl.loc[max(ym_common), self.cur_sel]
        else:
            tbl['cur_rate'] = self.cur_tbl.loc[max(self.cur_tbl.index), self.cur_sel]
        
        return True
    
    def calc_dir_cost(self, gross_tbl):
        ''' розрахувати щомісячні загальні прямі витрати (крім накладних) і 
        планові накладні окремо'''
        grs_tbl = gross_tbl.copy()
        # обрані стовпчики
        grs_tbl_cols = ['prj_name', 'y', 'm', 'currency', 'cache']
        # окремо порахувати планові накладні
        grs_over = grs_tbl[grs_tbl_cols].loc[grs_tbl.cat == 'Overheads', :]
        # залишити лише обрані стовпчики і викреслити накладні
        grs_tbl = grs_tbl[grs_tbl_cols].loc[grs_tbl.cat != 'Overheads', :]
        # згрупувати по ['y', 'm', 'currency'] і просумувати по 'cache' з мінусом
        idx_cols = ['y', 'm', 'currency']
        grs_tbl_res = -grs_tbl.groupby(idx_cols).agg({'cache':['sum']})
        grs_over_res = grs_over.groupby(idx_cols).agg({'cache':['sum']})
        # викреслити нижній рівень мультиіндексу стовчиків (після groupby.agg)
        grs_tbl_res.columns = grs_tbl_res.columns.droplevel(1)
        grs_over_res.columns = grs_over_res.columns.droplevel(1)
        # приєднати стовпчики cache та overhead
        grs_tbl_res.rename(columns={'cache': 'dir_cost'}, inplace=True)
        grs_over_res.rename(columns={'cache': 'overhead'}, inplace=True)
        grs_tbl_res = pd.concat([grs_tbl_res, grs_over_res], axis=1)
        # повернути індекси до стовпчиків (після groupby.agg)
        grs_tbl_res.reset_index(inplace=True)
        # додати стовпчик з категорією 'прямі витрати'
        grs_tbl_res['cat'] = 'dir_cost'
        # створити об'єднане поле РРММ
        grs_tbl_res['ym'] = (grs_tbl_res.y - 2000) * 100 + grs_tbl_res.m
        
        return grs_tbl_res
    
    def main(self):
        self.check_currency() # gets 'cur_list', 'cur_tbl'
        
        print(f'== ДОДАЄМО ФІНПЛАНИ ПРОЄКТІВ, ПРИВОДИМО ПЛАТЕЖІ ДО {self.cur_sel} ==')
        finplans_list = glob(os.path.join(self.finplans_dir, '*_finplan.xlsx'))
        cols_ = self.cols + ['cur_rate',]
        gross_tbl = pd.DataFrame(columns=cols_)
        for i, file in enumerate(finplans_list):
            tbl = pd.read_excel(file, sheet_name='Long table')
            if self.check_tbl(tbl, file):
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
        
        print('-- щомісячні загальні прямі витрати (крім накладних)')
        grs_tbl_res = self.calc_dir_cost(gross_tbl)
        
        ''' Зберегти всі таблиці в окремих аркушах ексель-файлу '''
        
        with pd.ExcelWriter(self.report_file, engine='xlsxwriter') as writer:
            user_month_pay.to_excel(writer, sheet_name=f'Monthly payment, {self.cur_sel}')
            user_person_month.to_excel(writer, sheet_name='Person-months, pers-day')
            cat_month_pay.to_excel(writer, sheet_name=f'All monthly payments, {self.cur_sel}')
            grs_tbl_res.to_excel(writer, sheet_name='Long table monthly forecast', index=False)
        
        print(f'Всі дані збережено у файлі `{self.report_file}`!')
        print('== DONE! ==')
        
        return
    
if __name__ == '__main__':
    
    # Вхідні параметри процедури побудови Загального Фінплану
    input_pars = {
        'start_month':      (2024, 10),
        'cur_sel':          'EUR',
        'finplans_dir':     r'../../240305 GROSSPLAN/Finplans',
        'reports_dir':      r'../../240305 GROSSPLAN/Reports',
        'prj_usr_cur_path': r'../../230708 GROSSBUCH/Загальне/Проєкти-Люди-Курс.xlsx'
        }

    gf = GrossFinplan(**input_pars)   
    gf.main()

