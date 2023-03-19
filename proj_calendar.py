''' Створюємо CSV-файл для імпорту початку і завершення 
    завдань проєкту до Гугл Календаря.
    Розміщення даних таблиці гарно описано у відео:
    https://www.youtube.com/watch?v=Yd1bQ3JDDLY
'''

import pandas as pd
from datetime import datetime

def make_calendar(tsk, users, prj_name):
    def append_rows(df, d_lbl):
        dct = {}
        for i, t in tsk.iterrows():
            dct['Subject'] = ', '.join([prj_name, t['Task/Deliv #'], d_lbl])
            dct['Start Date'] = datetime.strftime(t[d_lbl], '%d.%m.%Y')
            dct['Start Time'] = '09:00'
            dct['End Date'] = dct['Start Date']
            dct['End Time'] = '18:00'
            dct['All Day Event'] = 'FALSE'
            user_list = 'Виконавці: ' + ', '.join(t[users].dropna().index)
            dct['Description'] = '\n'.join([t['Description of task'], user_list])
            dct['Private'] = 'TRUE'
            df = pd.concat([df, pd.DataFrame(dct,index=[0])])
        return df

    cal_cols = ['Subject', 'Start Date', 'Start Time', 'End Date', 'End Time',
                'All Day Event', 'Description', 'Location', 'Private']

    df = pd.DataFrame(columns=cal_cols)

    d_lbl = 'DateStart'
    df = append_rows(df, d_lbl)

    d_lbl = 'DateEnd'
    df = append_rows(df, d_lbl)
    
    return df

