from datetime import datetime, date
from dateutil.rrule import rrule, MONTHLY
from dateutil.relativedelta import relativedelta
from calendar import monthrange

import pandas as pd


""" Init test dates """

def create_date(d):
    format_string = '%d.%m.%Y'
    return datetime.strptime(d, format_string)

def check_dates_consistency(task_dates, dps, dpe):
    def norm_date(d):
        return datetime(d.year, d.month, d.day)
    dps, dpe = norm_date(dps), norm_date(dpe)
    if dps > dpe:
        raise Exception(f'Project start date `{dps}` should be earlier project end date `{dpe}`')
    for i, r in task_dates.iterrows():
        ds, de = norm_date(r.ds), norm_date(r.de)
        if ds > de:
            raise Exception(f'Task {i} start date `{r.ds}` should be earlier task end date `{r.de}`')
        if ds < dps or ds > dpe:
            raise Exception(f'Task {i} start date `{r.ds}` is out the project time frame')
        if de < dps or de > dpe:
            raise Exception(f'Task {i} end date `{r.de}` is out the project time frame')
        
    return True

def days_in_month(d):
    # d - date format value
    return monthrange(d.year, d.month)[1]

def working_days_num(ds, de):
    return len(pd.bdate_range(ds, de, freq='B', inclusive='both'))

def working_days_in_month(r):
    # row from `drange`
    fd = date(r.y, r.m, 1)
    ld = date(r.y, r.m, r.ld)
    return working_days_num(fd, ld)
    
def create_proj_time_line(task_dates, dps, dpe):
    """ Make DataFrame index = month_index and columns = y, m """

    dclose = date(dpe.year, dpe.month, dpe.day) + relativedelta(months=1)

    drange = pd.date_range(dps, dclose, freq='MS') 
    drange = pd.Series(drange, name='prj_mon')
    drange = pd.DataFrame(drange)

    drange['y'] = [r.prj_mon.year for _, r in drange.iterrows()]
    drange['m'] = [r.prj_mon.month for _, r in drange.iterrows()]
    drange['ld'] = [days_in_month(r.prj_mon) for _, r in drange.iterrows()]
    drange['prjm'] = [i + 1 for i in drange.index]
    
    """ Find proj month index for arbitrary proj date """
    drange.set_index(['y','m'], inplace=True)

    task_dates['prjms'] = [drange.loc[(r.ds.year, r.ds.month),:].prjm \
                           for _, r in task_dates.iterrows()]
    task_dates['prjme'] = [drange.loc[(r.de.year, r.de.month),:].prjm \
                           for _, r in task_dates.iterrows()]

    """ Find working days in each period """
    task_dates.loc[:, 'wdn'] = [working_days_num(r.ds, r.de)\
                                for _, r in task_dates.iterrows()]
    return drange

def wdays_in_each_month_of_period(drange, task_dates):
    ''' Для кожного завдання типу prop обчислює серію з кількістю роб днів для 
        кожного проєктного місяця завдання і зберігає її у колонці wdays_m '''
    drange.reset_index(inplace=True)
    drange.set_index('prjm', inplace=True)

    wdays_m_lst = []
    for _, r in task_dates.iterrows():
        wdays_m = pd.Series(dtype=int, index=range(r.prjms, r.prjme + 1))
        if len(wdays_m) == 1: # prjms == prjme 
            wdays_m.loc[r.prjms] = working_days_num(r.ds, r.de)
        elif len(wdays_m) >= 2:
            dsld = date(r.ds.year, r.ds.month, days_in_month(r.ds)) # last date in month
            wdays_m.loc[r.prjms] = working_days_num(r.ds, dsld)
            defd = date(r.de.year, r.de.month, 1) # first date in month
            wdays_m.loc[r.prjme] = working_days_num(defd, r.de)
            for m in wdays_m.index[1:-1]:
                r = drange.loc[m]
                wdays_m.loc[m] = working_days_in_month(r)
        wdays_m_lst.append(wdays_m)

    task_dates.loc[:, 'wdays_m'] = wdays_m_lst

    return task_dates

def timestamp2date(ts):
    return datetime.fromtimestamp(datetime.timestamp(ts)).date()

def days_in_each_month_of_period(drange, task_dates):
    ''' Для кожного завдання типу cache обчислює серію з кількістю днів для 
        кожного проєктного місяця завдання і зберігає її у колонці days_m '''
    drange.reset_index(inplace=True)
    drange.set_index('prjm', inplace=True)

    days_m_lst = []
    for _, r in task_dates.iterrows():
        days_m = pd.Series(dtype=int, index=range(r.prjms, r.prjme + 1))
        if len(days_m) == 1 : # prjms == prjme 
            days_m.loc[r.prjms] = (r.de - r.ds).days
        elif len(days_m) >= 2:
            dsld = date(r.ds.year, r.ds.month, days_in_month(r.ds)) # last date in month
            days_m.loc[r.prjms] = (dsld - timestamp2date(r.ds)).days
            defd = date(r.de.year, r.de.month, 1) # first date in month
            days_m.loc[r.prjme] = (timestamp2date(r.de) - defd).days
            for m in days_m.index[1:-1]:
                r = drange.loc[m]
                days_m.loc[m] = drange.loc[m].ld
        days_m_lst.append(days_m)

    task_dates.loc[:, 'days_m'] = days_m_lst

    return task_dates
    
def working_time_plan(task_dates, dps, dpe):
    check_dates_consistency(task_dates, dps, dpe)
    drange = create_proj_time_line(task_dates, dps, dpe)
    task_dates = wdays_in_each_month_of_period(drange, task_dates)
    return task_dates, drange

def purchase_time_plan(task_dates, dps, dpe):
    check_dates_consistency(task_dates, dps, dpe)
    drange = create_proj_time_line(task_dates, dps, dpe)
    task_dates = days_in_each_month_of_period(drange, task_dates)
    return task_dates, drange

def set_proj_months(date_start, date_end):
    ''' # Таблиця проєктних місяців містить на кожний місяць значення
        {'y': years, 'm': months, 'q': quarters, 'prjm': prj_months, 
         'ld': last_day_in_month, 'wd': working_days_in_month}
    '''
    month_steps = [dt for dt in rrule(MONTHLY, dtstart=date_start, until=date_end)]
    years = [dt.year for dt in month_steps]
    months = [dt.month for dt in month_steps]
    quarters = [int((dt.month - 1)/3) + 1 for dt in month_steps]
    prj_months = [m + 1 for m in range(len(month_steps))] # number of project month
    last_days_in_month = [days_in_month(dt) for dt in month_steps]
    df = pd.DataFrame({'y': years, 'm': months, 'q': quarters, 
                       'prjm': prj_months, 'ld': last_days_in_month})
    df['wd'] = [working_days_in_month(r) for _, r in df.iterrows()]
    
    return df

def wdays_in_each_task_month(tsk, prj_months):
    ''' Обчислюємо кількість робочих днів 'twd' для кожного місяця, на який 
        припадає завдання (додатково до інших параметрів проєктного місяця 
        з `prj_months`) '''
    def get_prj_month_from_ym(prj_months, y, m):
        return prj_months.loc[(prj_months.y==y) & (prj_months.m==m)].index[0]    
    
    tsk_months = pd.DataFrame()
    for tsk_id, t in tsk.iterrows():
        prjm_s = get_prj_month_from_ym(prj_months, t.ds.year, t.ds.month)
        prjm_e = get_prj_month_from_ym(prj_months, t.de.year, t.de.month)
        # Місяці, на які припадає завдання
        t_months = prj_months.loc[range(prjm_s, prjm_e + 1)]
        tm_num = len(t_months) # Кількість місяців, на які припадає завдання
        wdays = []
        for i, tm in t_months.iterrows():
            day_s = t.ds.day if i == 1 else 1 # Перший день завдання у місяці
            day_e = t.de.day if i == tm_num else tm.ld # Останній день завдання у місяці
            try:
                wdays.append(working_days_num(date(tm.y, tm.m, day_s), 
                                              date(tm.y, tm.m, day_e)))
            except:
                wdays.append(working_days_num(date(tm.y, tm.m, day_s), 
                                              date(tm.y, tm.m, day_e-1)))
                #print(tm.y, tm.m, day_s, day_e)
        t_months['twd'] = wdays
        t_months['id'] = tsk_id
        tsk_months = pd.concat([tsk_months, t_months])
    
    return tsk_months


if __name__ == '__main__':

    """ Test data """
    
    dps, dpe = '01.01.2023', '31.12.2026'

    def create_test_dates():
        ds = ['31.01.2023', '02.06.2023', '01.12.2023', '01.06.2024', 
              '31.08.2024', '30.11.2026', '31.01.2023', '31.01.2023']
        de = ['02.07.2023', '02.07.2023', '01.01.2024', '01.07.2024', 
              '31.12.2026', '31.12.2026', '01.01.2024', '01.05.2023']

        ds = [create_date(d) for d in ds]
        de = [create_date(d) for d in de]

        task_dates = pd.DataFrame({'ds':ds, 'de': de})
    
        return task_dates

    dps = create_date(dps) # project start date
    dpe = create_date(dpe) # project end date
    task_dates = create_test_dates() # two columns: ds, de

    """ Start procedure """
    
    task_dates, drange = working_time_plan(task_dates, dps, dpe)

    for _, r in task_dates.iterrows():
        print(r.prjms, r.prjme, r.wdn, int(r.wdays_m.sum()))

    prjm = set_proj_months(dps, dpe)
    
