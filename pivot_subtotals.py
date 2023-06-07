''' Pivot table with all subtotals by vertical and horizon.

    Inspired by the article: "Tabulating Subtotals Dynamically in Python Pandas 
    Pivot Tables" by Will Keefe, Jun 24, 2022, https://medium.com/p/6efadbb79be2
    
    Author: @oleghbond, https://protw.github.io/oleghbond
    Date: April 13, 2023
'''

import pandas as pd, numpy as np, os

def pivot_w_subtot(df, values, indices, columns, aggfunc=np.nansum, 
                   fill_value=np.nan, margins=False):
    '''
    Adds tabulated subtotals to pandas pivot tables with multiple hierarchical 
    indices.
    
    Args:
    - df - dataframe used in pivot table
    - values - values used to aggregrate
    - indices - ordered list of indices to aggregrate by
    - columns - columns to aggregrate by
    - aggfunc - function used to aggregrate (np.max, np.mean, np.sum, etc)
    - fill_value - value used to in place of empty cells
    
    Returns:
    -flat table with data aggregrated and tabulated
    
    '''
    listOfTable = []
    for indexNumber in range(len(indices)):
        n = indexNumber+1
        if n == 1:
            table = pd.pivot_table(df, values=values, index=indices[:n], 
                                   columns=columns, aggfunc=aggfunc, 
                                   fill_value=fill_value, margins=margins)
        else:
            table = pd.pivot_table(df, values=values, index=indices[:n], 
                                   columns=columns, aggfunc=aggfunc, 
                                   fill_value=fill_value)
        table = table.reset_index()
        for column in indices[n:]:
            table[column] = ''
        listOfTable.append(table)
    concatTable = pd.concat(listOfTable).sort_index()
    concatTable = concatTable.set_index(keys=indices)
    concatTable.sort_index(axis=0, ascending=True, inplace=True)

    return concatTable

def pivot_w_subtot2(df, values, indices, columns, aggfunc=np.nansum, 
                    fill_value=np.nan):
    ''' ЦЕЙ МЕТОД НЕ ДОВЕДЕНИЙ ДО ВИКОРИСТАННЯ!!
        Проблема методу pivot_w_subtot полягає в тім, що вона гарно будує 
        зведену таблицю лише по вертикалі. Окрім того не проводяться загальний 
        підсумок (найверхнього рівня) по всіх колонках і рядках. Додавання 
        аргумента margins (див. оригінальну статтю Will Keefe) не вирішує 
        повністю цієї проблеми. 
        
        Тому призначенням pivot_w_subtot2 є побудова, на базі pivot_w_subtot,
        зведеної таблиці для всіх рівнів по вертикалі і горизонталі, включно із
        загальним підсумком найверхнього рівня. 
    '''

    df_long = df.copy()
    df_long.fillna('-', inplace=True)
    
    ''' Створюємо два нові тимчасові (dummy) стовпчики для отримання загальних 
        підсумків по вертикалі і горизонталі лише на найверхньому рівні. '''
    df_long['All_cols'] = 0 # Для підсумків по всіх колонках
    df_long['All_rows'] = 0 # Для підсумків по всіх рядках
    
    ''' Спочатку створюємо зведення по вертикалі, тобто для всіх рядків, включно 
        із загальним підсумком найверхнього рівня вхідної таблиці df (за рахунок 
        включення стовпчика All_cols і відповідного індекса).  '''
    df_wide = pivot_w_subtot(df=df_long, values=values, indices=['All_rows',]+indices, 
                             columns=['All_cols',]+columns, aggfunc=aggfunc, 
                             fill_value=fill_value)

    ''' Розгортаємо "широку" таблицю df_wide знов у довгу df_long, але цього разу 
        вона містить підсумки всіх рівнів по вертикалі. '''
    df_long = pd.melt(df_wide, value_name='value', ignore_index=False).reset_index()

    ''' Тепер з довгої таблиці df_long створюємо зведення по горизонталі (тобто по 
        всіх стовпчиках), але розміщуючи їх по вертикалі. Це досягається 
        призначенням аргументу indices всіх колонок і, навпаки, - аргументу columns
        всіх рядків. І, наприкінці, результуючу таблицю транспонуємо (T), тобто
        міняємо рядки (axis=0) і колонки (axis=1) місцями. '''
    df_wide = pivot_w_subtot(df=df_long, values='value', columns=['All_rows',]+indices, 
                             indices=['All_cols',]+columns, aggfunc=aggfunc, 
                             fill_value=fill_value).T

    ''' Модифікуємо індекси по вертикалі і горизонталі, обрізваши у них
        тимчасовий найверхній рівень - All_cols і All_rows. '''
    df_wide.index = df_wide.index.droplevel(0)
    df_wide.columns = df_wide.columns.droplevel(0)
    
    return df_wide

if __name__ == '__main__':
    
    data_file = 'test_data/sampledatafoodsales.xlsx'

    df = pd.read_excel(data_file)
    
    val = 'TotalPrice'
    idx = ['Category', 'Product']
    cols = ['Region', 'City']

    df_rep0 = pd.pivot_table(df, values=val, index=idx, columns=cols, 
                             aggfunc=np.nansum)
    df_rep1 = pivot_w_subtot(df=df, values=val, indices=idx, columns=cols, 
                             aggfunc=np.nansum, margins=True)
    df_rep2 = pivot_w_subtot2(df=df, values=val, indices=idx, columns=cols, 
                             aggfunc=np.nansum)

    file_name, file_extension = os.path.splitext(data_file)
    result_file = file_name + '_pivtab.xlsx'

    with pd.ExcelWriter(result_file, engine='xlsxwriter') as writer:
        df_rep0.to_excel(writer, sheet_name="Standard Pivot")
        df_rep1.to_excel(writer, sheet_name="Will Keefe's Pivot")
        df_rep2.to_excel(writer, sheet_name="Olegh Bond's Pivot")
