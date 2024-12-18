''' Підготовка аргументів для побудови фінплану проєкту з план-графіка завдань.

Використання в коді fin_plan4.py:
    import sys
    from get_args import get_args
    
    data_file, result_file = get_args() # sys.argv за замовчанням

Використання з командного рядка: 
    python fin_plan4.py [-h] [-f FILE] [-d DIR] [-r RESULT]

Аргументи командного рядка:
  -h, --help            show this help message and exit
  -f FILE, --file FILE  Шлях до XLSX план-графіка завдань
  -d DIR, --dir DIR     Шлях до теки лише з одним XLSX план-графіком завдань
                        всередині
  -r RESULT, --result RESULT
                        Шлях до результуючого XLSX фінплану.

Обов'язково використовується лише один аргумент: --file або --dir. Аргумент
--result необов'язковий і може бути файлом або текою. У останньому випадку
використовується ім'я вхідного файлу із суфіксом _finplan.

'''
import os, glob, sys
import argparse
import pandas as pd

def _get_args(argv_list:list[str]) -> pd.Series:
    res = pd.Series(data='', index=['error','warning','file','result'], 
                    dtype=str, name='value')
    
    file_dir_usage = '''Обов'язково використовується лише один аргумент: 
                     --file або --dir. '''
    result_usage = '''Аргумент --result необов'язковий і може бути файлом або текою. У останньому 
                   випадку використовується ім'я вхідного файлу із суфіксом _finplan.'''
    
    parser = argparse.ArgumentParser(
        prog='get_args.py',
        description='Побудова фінплану проєкту з план-графіка завдань.',
        epilog=file_dir_usage + result_usage)
    parser.add_argument('-f', '--file', 
                        help='Шлях до XLSX план-графіка завдань')
    parser.add_argument('-d', '--dir', 
                        help='Шлях до теки лише з одним XLSX' +
                        ' план-графіком завдань всередині')
    parser.add_argument('-r', '--result', 
                        help='Шлях до результуючого XLSX фінплану.')
    parser.print_usage = parser.print_help
    
    args = parser.parse_args(argv_list[1:])
    
    # Якщо нема жодного аргументу
    if len(argv_list[1:]) == 0:
        parser.print_help()
        return res
    # Якщо в командному рядку одразу 2 аргументи --file або --dir
    if not (bool(args.file) ^ bool(args.dir)):
        res.error = file_dir_usage
        return res
    # Формуємо шлях file до XLSX план-графіка завдань
    if bool(args.file):
        if not os.path.isfile(args.file):
            res.error = f'Не знайдений вхідний файл: {args.file}'
            return res
        file = args.file
    else:
        if not os.path.isdir(args.dir):
            res.error = f'Не знайдена вхідна тека: {args.dir}'
            return res
        file_list = glob.glob(os.path.join(args.dir, '*.xlsx'))
        if len(file_list) != 1:
            res.error = f'Має бути лише 1 XLSX-файл у вхідній теці: {args.dir}'
            return res
        file = file_list[0]
    # Перевірка розширення шляху файлу 
    file_name, file_ext = os.path.splitext(file)
    if file_ext.lower() != '.xlsx':
        res.error = f'Файл має мати розширення ".XLSX": {file}'
        return res
    # Формуємо шлях до результуючого XLSX фінплану
    if not bool(args.result):
        result = file_name + '_finplan' + file_ext
    else:
        result_name, result_ext = os.path.splitext(args.result)
        # Якщо розширення відсутнє, result_name - тека
        if len(result_ext) == 0:
            if not os.path.isdir(result_name):
                res.error = f'Не знайдена вихідна тека: {result_name}'
                return res
            file_name, file_ext = os.path.splitext(os.path.split(file)[1])
            result = os.path.join(result_name, file_name + '_finplan' + file_ext)
        else:
            result = args.result
            result_dir = os.path.split(result)[0]
            if not os.path.isdir(result_dir):
                res.error = f'Не знайдена вихідна тека: {result_dir}'
                return res
            if os.path.isfile(result):
                res.warning = 'Файл з таким іменем існує, буде перезаписаний:' +\
                              ' {result}'
    res.file = file
    res.result = result
    return res

def get_args(args=sys.argv):
    res = _get_args(args)
    if all(res.str.len() == 0):
        sys.exit(0)
    elif res.error:
        sys.exit(res.error)
    elif res.warning:
        print(res.warning)
    return res.file, res.result

if __name__ == '__main__':
    args = sys.argv
    #args = ['', '-r', 'qwert/asdf.asd']
    
    data_file, result_file = get_args(args)
    print(f'Вхідний файл: {data_file}\nВихідний файл: {result_file}')
    
