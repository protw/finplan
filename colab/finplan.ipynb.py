# -*- coding: utf-8 -*-
"""finplan.ipynb

Automatically generated by Colaboratory.

Original file is located at
    https://colab.research.google.com/drive/1P_vsN_N5hXY1n_WKkSfgbxKW_9j7B_Pm

# Фінансовий план грантового проєкту

## Перед початком роботи:

1. ознайомтесь зі [стислою документацією застосунку](https://github.com/protw/finplan) та [короткою інструкцією з підготовки фінплану](https://drive.google.com/file/d/1v-u1tpZBFsTDaco3bXuTO1yN5YrfGzCP/view?pli=1);
2. упевніться, що маєте доступ до теки на Гугл Диску `230904 finplan/` (або, якщо маєте доступ до теки `UNITEAM`, зробіть те саме з нею);
3. зробіть ярлик теки `230904 finplan/` у свою теку `Мій диск` (або `My Drive`); якщо ви робили цей ярлик колись, то цю операцію можна пропустити.

## Порядок роботи

1. заповніть вашу форму план-графіка завдань (шаблон можна скачати [тут](https://github.com/protw/finplan/blob/master/data/Task%20schedule%20(template).xlsx). Ім'я файлу може бути довільним, розширення - `.xlsx`;
2. розмістіть заповнену форму у теці `230904 finplan/finplandata/`; **переконайтесь, що в теці це єдиний xlsx-файл**; тримайте теку [finplandata](https://drive.google.com/drive/folders/1XyFHw9Olp_NNoNR4ZTFqmklpkWG06RT_) відкритою в окремій вкладці бравзера;
3. через Гугл Диск у теці `_DOCS/FIN/230904 FINPLAN/` відкрийте (подвійним кліком) блокнот (_Jupyter Notebook_) з назвою `finplan.ipynb`; якщо ви користуєтесь _Google Colab_ вперше, то, можливо, вам прийдеться натиснути додаткову кнопку (по центру вгорі) типу "Відкрити з допомогою ..." і там обрати _Google Colab_;
4. запустіть цей блокнот через меню `Середовише виконання` / `Виконати всі` або з клавіатури - `CTRL+F9`;
5. у разі зазначення помилок виправте їх і запустіть знов (п. 4), до отримання результату;
6. результуючий файл матиме таке ж ім'я як і вхідний, але з префіксом `_finplan`, і розташований поруч із вхідним файлом;
7. ви можете змінити дані у вхідному файлі даних (див. п. 2) і перерахувати результат; **у такому разі перед повторним запуском видаліть результуючий файл з префіксом `_finplan`** (див. п. 6);
8. зробіть копію результуючого файлу змінивши префікс на `_finplan_formatted` і відформатуйте його (краще на локальному комп'ютері); не забудьте перенести отримані результати в роб. теку свого проєкту.
"""

#@title ## Конфігурація робочих директорій

gdrive = '/content/drive'
shdrive = '/MyDrive/' # MyDrive, Shareddrives
calc_dir = 'UNITEAM/_DOCS/FIN/230904 FINPLAN/' # @param ['230904 FINPLAN/', 'UNITEAM/_DOCS/FIN/230904 FINPLAN/']
calc_dir = gdrive + shdrive + calc_dir

# монтування Гугл Диску
from google.colab import drive
drive.mount(gdrive)

# надання доступу до фолдера коду 'code_dir'
code_dir = calc_dir + 'finplan'
import sys
sys.path.insert(0, code_dir)

!pip install XlsxWriter > /dev/null

#@title ## Обробка XLSX таблиці
from fin_plan4 import finplan_main

data_dir = calc_dir + 'finplandata'

finplan_main(data_dir)