# Автор: Рабаданов Ильяc, геолог компании Petrotrace
# Версия: 1.1
# Дата: 10.07.2024
#
# Описание:
# Код выполняет проверку на наличие артефактов в файле Excel лист проверки относительно эталонного листа по критериям:
# Имени скважины, скважинным отбивкам, глубине по  MD
# C возможностью  выбора эталонного и проверочного листов и указанием интервала поиска от среднего значения глубины имеющихся в списке горизонтов на вхождение в интервал. 
# Выдает результаты проверки в терминале и дублирует в текстовый файл внутри директории в формате: <имя_файла>_result.txt
#
# ВАЖНО!
# Для работы скрипта требуется установить бибилиотеки  pandas, numpy, openpyxl
# Cкопируйте код ниже в терминал и нажмите Enter
#
# pip install pandas
# pip install numpy
# pip install openpyxl

import pandas as pd
import numpy as np
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Функция для выбора файла
def browse_file():
    root = Tk()
    root.withdraw()
    file_path = askopenfilename()
    return file_path

# Вызываем функцию выбора файла
file_path = browse_file()

# Получаем все листы из файла Excel
excel_file = pd.ExcelFile(file_path)
sheets = excel_file.sheet_names

# Выводим имена всех листов с номерами в терминале
print('Available sheets:')
for i, sheet in enumerate(sheets):
    print(f'{i+1}. {sheet}')

# Просим пользователя выбрать номер эталонного листа
print('Укажите имя эталонного листа:')
reference_sheet_number = int(input())
reference_sheet = sheets[reference_sheet_number - 1]

# Просим пользователя выбрать номер проверочного листа
print('Укажите имя проверочного листа:')
check_sheet_number = int(input())
check_sheet = sheets[check_sheet_number - 1]

# Загружаем данные из выбранных листов в DataFrame
df1 = pd.read_excel(file_path, sheet_name=reference_sheet)
df2 = pd.read_excel(file_path, sheet_name=check_sheet)

# Группируем данные в эталонной таблице по столбцам 'Well' и 'Surface'
grouped_ref = df1.groupby(['Well', 'Surface'])

# Создаем словарь для хранения средних значений глубин для каждой скважины и горизонта
mean_depths = {}
for (well_name, surface), group in grouped_ref:
    mean_depth = round(group['MD'].mean(), 2)
    mean_depths[(well_name, surface)] = mean_depth

# Запрашиваем у пользователя значение окна проверки
delta = float(input('Укажите значение окна проверки: '))

# Проверяем дубликаты отбивок у скважин
duplicates = df2.duplicated(subset=['Well', 'Surface', 'MD'], keep=False)
if duplicates.any():
    print('\nДубликаты отбивок у скважин:')
    for index, row in df2[duplicates].iterrows():
        print(f'Скважина {row["Well"]} в строке {index + 2} имеет дубликат отбивки {row["Surface"]} на глубине {row["MD"]}')

# Проверяем глубины для каждого горизонта, который вскрыт скважиной
for well_name, well_data in df2.groupby('Well'):
    if well_name not in df1['Well'].unique():
        print(f'\nСкважина {well_name}  в строке {group.index[0] + 2} отсутствует в эталонном списке')
        continue
    for surface, group in well_data.groupby('Surface'):
        if (well_name, surface) not in mean_depths:
            print(f'Скважина {well_name} в строке {group.index[0] + 2} имеет отбивку {surface}, которой нет в эталонном списке')
            continue
        mean_depth = mean_depths[(well_name, surface)]
        if not (group['MD'].min() >= mean_depth - delta and group['MD'].max() <= mean_depth + delta):
            print(f'Скважина {well_name} в строке {group.index[0] + 2} имеет глубину {group["MD"].iloc[0]:.2f}, которая не входит в интервал +-{delta}м от средней глубины {mean_depth:.2f}м для горизонта {surface}')

# Сохраняем результат проверки в виде txt файла с одинаковым именем как у открытого файла + текст (результат проверки)
file_name = os.path.splitext(file_path)[0]
with open(f'{file_name}_result.txt', 'w') as f:
    f.write('Результат проверки:\n')
    if duplicates.any():
        f.write('\nРеузльтаты проверки:')
        for index, row in df2[duplicates].iterrows():
            f.write(f'Скважина {row["Well"]} в строке {index + 2} имеет дубликат отбивки {row["Surface"]} на глубине {row["MD"]}')
        for well_name, well_data in df2.groupby('Well'):
            if well_name not in df1['Well'].unique():
                f.write(f'\nСкважина {well_name}  в строке {group.index[0] + 2} отсутствует в эталонном списке')
                continue
        for surface, group in well_data.groupby('Surface'):
            if (well_name, surface) not in mean_depths:
                f.write(f'Скважина {well_name} в строке {group.index[0] + 2} имеет отбивку {surface}, которой нет в эталонном списке')
                continue
        mean_depth = mean_depths[(well_name, surface)]
        if not (group['MD'].min() >= mean_depth - delta and group['MD'].max() <= mean_depth + delta):
            f.write(f'Скважина {well_name} в строке {group.index[0] + 2} имеет глубину {group["MD"].iloc[0]:.2f}, которая не входит в интервал +-{delta}м от средней глубины {mean_depth:.2f}м для горизонта {surface}')