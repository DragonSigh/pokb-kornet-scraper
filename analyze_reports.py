import pandas as pd
import os
import json
import shutil

from datetime import date
from datetime import timedelta

# Настройки путей и дат
reports_path = os.path.join(os.path.abspath(os.getcwd()), 'reports')                # путь до общей папки с отчётами
credentials_path = os.path.join(os.path.abspath(os.getcwd()), 'auth-kornet.json')   # путь до файла со списком подразделений
#first_date = date.today() - timedelta(days=date.today().weekday())                  # начало недели
first_date = date.today() - timedelta(days=date.today().weekday()) - timedelta(days=7)
yesterday_date = date.today() - timedelta(days=1)                                   # вчера
today_date = date.today()                                                           # сегодня

# Если сегодня понеделеьник, то берем всю прошлую неделю
if date.today() == (date.today() - timedelta(days=date.today().weekday())):
    first_date = date.today() - timedelta(days=date.today().weekday()) - timedelta(days=7) # начало прошлой недели

# Функция для нормализации ФИО из ЕМИАС
def complex_rename(x):
    if isinstance(x, str):
        first_name = x.split(' ')[1]
        second_name = x.split(' ')[2]
        last_name = x.split(' ')[3].replace(',', '')
        return f'{first_name} {second_name} {last_name}'.upper()
    else:
        return 0

# Функция для сохранения датафрейма в Excel с автоподбором ширины столбца
def save_to_excel(dframe: pd.DataFrame, path, index_arg=False):
    with pd.ExcelWriter(path, mode='w', engine='openpyxl') as writer:
         dframe.to_excel(writer, index=index_arg)
         for column in dframe:
            column_width = max(dframe[column].astype(str).map(len).max(), len(column))
            col_idx = dframe.columns.get_loc(column)
            writer.sheets['Sheet1'].column_dimensions[chr(65+col_idx)].width = column_width + 5

# Очистить папку с отчётами и пересоздать если её нет в системе

shutil.rmtree(reports_path + '\\result\\', ignore_errors=True)

try:
    os.mkdir(reports_path + '\\result\\') 
except FileExistsError:
    pass   

# Соединение датафреймов из ЕМИАСа в один
df_list = []
with os.scandir(reports_path + '\\from_emias') as it:
    for entry in it:
        if entry.is_file():
            df_temp = pd.read_excel(entry.path, usecols = 'A, C, H, K, O', skiprows=range(1, 3), skipfooter=4, header=0)
            df_list.append(df_temp)
df_emias = pd.concat(df_list)
df_emias.columns = ['Подразделение', 'Кабинет', 'ФИО пациента', 'Время приема по записи', 'Отметка о приеме']
df_emias['ФИО пациента'] = df_emias['ФИО пациента'].apply(complex_rename)

# Очистка датафрейма ЕМИАСа
values_to_remove = ['Максимум','Итог','Среднее','Количество'] 
pattern = '|'.join(values_to_remove)
df_emias = df_emias.loc[~df_emias['Подразделение'].str.contains(pattern, case=False)]

# Сконвертировать время приема по записи в дату
df_emias['Время приема по записи'] = pd.to_datetime(df_emias['Время приема по записи'])
# Выделить столбец только с датой
df_emias['Дата записи'] = pd.to_datetime(df_emias['Время приема по записи'], dayfirst=True).dt.date
# Сгруппировать по ФИО и дате записи, если есть дубли
df_emias = df_emias.groupby(['ФИО пациента', 'Дата записи']).head(1).reset_index(drop=True)

# Выделить в отдельный датафрейм все фактические неявки за прошедшие дни
df_noshow = df_emias[(df_emias['Отметка о приеме'] == 'Неявка') & (df_emias['Время приема по записи'] < pd.to_datetime('today').normalize())]
# Убрать все фактические неявки из основного датафрейма ЕМИАС
df_emias = df_emias[~((df_emias['Отметка о приеме'] == 'Неявка') & (df_emias['Время приема по записи'] < pd.to_datetime('today').normalize()))]
# Убрать пустые значения, отмены и переносы
df_emias = df_emias[(~df_emias['Отметка о приеме'].isnull()) &
                    (df_emias['Отметка о приеме'] != 'Запись отменена') &
                    (df_emias['Отметка о приеме'] != 'Запись перенесена')]
# Пометить все неявки в будущем как запланированные посещения
df_emias.loc[df_emias['Отметка о приеме'] == 'Неявка', 'Отметка о приеме'] = 'Запланирован'

# Соединение датафреймов из Корнета в один и поиск ФИО, которых нет в записи в кабинет ЕМИАС
f = open(credentials_path, 'r', encoding='utf-8')
data = json.load(f)
f.close()
df_list = []
for _departments in data['departments']:
    df_kornet = pd.read_excel(reports_path + '\\from_kornet\\'+ _departments['department'] + '.xlsx', header=0 , usecols = "A, C, E, F")
    df_kornet['ФИО пациента'] = df_kornet['ФИО пациента'].str.upper()
    df_kornet = df_kornet.groupby(['ФИО пациента', 'Дата выписки']).head(1).reset_index(drop=True)
    df_kornet['Дата выписки'] = pd.to_datetime(df_kornet['Дата выписки'], dayfirst=True).dt.date
    df_list.append(df_kornet)
    result = df_kornet[~df_kornet['ФИО пациента'].isin(df_emias['ФИО пациента'])]
    result = result[result['Дата выписки'] >= today_date]
    save_to_excel(result, reports_path + '\\result\\' + _departments['department'] +' - нет записи в кабинет выписки рецептов на ' + str(today_date) + '.xlsx')

df_kornet = pd.concat(df_list)

# Поиск людей, которым не отмечена явка в ЕМИАС, но выдан рецепт в Корнет
# Соединение датафреймов неявок и корнета
df_noshow = df_noshow \
    .merge(df_kornet,  
           left_on=['ФИО пациента', 'Дата записи'],  
           right_on=['ФИО пациента', 'Дата выписки'], 
           how='inner') \
    .drop(['Отметка о приеме', 'Отделение', 'СНИЛС', 'Дата записи', 'Дата выписки'], axis=1)

# Наглядное выражение времени вне расписания 
df_noshow['Время приема по записи'] = df_noshow['Время приема по записи'].apply(lambda x: x.strftime('%Y-%m-%d %H:%M').replace('00:00', 'вне расписания'))
# Сохранить отчет по непроставленным явкам
save_to_excel(df_noshow, reports_path + '\\result\\' + '_Не проставлена явка о приеме, но выписан рецепт.xlsx')
# Сегодняшний день исключаем, так как рецепт ещё могут выписать позже
df_kornet = df_kornet[df_kornet['Дата выписки'] < today_date]
# ЕСЛИ НУЖНО сохранить объединенные отчёты для обеих систем для дебага
#save_to_excel(df_kornet, reports_path + '\\result\\' +'КОРНЕТ.xlsx', index_arg=True)
#save_to_excel(df_emias, reports_path + '\\result\\' +'ЕМИАС.xlsx', index_arg=True)

# Сводим статистику по выписанными не по регламенту рецептам в единую таблицу
df_kornet = df_kornet \
    .merge(df_emias,  
           left_on=['ФИО пациента', 'Дата выписки'],  
           right_on=['ФИО пациента', 'Дата записи'], 
           how='left') \
    .assign(reglament = lambda x: ~x['Время приема по записи'].isna()) \
    .groupby('Отделение') \
    .agg({'reglament': ['count', 'sum']}) \
    .assign(rate_correct = lambda x: round(100 * x['reglament']['sum'] / x['reglament']['count']))
# Сохраняем свод
save_to_excel(df_kornet, reports_path + '\\result\\' +'_Свод по выписанным рецептам не по регламенту ' + str(first_date) + '_' + str(yesterday_date) + '.xlsx', index_arg=True)