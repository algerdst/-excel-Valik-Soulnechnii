import csv
import json
from excel2json import convert_from_file
import datetime
import glob
import os
import traceback


def convert_file():
    """
    конвертирует xlsx в json
    :return:
    """
    file = []
    path = os.getcwd()
    for filename in glob.glob(os.path.join(path, '*.xlsx')):
        file.append(filename)
    filename = file[0]
    print('Конвертация файла..')
    convert_from_file(filename)
    print('Конвертация завершена')


def convert_date(date):
    """
    конвертирует дату в нормальный формат
    :param date:
    :return:
    """
    date_number = date
    date = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=date_number)

    # Вывод в нужном формате
    formatted_date = date.strftime("%d.%m.%Y %H:%M:%S")
    return formatted_date


dict_sum1 = 0
dict_sum2 = 0
dict1 = {}
dict2 = {}


def make_dicts():
    """
    возвращает кортеж из 2 словарей, где ключи - это id,
    а значение - кортеж из даты и номера
    :return:
    """
    global dict_sum1
    global dict_sum2
    global dict1
    global dict2
    print('Поиск значений')
    file = []
    path = os.getcwd()
    for filename in glob.glob(os.path.join(path, '*.json')):
        file.append(filename)
    filename_1 = file[0]
    filename_2 = file[1]

    with open(filename_1, 'r', encoding='utf-8') as file:
        file_list1 = json.load(file)
    with open(filename_2, 'r', encoding='utf-8') as file:
        file_list2 = json.load(file)

    date_key = [i for i in file_list1[0]][0]

    for index in range(len([file_list1, file_list2])):
        file = [file_list1, file_list2][index]
        for i in file:
            try:
                i[date_key] = convert_date(i[date_key])
            except:
                pass
            try:
                i_d = int(i['id'])
            except:
                i_d = i['id']
            try:
                number = int(i['number'])
            except:
                number = i['number']
            if index == 0:
                dict1[i_d] = (i[date_key], number)
                try:
                    dict_sum1 += number
                except:
                    dict_sum1 += 0
            else:
                dict2[i_d] = (i[date_key], number)
                try:
                    dict_sum2 += number
                except:
                    dict_sum2 += 0


def find_matches_and_and_differences():
    global dict1
    global dict2

    ids = list(dict2)
    for _id in dict1:
        if _id in ids:
            with open('таблица совпадений.csv', 'a', newline='', encoding='utf-8-sig') as file:
                writer = csv.writer(file, delimiter=';')
                writer.writerow([_id, dict1[_id][0], dict1[_id][1]])
                print(f'Нашел совпадение по {_id}')
        else:
            with open('Значения которые есть в таблице 1, но нет в таблице 2.csv', 'a', newline='',
                      encoding='utf-8-sig') as file:
                writer = csv.writer(file, delimiter=';')
                writer.writerow([_id, dict1[_id][0], dict1[_id][1]])
                print(f'Нашел расхождение по {_id}')

    for _id in dict2:
        if _id not in dict1:
            with open('Значения которые есть в таблице 2, но нет в таблице 1.csv', 'a', newline='',
                      encoding='utf-8-sig') as file:
                writer = csv.writer(file, delimiter=';')
                writer.writerow([_id, dict2[_id][0], dict2[_id][1]])
                print(f'Нашел расхождение по {_id}')


def check_sum():
    """
    Проверяет суммы numbers в обоих словарях
    :return:
    """
    global dict_sum1
    global dict_sum2
    global dict1
    global dict2
    if dict_sum2 == dict_sum1:
        with open('суммы numbers в таблицах равны.txt', 'w', encoding='utf-8') as file:
            file.write('')
    else:
        print('cуммы не равны')
        with open('суммы numbers в таблицах не равны.txt', 'w', encoding='utf-8') as file:
            file.write(f'сумма в таблице 1 = {dict_sum1}\nсумма в таблице 2 = {dict_sum2}')
        try:
            with open('таблица совпадений.csv') as file:
                reader = csv.reader(file, delimiter=';')
                rows = list(reader)
        except:
            print('таблицы совпадений нет')
        for i in rows:
            _id = i[0]
            try:
                _id = int(_id)
            except:
                pass
            if _id in dict1 and _id in dict2:
                number1 = dict1[_id][1]
                number2 = dict2[_id][1]
                if number1 != number2:
                    with open('таблица несоответствий по numbers со строками из таблицы 1.csv', 'a', newline='',
                              encoding='utf-8') as file:
                        writer = csv.writer(file, delimiter=';')
                        writer.writerow([_id, dict1[_id][1], dict1[_id][0]])
                    with open('таблица несоответствий по numbers со строками из таблицы 2.csv', 'a', newline='',
                              encoding='utf-8') as file:
                        writer = csv.writer(file, delimiter=';')
                        writer.writerow([_id, dict2[_id][1], dict2[_id][0]])

        path = os.getcwd()
        flag = False
        for filename in glob.glob(os.path.join(path, '*.csv')):
            if 'таблица несоответствий по numbers со строками из таблицы 1.csv' in filename or 'таблица несоответствий по numbers со строками из таблицы 2.csv' in filename:
                flag = True
        if flag is False:
            with open('несоответсвий по numbers в таблицах не найдено.txt', 'w', encoding='utf-8') as file:
                file.write('')

def delete_json():
    path = os.getcwd()
    for file in glob.glob(os.path.join(path, '*.json')):
        os.remove(file)


convert_file()
make_dicts()
find_matches_and_and_differences()
check_sum()
delete_json()