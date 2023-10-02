import pandas as pd
import re
from number_to_text import num2text
import openpyxl
from openpyxl import load_workbook
import xlsxwriter
def get_key_by_value(dictionary, value):
    keys = list(dictionary.keys())
    if value in dictionary.values():
        index = list(dictionary.values()).index(value)
        return keys[index]
    return None  # Если значение не найдено
def convert_numbers_to_words(sentence):
    sentence = sentence.replace('.',' . ')
    sentence = sentence.replace('.', 'точка')
    words = sentence.split()  # разделение предложения на отдельные слова
    converted_words = []  # список для хранения преобразованных слов

    for word in words:
        if word.isdigit():  # проверка, является ли слово числом
            converted_word = num2text(int(word))  # преобразование числа в слово
            converted_words.append(converted_word)
        else:
            converted_words.append(word)  # оставляем слово без изменений

    converted_sentence = ' '.join(converted_words)  # объединение слов обратно в предложение
    return converted_sentence

df = pd.read_excel('table.xlsx')
months=['январь','февраль','март','апрель','май','июнь','июль','август','сентябрь','октябрь','ноябрь','декабрь']
items=[]

sootvetstvie={}
for item in pd.array(df['номенклатура']):
    converted = convert_numbers_to_words(item)
    items.append(converted)
    sootvetstvie[item] = converted
VA_ALIAS = ('алиса', 'алис', 'алисе', 'элиса')

VA_TBR = (
          'по','предмету','предмет ','предмету','за','по','на','номенклатуре','номенклатура')


VA_CMDS = {
    'showplan':('озвучь план'),
    'showfact':('озвучь факт'),
    'decplan':('уменьши план'),
    'incplan':('увеличь план'),
    'decfact': ('уменьши факт'),
    'incfact': ('увеличь факт'),
    'setplan':('установи план'),
    'setfact':('установи факт')
}

from openpyxl.utils.dataframe import dataframe_to_rows
def savetable(input,output,df):
    source_book = openpyxl.load_workbook(input)
    # Загрузка исходного листа
    source_sheet = source_book['Лист1']
    # Создание новой книги
    target_book = openpyxl.Workbook()
    # Создание нового листа
    target_sheet = target_book.active
    ws=target_book.active
    # Перенос стилей ячеек
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    for row in source_sheet.iter_rows():
        for cell in row:
            target_cell = target_sheet[cell.coordinate]
            target_cell.font = cell.font.copy()
            target_cell.border = cell.border.copy()
            target_cell.fill = cell.fill.copy()
            target_cell.number_format = cell.number_format
            target_cell.protection = cell.protection.copy()
            target_cell.alignment = cell.alignment.copy()

    # Сохранение изменений в целевой книге
    target_book.save(output)
