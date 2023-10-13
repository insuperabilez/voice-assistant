﻿import config
import stt
import tts
from fuzzywuzzy import fuzz
import pandas as pd
from sms import send_message
import re
import numpy as np
from extractor import NumberExtractor
from number_to_text import num2text
from openpyxl import load_workbook
tts.play_sound('Голосовой ассистент готов.')
df = pd.read_excel('table.xlsx',decimal=',')
ints = ['Склад факт','Недогруз/Перегруз','План производства','Факт производства цеха','Сум. мес. потребность','Сум. мес. отгрузка']
df[ints]=df[ints].astype(float)
extractor = NumberExtractor()
def va_respond(voice: str):
    print(voice)
    if voice.startswith(config.VA_ALIAS):
        cmd = filter_cmd(voice)
        if (cmd['cmd'] in config.VA_CMDS):
            execute_cmd(cmd)
        else:
            print('не распознано')

def filter_cmd(raw_voice: str):
    cmd = raw_voice
    for x in config.VA_ALIAS:
        cmd = cmd.replace(x, "").strip()
    com = recognize_cmd(' '.join(cmd.split()[0:2]))
    cmd = ' '.join(cmd.split()[2:])
    id=''
    comment=''
    company = ''
    if com=='comment':
        pattern=r'для строки\s(.*?)\sтекст'
        matches=re.findall(pattern, cmd)
        for match in extractor(' '.join(matches)):
            id += str(match.fact.int)
        print(f'Найдена строка с номером {id}')
        pattern2 = r"текст (.*)"
        match = re.search(pattern2, cmd)
        if match:
            comment = match.group(1)
            print(comment)
    if com == 'sendmail':
        pattern = r'на номер\s(.*?)\sтекст'
        matches = re.findall(pattern, cmd)
        for match in extractor(' '.join(matches)):
            id += str(match.fact.int)
        print(f'распознан номер {id}')
        pattern2 = r"текст (.*)"
        match = re.search(pattern2, cmd)
        if match:
            comment = match.group(1)
            print(f'текст сообщения {comment}')
    for x in config.VA_TBR:
        cmd = cmd.replace(x, "").strip()
    company=''
    if com=='show1' or com=='show2':
        company=recognize_company(cmd)
    return {'cmd':com,'company':company,'id':id,'comment':comment}

def recognize_cmd(cmd: str):
    rc = {'cmd': '', 'percent': 0}
    for c, v in config.VA_CMDS.items():
        vrt = fuzz.ratio(cmd, v)
        if vrt > rc['percent']:
            rc['cmd'] = c
            rc['percent'] = vrt
    print(f'Команда {cmd}  распознана как {rc["cmd"]} с уверенностью {rc["percent"]} процентов.')
    if rc['percent']>50:
        return rc['cmd']
    else:
        return 'None'

def recognize_company(cmd: str):
    rc = {'item': '', 'percent': 0}
    for x in config.items:
        vrt = fuzz.ratio(cmd, x)
        if vrt > rc['percent']:
            rc['item'] = x
            rc['percent'] = vrt
    if rc['percent']>50:
        print(f'Распознана компания {rc["item"]} с уверенностью {rc["percent"]} процентов.')
        return rc['item']
    else:
        print('Компания не распознана')
        return ''

def execute_cmd(cmd):
    print('текущая команда:',cmd)
    if (cmd['cmd']=='show1' or cmd['cmd']=='show2') and (cmd['company']!=''):
        companykey = config.get_key_by_value(config.sootvetstvie, cmd['company'])
        filtered_df = df[(df['Заказчик'] == companykey) & (df['Недогруз/Перегруз'] < 0)]
        items = filtered_df['Синоним'].to_numpy()
        tts.play_sound(f'Для заказчика {cmd["company"]} найдено {num2text(len(items))} позиций с отклонениями')
        for item in items:
            n1=filtered_df[filtered_df['Синоним']==item]["Недогруз/Перегруз"].values[0]*-1
            n2=filtered_df[filtered_df['Синоним']==item]["Склад факт"].values[0]
            n3=filtered_df[filtered_df['Синоним']==item]["Сум. мес. потребность"].values[0]
            n4=filtered_df[filtered_df['Синоним']==item]["План производства"].values[0]
            n5=n2-n1-n3+n4
            #ei=filtered_df[filtered_df['Синоним']==item]['ЕИ'].values[0]
            ei='кэгэ'
            s1=f'долг за предыдущий период {num2text(n1)} {ei}'
            s2=f'на складе {config.convert_numbers_to_words(str(n2))} {ei}'
            s3=f'отгрузка текущего месяца {num2text(n3)} {ei}'
            s4=f'производственная программа {num2text(n4)} {ei}'
            if n5<0:
                s5=f'прогнозное отклонение на конец месяца {num2text(n5*-1)} {ei}'
            else:
                s5=f'прогнозное отклонение на конец месяца отсутствует'
            #tts.play_sound(item+s1+s2+s3+s4+s5)
            ssml = f"""
                    <speak>
                        <p>
                            {item}
                        </p>
                        <p>
                            {s1}
                        </p>
                        <p>
                            {s2}
                        </p>
                        <p>
                            {s3}
                        </p>
                        <p>
                            {s4}
                        </p>
                        <p>
                            {s5}
                        </p>
                    </speak>
                    """
            tts.play_ssml_sound(ssml)
    elif cmd['cmd']=='comment':
        id=int(cmd['id'])
        comment = cmd['comment']
        if str(cmd['id']).isnumeric() and cmd['comment']!='':
            id = int(cmd['id'])
            comment = cmd['comment']
            df.loc[df['ID']== id ,'Комментарий'] = comment
            print(comment)
            tts.play_sound('Комментарий добавлен')
            config.savetable('table.xlsx', 'table.xlsx', df)
        else:
            tts.play_sound('Не распознан идентификационный номер')
    elif cmd['cmd'] == 'sendmail':
        number=cmd['id']
        text=cmd['comment']
        if number.isnumeric():
            send_message(number,text)
            tts.play_sound('Сообщение отправлено')
stt.va_listen(va_respond)
#va_respond('алиса отправь сообщение на номер восемь девять шесть пять восемь четыре девять восемь четыре семь восемь текст это текст сообщения голосового помощника')
#va_respond('алиса отправь сообщение на номер три три семь текст это текст сообщения')
#va_respond('алиса добавь комментарий для строки три один восемь девять текст это второй комментарий')
#va_respond('алиса выполнение договоров для фирмы технологии')