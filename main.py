import config
import stt
import tts
from fuzzywuzzy import fuzz
import time
import sounddevice as sd
import multiprocessing
import threading
import pandas as pd
from sms import send_message
import re
import numpy as np
from extractor import NumberExtractor
from number_to_text import num2text
from openpyxl import load_workbook
df = pd.read_excel('table.xlsx',decimal=',')
ints = ['Склад факт','Недогруз/Перегруз','План производства','Факт производства цеха','Сум. мес. потребность','Сум. мес. отгрузка']
df[ints]=df[ints].astype(float)
#df['Комментарий']=df['Комментарий'].astype('string')
extractor = NumberExtractor()
cfg=config.cfg
pause_between_rows = int(cfg.get('AUDIO', 'pause_between_rows'))
pause_between_items = int(cfg.get('AUDIO', 'pause_between_items'))
round_values = int(cfg.get('AUDIO', 'round_values'))
tts.play_sound('Голосовой ассистент готов')
class StoppableThread(threading.Thread):
    """Thread class with a stop() method. The thread itself has to check
    regularly for the stopped() condition."""

    def __init__(self,  *args, **kwargs):
        super(StoppableThread, self).__init__(*args, **kwargs)
        self._stop_event = threading.Event()

    def stop(self):
        self._stop_event.set()

    def stopped(self):
        return self._stop_event.is_set()
def thread_process(voice):

    cmd = filter_cmd(voice)
    if (cmd['cmd'] in config.VA_CMDS):
        execute_cmd(cmd)
    else:
        tts.play_sound('не распознано')
def va_respond(voice: str):
    print(voice)
    if voice.startswith(config.VA_ALIAS):
        tts.play_sound('Запрос принят')
        th = StoppableThread(target=thread_process,args=(voice,))
        th.start()
def filter_cmd(raw_voice: str):
    cmd = raw_voice
    for x in config.VA_ALIAS:
        cmd = cmd.replace(x, "").strip()
    com = recognize_cmd(' '.join(cmd.split()[0:2]))
    cmd = ' '.join(cmd.split()[2:])
    if com=='stop':
        tts.play_sound('Остановка потоков')
        all_threads = threading.enumerate()
        other_threads = [thread for thread in all_threads if thread != threading.main_thread() and thread != threading.current_thread()]
        for thread in other_threads:
            print('Остановлен поток:', thread)
            thread.stop()
            thread.join()
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
    if com == 'sendmail':
        pattern = r'\s(.*?)\sтекст'
        matches = re.findall(pattern, cmd)
        department=' '.join(matches)
        id=recognize_department(department)
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
        print(f'Распознан заказчик {rc["item"]} с уверенностью {rc["percent"]} процентов.')
        return rc['item']
    else:
        tts.play_sound('Заказчик не распознан')
        return ''

def recognize_department(cmd: str):
    rc = {'item': '', 'percent': 0}
    for v,x in config.departments.items():
        vrt = fuzz.ratio(cmd, x)
        if vrt > rc['percent']:
            rc['item'] = x
            rc['percent'] = vrt
    if rc['percent']>50:
        print(f'Распознан отдел {rc["item"]} с уверенностью {rc["percent"]} процентов.')
        return rc['item']
    else:
        tts.play_sound('Отдел не распознан')
        return ''
def execute_cmd(cmd):
    cur_tr = threading.current_thread()
    if cur_tr.stopped():
        return
    if (cmd['cmd']=='show1' or cmd['cmd']=='show2') and (cmd['company']==''):
        return
    if (cmd['cmd']=='show1' or cmd['cmd']=='show2') and (cmd['company']!=''):
        companykey = config.get_key_by_value(config.sootvetstvie, cmd['company'])
        filtered_df = df[(df['Заказчик'] == companykey) & (df['Недогруз/Перегруз'] < 0)]
        baditems = filtered_df['Синоним'].to_numpy()
        vectorized_convert = np.vectorize(config.convert_numbers_to_words)
        if len(baditems) >0 :
            items=vectorized_convert(baditems)
        else:
            items=[]
        tts.play_sound(f'Для заказчика {cmd["company"]} найдено {num2text(len(items))} позиций с отклонениями')
        for idx,item in enumerate(baditems):
            if cur_tr.stopped():
                return
            n1=filtered_df[filtered_df['Синоним']==item]["Недогруз/Перегруз"].values[0]*-1
            n2=filtered_df[filtered_df['Синоним']==item]["Склад факт"].values[0]
            n3=filtered_df[filtered_df['Синоним']==item]["Сум. мес. потребность"].values[0]
            n4=filtered_df[filtered_df['Синоним']==item]["План производства"].values[0]
            n5=n2-n1-n3+n4
            arr=[n1,n2,n3,n4,n5]
            if round_values==1:
                arr=[round(x) for x in arr]
            else :
                arr[4]=round(arr[4],1)
                for i in range(len(arr)):
                    if arr[i].is_integer():
                        arr[i] = int(arr[i])

            EI=filtered_df[filtered_df['Синоним']==item]['ЕИ'].values[0]
            ei=''
            match EI:
                case 'КГ':
                    ei='килограмм'
                case 'ТН':
                    ei = 'тээн'
                case 'М':
                    ei = 'эм'
                case 'ШТ':
                    ei='штук'
                case 'М':
                    ei='метр'
            s1=f'долг за предыдущий период {config.convert_numbers_to_words(str(arr[0]))} {ei}'
            s2=f'на складе {config.convert_numbers_to_words(str(arr[1]))} {ei}'
            s3=f'отгрузка текущего месяца {config.convert_numbers_to_words(str(arr[2]))} {ei}'
            s4=f'производственная программа {config.convert_numbers_to_words(str(arr[3]))} {ei}'
            if n5<0:
                s5=f'прогнозное отклонение на конец месяца {config.convert_numbers_to_words(str(arr[4]*-1))} {ei}'
            else:
                s5=f'прогнозное отклонение на конец месяца отсутствует'
            ssml = f"""
                    <speak>
                            {items[idx]}
                        <break time="{pause_between_rows}ms"/>
                            {s1}
                        <break time="{pause_between_rows}ms"/>
                            {s2}
                        <break time="{pause_between_rows}ms"/>
                            {s3}
                        <break time="{pause_between_rows}ms"/>
                            {s4}
                        <break time="{pause_between_rows}ms"/>
                            {s5}
                        
                    </speak>
                    """
            tts.play_ssml_sound(ssml)

    elif cmd['cmd']=='comment':
        if str(cmd['id']).isnumeric() and cmd['comment']!='':
            id = int(cmd['id'])
            comment = cmd['comment']
            df.loc[df['ID']== id ,'Комментарий'] = comment
            try:
                config.savetable('table.xlsx', 'table.xlsx', df)
                tts.play_sound('Комментарий добавлен')
            except:
                tts.play_sound('Доступ к файлу заблокирован')
            
        else:
            tts.play_sound('Не распознан идентификационный номер')
    elif cmd['cmd'] == 'sendmail':
        if cmd['id']=='' or cmd['comment']=='':
            return
        department = config.get_key_by_value(config.departments, cmd['id'])

        text=cmd['comment']
        options = cfg.options(department)
        sended=0
        for option in options:
            value = cfg.get(department, option)
            try:
                send_message(value,text)
                sended=sended+1
            except:
                print('Ошибка отправки сообщения')
        if sended!=len(options):
            tts.play_sound('Ошибка отправки сообщения')
        elif sended>0:
            tts.play_sound('Сообщения отправлены')
        elif sended==0 and len(options)==0:
            tts.play_sound('Адресат не указан')

def main():

    stt.va_listen(va_respond)
    #va_respond('алиса выполнение договоров для пао корпорация ависма')
    #va_respond('алиса стоп')
    #va_respond('алиса отправь сообщение в отдел сбыта текст проверка')
    #va_respond('алиса добавь комментарий для строки один два ноль текст комментарий')
    #va_respond('алиса выполнение договоров для корпорация ависмед')
    #va_respond('алиса добавь комментарий для строки один два ноль текст комментарий')

if __name__ == "__main__":
    main()


