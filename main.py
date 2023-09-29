import config
import stt
import tts
from fuzzywuzzy import fuzz
import pandas as pd
import re
from extractor import NumberExtractor
from number_to_text import num2text
tts.play_sound('Голосовой ассистент готов.')
df = pd.read_excel('table.xlsx')
extractor = NumberExtractor()
def va_respond(voice: str):
    print(voice)
    if voice.startswith(config.VA_ALIAS):
        cmd = filter_cmd(voice)
        if cmd['item']!='' and cmd['cmd']!='None':
            execute_cmd(cmd)
        else:
            print('не распознано')

def filter_cmd(raw_voice: str):
    cmd = raw_voice
    for x in config.VA_ALIAS:
        cmd = cmd.replace(x, "").strip()
    com = recognize_cmd(' '.join(cmd.split()[0:2]))
    cmd = ' '.join(cmd.split()[3:])
    month = recognize_month(cmd.split()[-1])
    cmd = cmd.replace(month,'').strip()
    res = re.findall(r'на\s(.*?)\sединиц', cmd)
    if res:
        res=res[-1]
    else:
        res=''
    cmd=cmd.replace(res,'').strip()
    if month=='':
        month='сентябрь'
    for x in config.VA_TBR:
        cmd = cmd.replace(x, "").strip()
    item=recognize_item(cmd)
    num = extractor.replace_groups(res)
    print(f'command: {com}, item: {item}, month: {month} , num: {num}')
    return {'cmd':com,'item':item,'month':month,'number':num}

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

def recognize_item(cmd: str):
    rc = {'item': '', 'percent': 0}
    for x in config.items:
        vrt = fuzz.ratio(cmd, x)
        if vrt > rc['percent']:
            rc['item'] = x
            rc['percent'] = vrt
    print(rc)
    if rc['percent']>50:
        print(f'Распознана номенклатура {rc["item"]} с уверенностью {rc["percent"]} процентов.')
        return rc['item']
    else:
        print('Номенклатура не распознана')
        return ''
def recognize_month(cmd: str):
    rc = {'month': '', 'percent': 0}
    for x in config.months:
        vrt = fuzz.ratio(cmd, x)
        if vrt > rc['percent']:
            rc['month'] = x
            rc['percent'] = vrt
    print(rc)
    if rc['percent']>50:
        return rc['month']
    else:
        return ''

def execute_cmd(cmd: str):
    itemkey = config.get_key_by_value(config.sootvetstvie, cmd['item'])
    m = cmd['month']
    i = cmd['item']
    try:
        n= int(cmd['number'])
    except:
        n=0
    tts.play_sound(f'Текущая команда: {config.VA_CMDS[cmd["cmd"]]}, по номенклатуре {i}, на {num2text(n)} единиц, месяц {m}')
    plan = df[df['номенклатура'] == itemkey][cmd['month'] + "1"].iloc[0]
    fact = df[df['номенклатура'] == itemkey][cmd['month'] + "2"].iloc[0]

    if cmd['cmd']=='showplan':
        tts.play_sound(f'По номенклатуре {i} в месяце {m} план равен {num2text(plan)}')
    if cmd['cmd'] == 'showfact':
        tts.play_sound(f'По номенклатуре {i} в месяце {m} по факту произведено {num2text(fact)}')
    if cmd['cmd'] == 'decplan':
        if pd.isna(plan):
            tts.playsound('Невозможно уменьшить несуществующее значение')
        else:
            plan=int(plan)-n
            df.loc[df['номенклатура'] == itemkey, f'{m}1'] = plan
            print('done')
    if cmd['cmd'] == 'incplan':
        if pd.isna(plan):
            tts.playsound('Невозможно увеличить несуществующее значение')
        else:
            plan=int(plan)+n
            df.loc[df['номенклатура'] == itemkey, f'{m}1'] = plan
            print('done')
    if cmd['cmd'] == 'decfact':
        if pd.isna(plan):
            tts.playsound('Невозможно уменьшить несуществующее значение')
        else:
            fact = int(fact) - n
            df.loc[df['номенклатура'] == itemkey, f'{m}1'] = fact
            print('done')
    if cmd['cmd'] == 'incfact':
        if pd.isna(plan):
            tts.playsound('Невозможно увеличить несуществующее значение')
        else:
            fact=int(fact)+n
            df.loc[df['номенклатура'] == itemkey, f'{m}1'] = fact
            print('done')
    if cmd['cmd'] == 'setplan':
        plan = n
        df.loc[df['номенклатура'] == itemkey, f'{m}1'] = plan
        print('done')
    if cmd['cmd'] == 'setfact':
        fact = n
        df.loc[df['номенклатура'] == itemkey, f'{m}1'] = fact
        print('done')
    df.to_excel('output.xlsx')




# начать прослушивание команд
stt.va_listen(va_respond)