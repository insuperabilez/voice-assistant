# voice-assistant

Чтобы установить необходимые пакеты - pip install -r requirements.txt <br />
Для запуска - python main.py <br />
При запуске ожидать фразы "Голосовой ассистент готов", после фразы можно обращаться по шаблонам:
"Алиса, озвучь план(факт)  по номенклатуре {номенклатура} за {месяц}" ,"Алиса, увеличь(уменьши/установи) план(факт)  по номенклатуре {номенклатура} на {число} единиц за {месяц}" <br />
Предметы находятся в таблице table.xlsx в столбце "номенклатура". <br />
Также необходимо скачать архив с моделью из Release и распаковать в папку model, находящуюся в директиве проекта
