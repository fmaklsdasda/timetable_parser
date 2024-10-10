# time parser 

Проект для парсинга расписания занятий из таблиц и отображения всех занятий для выбранного преподавателя. Инструмент извлекает и организует данные из расписания, чтобы упростить быстрый доступ к расписанию конкретного преподавателя.

## installation 

```bash
python -m venv ./env  
. ./env/Scripts/activate
pip install -r ./requirements.txt
```

## usage 

```bash
python ./gui.py  
```

## build 

```bash
pip install pyinstaller
pyinstaller --onefile gui.py
pyinstaller --onefile --icon=icon.ico gui.py

```