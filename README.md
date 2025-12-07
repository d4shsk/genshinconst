# Тирлист созвездий
Вся информация взята из [таблицы Arataki Ozi](https://docs.google.com/spreadsheets/d/1q5Bk9AukcEhnnKMB5H0qayEnwB0A5wQSTCRTpT3OqJk/edit?gid=1444628575#gid=1444628575)

## Что это?
Это вот такой сайт который отображает, соритрует и фильтрует значения из таблицы Ози.  
Базовое отображение сайта  

<img width="600" height="936" alt="image" src="https://github.com/user-attachments/assets/9d4e2ed1-0ae8-42b3-b02d-d7069718f8e4" />
  
Фильтрация по силе C1 (сразу можно посмотреть)  

<img width="600" height="927" alt="image" src="https://github.com/user-attachments/assets/bb632ed9-e146-4dfd-be0b-225689d75692" />

Отображение конст, если выбран фильтр по силе консты, то эта конста выделяется (например как тут C1)

<img width="600" height="948" alt="image" src="https://github.com/user-attachments/assets/87defd9d-4a67-457d-b241-daceb6bea4e3" />

## Как запустить сайт?
1. Скачать все файлы с этого репозитория
2. В скачанной папке нажать на место с путем и ввести команду, после нажать Enter:
```python -m http.server```
3. Перейти по ссылке:
```http://localhost:8000/```
<img width="600" alt="image" src="https://github.com/user-attachments/assets/455b80ba-2956-43d7-a390-a33ca5c3abbf" />


## Как обновлять сайт?
1. Скачать все файлы с этого репозитория
2. Открыть таблицу, перейти на вкладку созвездия и скачать ее в .csv
<img width="600" alt="image" src="https://github.com/user-attachments/assets/c801ebf9-f128-40f4-b5e0-1e228cb57e02" />

3. Переименовать скачаный файл в data.csv
4. Положить ее в одну папку со скриптами и запустить parser.py
5. Появится файл result.json, поздравляю, данные обновлены!
