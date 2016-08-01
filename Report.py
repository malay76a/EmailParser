from jinja2 import Template
import os
import sqlite3
from datetime import datetime
import logging

# подключаемся к файлу с логами
logging.basicConfig(
    filename='log_report.txt',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# проверяем наличие папки images, если нет, то создаем
if not("reports" in os.listdir(path=".")):
    os.mkdir("reports", mode=0o777)
    logging.debug('Создана папка reports')

try:
    db = sqlite3.connect("email.db")  # Открываем подключение к базе данных
    temp_data = db.execute('SELECT * FROM emails_table GROUP BY name_report'
                           ' having dateOfRec = max( dateOfRec )').fetchall()
except Exception:
    logging.debug('Ошибка при получении информации из БД')
    exit()
finally:
    db.close()

# получаем текущую дату
date = datetime.now()

# шаблон для создания HTML разметки
template = Template("""
                    <!DOCTYPE html>
                        <html lang="en">
                            <head>
                                <title>Отчет на {{ datetime }}</title>
                            </head>
                            <body>
                                {% for item in data %}
                                    <div>
                                        <h1><a href="{{ item[1] }}">{{ item[0] }}</a></h1>
                                        <h3>{{ item[3] }}</h3>
                                        <img src="../images/{{ item[2] }}">
                                    </div>
                                {% endfor %}
                            </body>
                        </html>
                    """)

# формируем название файла
report_name = "reports/" + str(date.date()) + ".html"

# записываем отчет в папку с отчетами
try:
    with open(report_name, "w") as fh:
        fh.write(template.render(data=temp_data, datetime=date.date()))
        logging.debug('Создан отчет ' + str(date.date()) + ".html")
except Exception:
    logging.debug('Ошибка при записи файла отчета. Файл: ' + str(date.date()) + ".html; Записываемая информация: " + str(temp_data))


