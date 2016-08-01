import imapy
import time
from html.parser import HTMLParser
import sqlite3
import os
import logging

# подключаемся к файлу с логами
logging.basicConfig(
    filename='log_downloadparser.txt',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

host = "imap.yandex.ru"     # хост email откуда будет парситься почта
user = "dagot321"           # логин на почте
password = "LJuME76A7T"     # пароль от почты
INBOX = "INBOX"             # название директории откуда будет парситься почта
OUTBOX = "outbox"           # название директории куда будет перемещаться почта после ее обработки
TIME = 30                   # количество секунд перед повторым запуском парсера

# класс для парсинга HTML из email
class EmailHTMLParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.links = []

    def handle_starttag(self, tag, attrs):
        try:
            if(tag == 'h2'):
                self.links.append('h2')
            if(tag == 'a'):
                for name, value in attrs:
                    if(name == 'href'):
                        self.links.append(value)
            if((tag == 'img')):
                for name, value in attrs:
                    if (name == 'alt'):
                        self.links.append(value)
        except:
            pass

    def handle_data(self, data):
        self.links.append(data)

    def report(self):
        out = []
        for i in range(len(self.links)):
            if(self.links[i] == "h2"):
                out.append(self.links[i+1:i+4])
        return out

# Функция для обработки формата даты
def formatDateTime(email_date):
    email_date = email_date.split()
    mounts = {"Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04",
              "May": "05", "Jun": "06", "Jul": "07", "Aug": "08",
              "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12"}
    return '{0}-{1}-{2} {3}:{4}:{5}'.format(email_date[3], mounts[email_date[2]], email_date[1],
                                            email_date[4][:2], email_date[4][3:5], email_date[4][6:])

# Проверяем подключение к базе данных и наличие рабочей таблицы
db = sqlite3.connect("email.db")
try:
    con = db.cursor()
    con.execute('''create table emails_table (name_report text, link text, images text, dateOfRec date)''')
    con.commit()
    logging.debug('Создана рабочая таблица в БД')
except Exception:
    pass
finally:
    db.close()
    logging.debug('Подключение к базе данных проверено')

# проверяем наличие папки images, если нет, то создаем
if not ("images" in os.listdir(path=".")):
    os.mkdir("images", mode=0o777)
    logging.debug('Создана папка images')

# запускаем парсер
while True:
    logging.debug('Парсер почты запущен')
    try:
        # подключение к email
        box = imapy.connect(
            host=host,
            username=user,
            password=password,
            ssl=True,
        )
    except Exception:
        logging.debug('Ошибка подключение к серверу почты')
        break

    # Открываем подключение к базе данных
    db = sqlite3.connect("email.db")
    con = db.cursor()

    # получаем кол-во писем доступных для парсинга
    status = box.folder(INBOX).info()
    total_messages = status['total']

    # получаем почту из папки
    emails = box.folder(INBOX).emails()

# обрабатываем каждое полученние письмо
    if emails:
        for email in emails:
            # получаем вложение картинки
            try:
                for attachment in email['attachments']:
                    file_name = attachment['filename']              # считываем имя вложенного файла
                    data = attachment['data']                       # загружаем картинку в виде потока битов
                    with open('images\\' + file_name, 'wb') as f:   # записываем полученный поток в файл, в папку images
                        f.write(data)
            except Exception:
                logging.debug("Ошибка при попытке получения изображения")

            # парсим HTML для получения информации о заголовке, ссылке
            parser = EmailHTMLParser()
            parser.feed(str(email['html']))
            reports = parser.report()

            # Добавляем информацию из письма в БД
            for report in reports:
                try:
                    con.execute('INSERT INTO emails_table VALUES (?,?,?,?)', (report[0], report[1], report[2],
                                                                              formatDateTime(email['date'])))
                    db.commit()
                except Exception:
                    logging.debug("Ошибка при попытке записи в БД. report = " + str(report))

            # перемещаем прочитанное сообщение в папку OUTBOX
            email.move(OUTBOX)


        # закрываем соединение с сервером почты
        box.logout()

    logging.debug('Получено ' + str(total_messages) + " писем")

    # закрываем соединение с базой данных
    db.close()

    # переход в режим ожидания на время = TIME сек.
    time.sleep(TIME)