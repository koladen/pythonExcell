from imaplib import IMAP4_SSL
import email
import os
import zipfile
import logging

YA_HOST = "mail.jino.ru"
YA_PORT = 993
YA_USER = "sverka@cottage-samara.ru"
YA_PASSWORD = "nedved1977"
SENDER = "accountopening@sberbank.ru"

PATH = r'\\10.37.2.3\Documents\01_Бухгалтерия\From Sberbank\ '
EXTRACT_PATH = r'\\10.37.2.3\Documents\01_Бухгалтерия\From Sberbank\Обработанные\ '
LOG_PATH = r'\\10.37.2.3\Documents\01_Бухгалтерия\From Sberbank\app.log'

logging.basicConfig(level=logging.INFO, filename=LOG_PATH, format='%(asctime)s - %(message)s')

def extract_attachment(file_name):
    with zipfile.ZipFile(file_name) as z_file:
        z_file.extractall(EXTRACT_PATH.strip())

    return filename[:-3] + 'xlsx'


files = list(filter(lambda x: x.endswith('.zip'), os.listdir(PATH.strip())))

connection = IMAP4_SSL(host=YA_HOST, port=YA_PORT)
connection.login(user=YA_USER, password=YA_PASSWORD)
status, msgs = connection.select('INBOX')
logging.info('Скрипт запущен ')
if status != 'OK':
    logging.info('Соединение не установлено! ' + YA_HOST + ' ' + YA_PORT + ' ' + YA_USER + ' ' + YA_PASSWORD)
else:
    logging.info('Соединение установлено!')
assert status == 'OK'

typ, data = connection.search(None, 'FROM', '"%s"' % SENDER)

for num in data[0].split():
    typ, message_data = connection.fetch(num, '(RFC822)')
    mail = email.message_from_bytes(message_data[0][1])

    if mail.is_multipart():
        for part in mail.walk():
            content_type = part.get_content_type()
            filename = part.get_filename()

            if filename:
                if filename not in files:
                    with open(PATH.strip() + filename, 'wb') as new_file:
                        new_file.write(part.get_payload(decode=True))
                    xls_file = extract_attachment(PATH.strip() + filename)
                    logging.info('Извлечен файл '+ xls_file)
                    print(xls_file)
                else:
                    print('Файл ' + filename + ' уже обрабатывался!')
                    logging.info('Файл ' + filename + ' уже обрабатывался!')
connection.close()
connection.logout()
