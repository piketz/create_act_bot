import telebot
from telebot import types, apihelper
import jinja2
import pdfkit
from datetime import datetime
import tempfile
from PyPDF2 import PdfMerger
import os
from dateutil import parser
import pandas as pd
import logging
import platform
from dotenv import load_dotenv
import json
import requests
import base64

logging.basicConfig(level=logging.ERROR, filename="py_log.log", filemode="w")

# Инициализация бота
dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)

apihelper.ENABLE_MIDDLEWARE = True
apihelper.SESSION_TIME_TO_LIVE = 5 * 60
admin_chat_id = os.environ.get('ADMIN_CHAT_ID')  # Загрузите ID чата администратора из переменной окружения

if platform.system() == 'Linux':
    wkhtmltopdf = '/bin/wkhtmltopdf'
    bot = telebot.TeleBot(os.environ.get('API_KEY'))
else:
    wkhtmltopdf = './wkhtmltox/bin/wkhtmltopdf.exe'
    bot = telebot.TeleBot(os.environ.get('API_KEY_dev'))

downloads_folder = 'downloads'
generated_folder = 'files'
max_folder_size = 1000 * 1024 * 1024  # 1000 МБ в байтах

def check_file_content(exelfile):
    try:
        # Загрузка данных из файла Excel
        df = pd.read_excel(exelfile)

        # Проверка наличия конкретных столбцов в файле
        required_columns = ['Дата создания', 'Конфигурационная единица', 'Объект обслуживания']

        if not set(required_columns).issubset(set(df.columns)):
            logging.error(f"File does not contain all required columns: {required_columns}")
            return False

        logging.info("File content is valid.")
        return True
    except Exception as e:
        logging.error(f"Error checking file content: {e}")
        return False


def generate_documents(exelfile, operation, fio_ispolnitel, day, month, year):
    """
    Генерирует PDF-документы на основе данных из файла Excel с учетом ограничения в 7 строк данных для каждого index_ops.

    Args:
        exelfile (str): Путь к файлу Excel с данными.
        operation (str): Операция или услуга, представленная в данных.
        fio_ispolnitel (str): ФИО исполнителя операции или услуги.
        day (str): День для заполнения PDF-файлов (в формате 'dd').
        month (str): Месяц для заполнения PDF-файлов (в формате 'mm').
        year (str): Год для заполнения PDF-файлов (в формате 'yyyy').

    Returns:
        list or None: Список путей к созданным PDF-файлам или None в случае ошибки.
    """
    logging.info(f"generate_documents exelfile={exelfile}, operation={operation}, fio_ispolnitel={fio_ispolnitel}, day={day}, month={month} , year={year}")

    # Загрузка данных из файла Excel
    df = pd.read_excel(exelfile)

    # Проверка на пустые значения дня, месяца и года
    if not day:
        day = '__'
    if not month:
        month = '__'
    if not year:
        year = '____'

    # Логирование информации о вызове функции
    logging.info(f'generate_documents '
                 f'exelfile={exelfile}, '
                 f'operation={operation}, '
                 f'fio_ispolnitel={fio_ispolnitel}, '
                 f'day={day}')

    # Инициализация списка для хранения путей к созданным PDF-файлам
    generated_docs = []

    # Словарь для хранения строк данных для каждого index_ops
    data_groups = {}

    # Обработка каждой строки данных из файла Excel
    for index, row in df.iterrows():

        try:
            # Получение данных из текущей строки
            index_ops = str(row['Объект обслуживания']).split()[0]
            config_data = str(row['Конфигурационная единица'])
            config_parts = [part.strip() for part in config_data.split('|')]
            name_file = f"{index_ops}_{str(row['NumberIn'])}_{str(row['Number'])}_{str(row['Задание']).replace('/', '-')}"
            date_str = str(row['Дата создания'])
            date_obj = parser.parse(date_str)
            logging.debug(f"[generate_documents] Обработка строки {index_ops}: {row}")

            # Проверка длины config_parts перед использованием
            if len(config_parts) >= 4:
                model_ke = f'{config_parts[1]} {config_parts[2]} {config_parts[3]}'
            else:
                logging.error(f"Недостаточно элементов в config_parts: {config_parts}")
                model_ke = "        "

            # Проверка значений num_rp и num_im
            #num_rp = str(row['NumberIn']) if pd.notna(row['NumberIn']) else "        "
            num_in_val = row.get('NumberIn')
            incoming_val = row.get('incomingNumber')
            if pd.notna(num_in_val) and str(num_in_val).strip():
                num_rp = str(num_in_val)
            elif pd.notna(incoming_val) and str(incoming_val).strip():
                num_rp = str(incoming_val)
            else:
                num_rp = "        "
            num_im = str(row['Number']) if pd.notna(row['Number']) else "        "

            # Добавление строки данных в соответствующую группу по index_ops
            if index_ops not in data_groups:
                data_groups[index_ops] = []

            data_groups[index_ops].append({
                'name_file': name_file,
                'nn': str(row['Задание']),
                'fio_ispolnitel': fio_ispolnitel,
                'day': day,
                'month': month,
                'year': year,
                'day_crt': date_obj.day,
                'month_crt': date_obj.month,
                'year_crt': date_obj.year,
                'index_adress': str(row['Объект обслуживания']),
                'model_ke': model_ke,
                'num_ke': config_parts[0],
                'work': operation,
                'num_rp': num_rp, #str(row['NumberIn']),
                'num_im': num_im #str(row['Number'])
            })
        except Exception as e:
            logging.exception(f"Error processing row {index_ops}: {row} — {e}")
            return None  # В случае ошибки возвращается None

    # Обработка данных для каждого index_ops
    for index_ops, data_group in data_groups.items():
        #print(f'data_group = {data_group}')
        try:
            # Инициализация списка для хранения путей к созданным PDF-файлам для текущего index_ops
            generated_docs_index_ops = []

            # Разбиение данных на группы по 7 строк для каждого index_ops
            for i in range(0, len(data_group), 7):
                data_chunk = data_group[i:i+7]

                # Создание контекста для текущей группы данных
                context = {}
                for j, data_row in enumerate(data_chunk, start=1):
                    for key, value in data_row.items():
                        if key == 'model_ke':
                            context[f'model_ke_{j}'] = value
                            context[f'num_ke_{j}'] = data_row['num_ke']  # Присваиваем значение num_ke каждой model_ke
                            context[f'num_rp_{j}'] = data_row['num_rp']
                            context[f'num_im_{j}'] = data_row['num_im']
                        else:
                            context[key] = value
                # Вместо отдельных model_ke_1, model_ke_2 ... создайте список словарей
                devices = []
                for j, data_row in enumerate(data_chunk, start=1):
                    devices.append({
                        'model_ke': safe_str(data_row.get('model_ke')),
                        'num_ke': safe_str(data_row.get('num_ke')),
                        'serial_number': safe_str(data_row.get('serial_number')),
                        'num_rp': safe_str(data_row.get('num_rp')),
                        'num_im': safe_str(data_row.get('num_im')),
                    })
                context['devices'] = devices

                # Создание PDF-файла на основе текущей группы данных
                logging.info(f"[generate_documents] Генерация PDF для index_ops={index_ops}, срез {i}–{i + 7}")
                file_path = one_pdf_crt(context)
                logging.info(f"[generate_documents] Результат one_pdf_crt: {file_path}")
                if file_path:
                    generated_docs_index_ops.append(file_path)

            # Добавление путей к созданным PDF-файлам в общий список
            generated_docs.extend(generated_docs_index_ops)
        except Exception as e:
            print(f"Error processing data group for index_ops={index_ops}: {e}")
            logging.error(f"Error processing data group for index_ops={index_ops}: {e}")
            return None  # В случае ошибки возвращается None

    # Логирование информации о созданных PDF-файлах
    logging.info(f'return generated_docs = {generated_docs}')

    # Возвращение списка путей к созданным PDF-файлам
    return generated_docs

def safe_str(value):
    return '' if pd.isna(value) or value in [None, 'nan', 'NaN'] else str(value)

def one_pdf_crt(context):
    logging.info(f"[one_pdf_crt] Начало генерации PDF для файла: {context.get('name_file', 'unknown')}")
    try:
        url = 'http://gooduser:secretpassword@192.168.1.229:5552/'
        logging.info(f"[one_pdf_crt] URL для генерации PDF: {url}")

        template_loader = jinja2.FileSystemLoader('./')
        template_env = jinja2.Environment(loader=template_loader)
        template = template_env.get_template('template.html')
        context = {k: v for k, v in context.items() if v is not None}
        output_text = template.render(context)

        encoded_content = base64.b64encode(output_text.encode('utf-8')).decode('utf-8')

        data = {
            'contents': encoded_content,
            'options': {
                'enable-local-file-access': '',
                'margin-top': '6',
                'margin-right': '6',
                'margin-bottom': '6',
                'margin-left': '6',
                'page-size': 'A3',
            }
        }

        logging.debug(f"[one_pdf_crt] Отправляем запрос на PDF-сервер...")
        response = requests.post(url, data=json.dumps(data), headers={'Content-Type': 'application/json'})
        logging.debug(f"[one_pdf_crt] Ответ от сервера: код {response.status_code}")

        if response.status_code == 200:
            pdf_path = os.path.join(tempfile.mkdtemp(), f"{context['name_file']}.pdf")
            with open(pdf_path, 'wb') as f:
                f.write(response.content)
            logging.info(f"[one_pdf_crt] PDF создан: {pdf_path}")
            return pdf_path
        else:
            logging.error(f"[one_pdf_crt] Ошибка при генерации PDF. Код: {response.status_code}, Тело: {response.text}")
            return None
    except Exception as e:
        logging.exception(f"[one_pdf_crt] Исключение при генерации PDF: {e}")
        return None


# Обработчик команды /start
@bot.message_handler(commands=['start'])
def start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    item = types.KeyboardButton("Start")
    markup.add(item)
    bot.send_message(message.chat.id, "Привет! Я бот для загрузки Excel-файлов. Отправь мне файл.. ")
    admin_message = f"Start от пользователя {message.from_user.username}."
    bot.send_message(admin_chat_id, admin_message)


# Обработчик текстовых сообщений
@bot.message_handler(content_types=['text'])
def handle_text(message):
    bot.send_message(message.chat.id, "Пожалуйста, отправьте мне Excel-файл, выгруженный из remo.itsm365.com "
                                      "c столбцами: Задание,	Описание (RTF),	Адрес,	Объект обслуживания,"
                                      "	Статус,	Крайний срок решения,	Дата создания,	Number,	NumberIn, "
                                      "incomingNumber, Конфигурационная единица для ТО надо ещё incomingNumber", reply_markup=types.ReplyKeyboardRemove())


# Обработчик загрузки документов (Excel-файлов)
@bot.message_handler(content_types=['document'])
def handle_document(message):
    try:
        if message.document.mime_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
            file_in = bot.get_file(message.document.file_id)

            if get_folder_size(downloads_folder) > max_folder_size:
                remove_file(downloads_folder)
            if get_folder_size(generated_folder) > max_folder_size:
                remove_file(generated_folder)

            file_path = os.path.join(downloads_folder,
                                     datetime.now().strftime('%Y%m%d%H%M%S') + message.document.file_name)
            downloaded_file = bot.download_file(file_in.file_path)

            if not os.path.exists(downloads_folder):
                os.makedirs(downloads_folder)


            if file_in.file_size > 20 * 1024 * 1024:
                bot.send_message(message.chat.id,
                                 "Размер файла превышает 20MB. Пожалуйста, загрузите файл размером не более 20MB.")
                return
            else:
                with open(file_path, 'wb') as new_file:
                    new_file.write(downloaded_file)

                if not check_file_content(file_path):
                    bot.send_message(message.chat.id, "Excel-файл имеет неверное содержание. "
                                                      "Он должен содержать столбцы Задание,	"
                                                      "Описание (RTF),	Адрес,	Объект обслуживания, "
                                                      "Статус, Крайний срок решения, "
                                                      "Дата создания, Number,	NumberIn, "
                                                      "Конфигурационная единица. "
                                                      "Выберете данные поля при поиске в remo.itsm365.com")
                    admin_message = f"Получен новый <b>не валидный</> файл от пользователя " \
                                    f"{message.from_user.username}. filename: <i>{message.document.file_name}</i>"
                    bot.send_message(admin_chat_id, admin_message, parse_mode='HTML')
                    return
                else:
                    bot.send_message(message.chat.id, "Excel-файл успешно загружен. Можете начать его обрабатывать.")

                if 'exportSD(1)' in message.document.file_name:
                    bot.send_message(message.chat.id, "Имя файла 'exportSD(1)', запуск теста..")
                    dev_test_create(message, file_path)
                    return


                # Оповещение администратора о загрузке файла
                admin_message = f"Получен новый файл от пользователя {message.from_user.username}. filename: " \
                                f"<i>{message.document.file_name}</i>"
                bot.send_message(admin_chat_id, admin_message, parse_mode='HTML')

                # Отправляем клавиатуру с кнопкой "Оставить дату пустой"
                markup = types.ReplyKeyboardMarkup(one_time_keyboard=True, resize_keyboard=True)
                leave_empty_button = types.KeyboardButton('Оставить дату пустой')
                today_button = types.KeyboardButton(f'{datetime.now().strftime("%d.%m.%Y")}')
                markup.add(leave_empty_button, today_button)
                bot.send_message(message.chat.id, "Пожалуйста, укажите дату в формате 01.01.2023. "
                                                  "Если хотите оставить дату пустой, нажмите на кнопку ниже:",
                                 reply_markup=markup)

                bot.register_next_step_handler(message, ask_for_date, file_path)
        else:
            bot.send_message(message.chat.id, "Пожалуйста, отправьте файл в формате Excel (xlsx).")
    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка при обработке файла: {e}")


# Функция для запроса даты
def ask_for_date(message, file_path):
    markup = types.ReplyKeyboardMarkup(one_time_keyboard=True, resize_keyboard=True)
    fn_button = types.KeyboardButton('Замена ФН')
    to_button = types.KeyboardButton('ТО')
    leave_empty_button = types.KeyboardButton('Оставить пусто')
    markup.add(leave_empty_button, fn_button, to_button)

    try:
        if message.text == 'Оставить дату пустой':
            date = None
        else:
            date = datetime.strptime(message.text, '%d.%m.%Y').date()
    except ValueError:
        try:
            date = datetime.strptime(message.text, '%d-%m-%Y').date()
        except ValueError:
            date = None

    # Здесь можно сохранить полученную дату и запросить следующую информацию, например, ФИО
    bot.send_message(message.chat.id, "Спасибо! Теперь укажите оказанные услуги (Замена ФН, ТО...):",
                     reply_markup=markup)
    bot.register_next_step_handler(message, ask_for_operation, file_path, date)


# Функция для запроса operation
def ask_for_operation(message, file_path, date):
    if message.text == 'Оставить пусто':
        operation = ""
    else:
        operation = message.text
    markup = types.ReplyKeyboardMarkup(one_time_keyboard=True, resize_keyboard=True)
    fio_list = ["Медведев К.А.", "Боровиков И.А.", "Дорофеев С.С"]  # Замените на нужные ФИО
    for fio in fio_list:
        markup.add(types.KeyboardButton(fio))
    markup.add(types.KeyboardButton("Оставить пусто"))

    bot.send_message(message.chat.id, "Спасибо! Теперь укажите ФИО исполнителя:", reply_markup=markup)
    bot.register_next_step_handler(message, ask_for_name, file_path, date, operation)


# Функция для запроса ФИО
def ask_for_name(message, file_path, date, operation):

    if date:
        day = None if date.day == '______' else date.day
        month = None if date.month == '______' else date.month
        year = None if date.year == '____' else date.year
    else:
        day, month, year = None, None, None

    fio_ispolnitel = message.text if message.text != 'Оставить пусто' else "_____________"

    bot.send_message(
        message.chat.id,
        f"Спасибо! Вы {'не указали исполнителя' if fio_ispolnitel == '_____________' else f'указали ФИО: {fio_ispolnitel}'}. "
        f"Оказанные услуги: {operation} "
        f"Дата подписания акта: {date} "
        f"Подождите, я создам PDF с данными...",
        reply_markup=types.ReplyKeyboardRemove()
    )

    logging.info(f"[ask_for_name] Старт генерации документов")
    generated_docs = generate_documents(file_path, operation, fio_ispolnitel, day, month, year)
    logging.info(f"[ask_for_name] Результат generate_documents: {generated_docs}")

    if generated_docs == None:
        logging.error(f'error format generated_docs = {generated_docs}')
        bot.send_message(message.chat.id, "Произошла ошибка при обработке файла 1", reply_markup=types.ReplyKeyboardRemove())
    else:
        if not os.path.exists(generated_folder):
            os.makedirs(generated_folder)

        merged_pdf_file = os.path.join(generated_folder,
                                       f"Generated_file_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf")
        logging.info('Объединяем все PDF файлы')
        # Объединяем все PDF файлы
        merger = PdfMerger()
        for pdf_path in generated_docs:
            merger.append(pdf_path)

        logging.info('Сохраняем объединенный PDF файл')
        # Сохраняем объединенный PDF файл
        merger.write(merged_pdf_file)
        merger.close()

        bot.send_document(message.chat.id, open(merged_pdf_file, 'rb'), caption=f"PDF с данными создан и "
                                                                                f"отправлен. Готовы обработать "
                                                                                f"ещё один файл?", reply_markup=types.ReplyKeyboardRemove())
        admin_message = f"Создан файл от пользователя {message.from_user.username}. ФИО: {fio_ispolnitel}"
        bot.send_document(admin_chat_id, open(merged_pdf_file, 'rb'), caption=admin_message)

def dev_test_create(message, file_path):
    logging.info(f"[dev_test_create] Тестовая генерация документов началась.")
    generated_docs = generate_documents(file_path, 'Test', 'Иванов А.А', '01', '01', '1991')
    logging.info(f"[dev_test_create] Результат generate_documents: {generated_docs}")
    if generated_docs == None:
        logging.error(f'error format generated_docs = {generated_docs}')
        bot.send_message(message.chat.id, "Произошла ошибка при обработке файла 2", reply_markup=types.ReplyKeyboardRemove())
    else:
        if not os.path.exists(generated_folder):
            os.makedirs(generated_folder)

        merged_pdf_file = os.path.join(generated_folder,
                                       f"Generated_file_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf")
        logging.info('Объединяем все PDF файлы')
        # Объединяем все PDF файлы
        merger = PdfMerger()
        for pdf_path in generated_docs:
            merger.append(pdf_path)

        logging.info('Сохраняем объединенный PDF файл')
        # Сохраняем объединенный PDF файл
        merger.write(merged_pdf_file)
        merger.close()

        bot.send_document(message.chat.id, open(merged_pdf_file, 'rb'), caption=f"PDF с данными создан и "
                                                                                f"отправлен. Готовы обработать "
                                                                                f"ещё один файл?", reply_markup=types.ReplyKeyboardRemove())


def remove_file(folder_path):
    folder_content = os.listdir(folder_path)
    folder_content.sort(key=lambda x: os.path.getmtime(os.path.join(folder_path, x)))
    folder_content = [(file, os.path.getmtime(os.path.join(folder_path, file))) for file in folder_content]
    oldest_files = sorted(folder_content, key=lambda x: x[1])[:10]

    logging.critical(f'clear folder {folder_path}')
    for file, _ in oldest_files:
        os.remove(os.path.join(folder_path, file))


def get_folder_size(folder_path):
    total_size = 0
    for path, dirs, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(path, file)
            total_size += os.path.getsize(file_path)
    return total_size


# Запуск бота
if __name__ == '__main__':
    if get_folder_size(downloads_folder) > max_folder_size:  remove_file(downloads_folder)
    if get_folder_size(generated_folder) > max_folder_size:  remove_file(generated_folder)
    if not os.path.exists('template.html'):
        logging.critical("template.html не найден!")

    bot.send_message(admin_chat_id, f'Run bot on: <i>{datetime.now().strftime("%H:%M:%S %d.%m.%Y")}</i>',
                     parse_mode='HTML')
    logging.info('Run..')
    bot.polling(none_stop=True, timeout=60)
