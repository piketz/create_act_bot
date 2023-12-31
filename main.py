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

logging.basicConfig(level=logging.INFO, filename="py_log.log", filemode="w")

# Инициализация бота
dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)

apihelper.ENABLE_MIDDLEWARE = True
apihelper.SESSION_TIME_TO_LIVE = 5 * 60
bot = telebot.TeleBot(os.environ.get('API_KEY'))
admin_chat_id = os.environ.get('ADMIN_CHAT_ID')  # Загрузите ID чата администратора из переменной окружения

if platform.system() == 'Linux':
    wkhtmltopdf = '/bin/wkhtmltopdf'
else:
    wkhtmltopdf = './wkhtmltox/bin/wkhtmltopdf.exe'

downloads_folder = 'downloads'
generated_folder = 'files'
max_folder_size = 1000 * 1024 * 1024  # 1000 МБ в байтах

def check_file_content(exelfile):
    try:
        # Загрузка данных из файла Excel
        df = pd.read_excel(exelfile)

        # Проверка наличия конкретных столбцов в файле
        required_columns = ['Дата создания', 'Конфигурационная единица', 'Объект обслуживания',
                             'NumberIn', 'Number', 'Задание']

        if not set(required_columns).issubset(set(df.columns)):
            logging.error(f"File does not contain all required columns: {required_columns}")
            return False

        logging.info("File content is valid.")
        return True
    except Exception as e:
        logging.error(f"Error checking file content: {e}")
        return False


def generate_documents(exelfile, operation, fio_ispolnitel, day, month, year):
    # Загрузка данных из файла Excel
    global generated_docs
    df = pd.read_excel(exelfile)

    if not day: day = '__'
    if not month: month = '__'
    if not year: year = '____'

    logging.info(f'generate_documents '
                 f'exelfile={exelfile}, '
                 f'operation={operation}, '
                 f'fio_ispolnitel={fio_ispolnitel}, '
                 f'day={day}')

    generated_docs = []
    for index, row in df.iterrows():
        try:
            date_str = str(row['Дата создания'])
            date_obj = parser.parse(date_str)

            config_data = str(row['Конфигурационная единица'])
            config_parts = [part.strip() for part in config_data.split('|')]
            index_ops = str(row['Объект обслуживания']).split()[0]
            name_file = f"{index_ops}_{str(row['NumberIn'])}_{str(row['Number'])}_{str(row['Задание']).replace('/', '-')}"

            # Создание контекста для заполнения шаблона
            context = {'name_file': name_file,
                       'nn': str(row['Задание']),
                       'fio_ispolnitel': fio_ispolnitel,
                       'day': day,
                       'month': month,
                       'year': year,
                       'day_crt': date_obj.day,
                       'month_crt': date_obj.month,
                       'year_crt': date_obj.year,
                       'index_adress': str(row['Объект обслуживания']),
                       'model_ke': f'{config_parts[1]} {config_parts[2]} {config_parts[3]}',
                       'num_ke': config_parts[0],
                       'work': operation,
                       'num_rp': str(row['NumberIn']),
                       'num_im': str(row['Number'])}

            file_path = one_pdf_crt(context)
            # logging.info(f'one_pdf_crt(context)  = {file_path}')
            generated_docs.append(file_path)
        except Exception as e:
            logging.error(f"Error processing row: {e}")
            return None

    logging.info(f'return generated_docs = {generated_docs}')
    return generated_docs


def one_pdf_crt(context):
    template_loader = jinja2.FileSystemLoader('./')
    template_env = jinja2.Environment(loader=template_loader)
    template = template_env.get_template('template.html')
    output_text = template.render(context)

    temp_dir = tempfile.mkdtemp()
    pdf_path = os.path.join(temp_dir, f"{context['name_file']}.pdf")
    options = {  "enable-local-file-access": None}
    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf)
    pdfkit.from_string(output_text, pdf_path, configuration=config, options=options)

    return pdf_path


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
                                      "	Статус,	Крайний срок решения,	Дата создания,	Number,	NumberIn,	"
                                      "Конфигурационная единица ", reply_markup=types.ReplyKeyboardRemove())


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
    markup = types.ReplyKeyboardMarkup(one_time_keyboard=True, resize_keyboard=True)
    leave_empty_button = types.KeyboardButton('Оставить пусто')
    markup.add(leave_empty_button)

    if message.text == 'Оставить пусто':
        operation = ""
    else:
        operation = message.text
    bot.send_message(message.chat.id, "Спасибо! Теперь укажите ФИО исполнителя:",reply_markup=markup)
    bot.register_next_step_handler(message, ask_for_name, file_path, date, operation)


# Функция для запроса ФИО
def ask_for_name(message, file_path, date, operation):

    if date:
        day = None if date.day == '______' else date.day
        month = None if date.month == '______' else date.month
        year = None if date.year == '____' else date.year
    else:
        day, month, year = None, None, None
    if message.text == 'Оставить пусто':
        fio_ispolnitel = "_____________"
        bot.send_message(message.chat.id, f"Спасибо! Вы не указали исполнителя. Оказанные услуги: {operation} "
                                          f"Дата подписания акта: {date} "
                                          f"Подождите, я создам PDF с данными...", reply_markup=types.ReplyKeyboardRemove())
    else:
        fio_ispolnitel = message.text
        bot.send_message(message.chat.id, f"Спасибо! Вы указали ФИО: {fio_ispolnitel}. Оказанные услуги: {operation} "
                                          f"Дата подписания акта: {date} "
                                          f"Подождите, я создам PDF с данными...", reply_markup=types.ReplyKeyboardRemove())

    generated_docs = generate_documents(file_path, operation, fio_ispolnitel, day, month, year)
    if generated_docs == None:
        logging.error(f'error format generated_docs = {generated_docs}')
        bot.send_message(message.chat.id, "Произошла ошибка или Excel-файл имеет неверное содержание. "
                                          "Он должен содержать столбцы Задание,	"
                                          "Описание (RTF),	Адрес,	Объект обслуживания, "
                                          "Статус, Крайний срок решения, "
                                          "Дата создания, Number,	NumberIn, "
                                          "Конфигурационная единица. "
                                          "Выберете данные поля при поиске в remo.itsm365.com", reply_markup=types.ReplyKeyboardRemove())
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
    bot.send_message(admin_chat_id, f'Run bot on: <i>{datetime.now().strftime("%H:%M:%S %d.%m.%Y")}</i>',
                     parse_mode='HTML')
    logging.info('Run..')
    bot.polling(none_stop=True, timeout=60)
