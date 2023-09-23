import telebot
from telebot import types
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

logging.basicConfig(level=logging.INFO, filename="py_log.log",filemode="w")


# Инициализация бота
dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)

bot = telebot.TeleBot(os.environ.get('API_KEY'))

if platform.system() == 'Linux':
    wkhtmltopdf = '/usr/bin/wkhtmltopdf'
else:
    wkhtmltopdf = './wkhtmltopdf.exe'


def generate_documents(exelfile, operation, fio_ispolnitel, day, month, year):
    # Загрузка данных из файла Excel
    global generated_docs
    df = pd.read_excel(exelfile)

    if not day: day = '__'
    if not month: month = '__'
    if not year: year = '____'

    logging.info(f'generate_documents   exelfile={exelfile},   operation={operation},  fio_ispolnitel={fio_ispolnitel}, day={day}')

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
            context = {'name_file':name_file,
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
            logging.info(f'one_pdf_crt(context)  = {file_path}')
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

    # Создаем временную папку
    temp_dir = tempfile.mkdtemp()

    # Генерируем уникальное имя файла
    filename = f"{context['name_file']}.pdf"

    # Путь к сгенерированному PDF файлу
    pdf_path = os.path.join(temp_dir, filename)

    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf)
    pdfkit.from_string(output_text, pdf_path, configuration=config)

    # Возвращаем путь к сгенерированному файлу
    return pdf_path

# Обработчик команды /start
@bot.message_handler(commands=['start'])
def start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    item = types.KeyboardButton("Загрузить Excel файл")
    markup.add(item)
    bot.send_message(message.chat.id, "Привет! Я бот для загрузки Excel-файлов. Нажмите кнопку, чтобы загрузить файл.",
                     reply_markup=markup)


# Обработчик текстовых сообщений
@bot.message_handler(content_types=['text'])
def handle_text(message):
    if message.text == "Загрузить Excel файл":
        bot.send_message(message.chat.id, "Пожалуйста, отправьте мне Excel-файл, выгруженный из remo.itsm365.com "
                                          "c столбцами: Задание,	Описание (RTF),	Адрес,	Объект обслуживания,"
                                          "	Статус,	Крайний срок решения,	Дата создания,	Number,	NumberIn,	"
                                          "Конфигурационная единица ")
    else:
        bot.send_message(message.chat.id, "Я понимаю только команду 'Загрузить Excel файл'.")

# Обработчик загрузки документов (Excel-файлов)
@bot.message_handler(content_types=['document'])
def handle_document(message):
    if message.document.mime_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
        file_info = bot.get_file(message.document.file_id)
        file_path = os.path.join("downloads", message.document.file_name)
        downloaded_file = bot.download_file(file_info.file_path)

        with open(file_path, 'wb') as new_file:
            new_file.write(downloaded_file)

        bot.send_message(message.chat.id, "Excel-файл успешно загружен. Можете начать его обрабатывать.")

        # Здесь можно добавить код для запроса дополнительных данных
        bot.send_message(message.chat.id, "Пожалуйста, укажите дату в формате 01.01.2023:")
        bot.register_next_step_handler(message, ask_for_date, file_path)
    else:
        bot.send_message(message.chat.id, "Пожалуйста, отправьте файл в формате Excel (xlsx).")


# Функция для запроса даты
def ask_for_date(message, file_path):
    date_str = message.text
    try:
        date = datetime.strptime(date_str, '%d.%m.%Y').date()
    except ValueError:
        try:
            date = datetime.strptime(date_str, '%d-%m-%Y').date()
        except ValueError:
            date = None

    # Здесь можно сохранить полученную дату и запросить следующую информацию, например, ФИО
    bot.send_message(message.chat.id, "Спасибо! Теперь укажите что делали (Замена ФН, ТО...):")
    bot.register_next_step_handler(message, ask_for_operation, file_path, date)

# Функция для запроса operation
def ask_for_operation(message, file_path, date):
    operation = message.text

    # Здесь можно сохранить полученную дату и запросить следующую информацию, например, ФИО
    bot.send_message(message.chat.id, "Спасибо! Теперь укажите ФИО:")
    bot.register_next_step_handler(message, ask_for_name, file_path, date, operation)

# Функция для запроса ФИО
def ask_for_name(message, file_path, date,operation):
    fio_ispolnitel = message.text
    # Здесь можно сохранить полученное ФИО и продолжить обработку файла
    bot.send_message(message.chat.id, f"Спасибо! Вы указали ФИО: {fio_ispolnitel}. Оказанные услуги: {operation} "
                                      f"Подождите, я создам PDF с данными...")

    if date:
        day = None if date.day == '____' else date.day
        month = None if date.month == '____' else date.month
        year = None if date.year == '____' else date.year
    else:
        day, month, year = None, None, None


    generated_docs = generate_documents(file_path, operation, fio_ispolnitel, day, month, year)
    if generated_docs == None:
        logging.error(f'error format generated_docs = {generated_docs}')
        bot.send_message(message.chat.id, "Excel-файл имеет неверное содержание . Он должен содержать столбцы Задание,	"
                                          "Описание (RTF),	Адрес,	Объект обслуживания,	Статус,	Крайний срок решения,	"
                                          "Дата создания,	Number,	NumberIn,	Конфигурационная единица. "
                                          "Выберете данные поля при поиске в remo.itsm365.com")

    else:
        merged_pdf_file = os.path.join('files', f"Generated_file_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf")
        logging.info('Объединяем все PDF файлы')
        # Объединяем все PDF файлы
        merger = PdfMerger()
        for pdf_path in generated_docs:
            merger.append(pdf_path)

        logging.info('Сохраняем объединенный PDF файл')
        # Сохраняем объединенный PDF файл
        merger.write(merged_pdf_file)
        merger.close()

        bot.send_document(message.chat.id, open(merged_pdf_file, 'rb'))
        bot.send_message(message.chat.id, f"PDF с данными создан и отправлен. Готовы обработать ещё один файл?")


# Запуск бота
if __name__ == '__main__':
    logging.info('Run..')
    bot.polling(none_stop=True)