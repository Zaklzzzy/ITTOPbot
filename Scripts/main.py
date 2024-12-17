import re
import telebot
import pandas as pd
import os
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton

# Telebot Fields
TOKEN = "7774431217:AAHiNLWfWzlQCx71maPMQpa3cAeYGmcsvAw"
bot = telebot.TeleBot(TOKEN)

TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)

DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)

USER_STATE = {}

def analyze_group_subjects(file_path: str):
    """
    Анализирует таблицу и выводит список дисцплин с количеством пар за неделю для группы

    :param file_path: Путь к файлу .xlsx
    :return: Форматированная строка со списком пар за неделю
    """
    df = pd.read_excel(file_path)

    group_name = df['Группа'].iloc[0]

    week_columns = [col for col in df.columns if re.search(r'Понедельник|Вторник|Среда|Четверг|Пятница|Суббота', col, re.IGNORECASE)]
    
    subject_count = {}

    # Calculating lessons count
    for day in week_columns:
        for cell in df[day]:
            if isinstance(cell,str):
                match = re.search(r'Предмет: (.*?)\n', cell)
                if match:
                    subject = match.group(1).strip()
                    subject_count[subject] = subject_count.get(subject, 0) + 1
    # Make result
    if not subject_count:
        return f"Для группы {group_name} не найдено расписание"
    
    result = f"\nУ группы {group_name} на этой неделе было:\n"
    for subject, count in subject_count.items():
        result += f"{subject} - {count}\n"
    
    # Return lessons list
    return result

def analyze_low_attendance(file_path: str):
    """
    Анализирует отчет по посещаемости и возвращает список преподавателей с посещаемостью ниже 65%

    :param file_path: Путь к файлу .xlsx
    :return: Форматированная строка со списком преподавателей
    """
    df = pd.read_excel(file_path)
    
    df = df[:-1]

    df['Средняя посещаемость'] = df['Средняя посещаемость'].str.replace('%', '').astype(float)

    low_attendance = df[df['Средняя посещаемость'] < 65]

    if low_attendance.empty:
        return "Все преподаватели имеют посещаемость выше 65%"
    
    result = "Список преподавателей с посещаемостью групп ниже 65%:\n"
    for _, row in low_attendance.iterrows():
        result += f"{row['ФИО преподавателя']} - {row['Средняя посещаемость']}%\n"

    return result


# Bot Functions

#regionStart Menu
@bot.message_handler(commands=['start'])
def start(message):
    markup = InlineKeyboardMarkup()
    button1 = InlineKeyboardButton("Количество пар группы", callback_data="group_subjects")
    button2 = InlineKeyboardButton("Посещаемость ниже 65%", callback_data="low_attendance")
    markup.add(button1, button2)
    bot.send_message(message.chat.id, "Выберите действие:", reply_markup=markup)
    #print(message.chat.id) Only for get admin chatID

# Number of lessons of the group
@bot.callback_query_handler(func=lambda call:call.data == "group_subjects")
def request_xlsx_file(call):
    USER_STATE[call.message.chat.id] = "group_subjects"
    bot.edit_message_text("Бот подсчитает количество проведенных пар по всем дисциплинам\nПришлите файл формата .xlsx", call.message.chat.id, call.message.message_id)
# Attendance below 65%
@bot.callback_query_handler(func=lambda call: call.data == "low_attendance")
def request_attendance_file(call):
    USER_STATE[call.message.chat.id] = "low_attendance"
    bot.edit_message_text("Бот выведет список преподавателей, средняя посещаемость которых ниже 65%\nПришлите отчет по посещаемости студентов в формате .xlsx", call.message.chat.id, call.message.message_id)

# Download handler .xlsx files
@bot.message_handler(content_types=['document'])
def handle_document(message):
    
    chat_id = message.chat.id
    user_state = USER_STATE.get(chat_id)
    
    if user_state not in ["group_subjects", "low_attendance"]:
        bot.reply_to(message, "Пожалуйста, выберите действие из меню.")
        return

    file_name = message.document.file_name
    _, file_extension = os.path.splitext(file_name)
    file_extension = file_extension.lower()

    # Catch incorrect file type
    if not file_extension == '.xlsx':
        bot.reply_to(message, "Пожалуйста, отправьте файл в формате .xlsx")
        return
    
    # Downloading file
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    file_path = os.path.join(TEMP_DIR, file_name)

    with open(file_path, "wb") as new_file:
        new_file.write(downloaded_file)

    try:
        if user_state == "low_attendance":
            result = analyze_low_attendance(file_path)
        elif user_state == "group_subjects":
            result = analyze_group_subjects(file_path)
        else:
            result = "Неизвестное действие. Попробуйте снова."

        bot.reply_to(message, result)
    except Exception as e:
        bot.reply_to(message, f"Error: {e}")
    finally:
        if os.path.exists(file_path):
            os.remove(file_path) # Clear temp files
#endregion

print("Bot started...")
bot.polling(none_stop=True)