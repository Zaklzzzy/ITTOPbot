import re
import telebot
import pandas as pd
import os
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton

# Telebot Fields
TOKEN = "token"
bot = telebot.TeleBot(TOKEN)

TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)

DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)

USER_STATE = {}

# 1. Number of lessons of the group
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
# 2. Checked homeworks
def analyze_checked_homeworks(file_path: str, period):
    """
    Анализирует проверенные ДЗ на процент выполнения < 75%

    :param file_path: Путь к файлу .xlsx
    :param period: Период, за который анализируются данные
    :return: Форматированная строка со списком преподавателей и процентом проверенных ДЗ
    """
    df = pd.read_excel(file_path, header=1)
    if period == "month":
        received_index, checked_index = 4, 5
    elif period == "week":
        received_index, checked_index = 9, 10
    else:
        return "Неверный период"
    
    results = []
    for _, row in df.iterrows():
        received = row.iloc[received_index]
        checked = row.iloc[checked_index]
        percentage = (checked / received * 100) if received else 0
        if percentage < 75:
            results.append(f"{row['Unnamed: 1']}: {percentage:.1f}% (Проверено {checked} из {received})")

    return "\n".join(results) if results else "Все ДЗ проверены более чем на 75%"
# 3. Given homeworks
def analyze_given_homeworks(file_path: str, period):
    """
    Анализирует выданные ДЗ на процент выполнения < 70%

    :param file_path: Путь к файлу .xlsx
    :param period: Период, за который анализируются данные
    :return: Форматированная строка со списком преподавателей и процентом выданных ДЗ
    """
    df = pd.read_excel(file_path, header=1)
    if period == "month":
        given_index, planned_index = 3, 6
    elif period == "week":
        given_index, planned_index = 8, 11
    else:
        return "Неверный период"
    
    results = []
    for _, row in df.iterrows():
        given = row.iloc[given_index]
        planned = row.iloc[planned_index]
        percentage = (given / planned * 100) if planned else 0
        if percentage < 70:
            results.append(f"{row['Unnamed: 1']}: {percentage:.1f}% (Выдано {given} из {planned})")

    return "\n".join(results) if results else "Все ДЗ выданы более чем на 70%."
# 4. Attendance below 65%
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
        result += f"{row['ФИО преподавателя']} - {row['Средняя посещаемость']:.0f}%\n"

    return result


# Bot Functions

#regionStart Menu
def send_menu(chat_id):
    markup = InlineKeyboardMarkup()
    group_subjects_button = InlineKeyboardButton("Пары группы", callback_data="group_subjects")
    checked_homeworks_button = InlineKeyboardButton("Проверенные ДЗ", callback_data="checked_homeworks")
    given_homeworks_button = InlineKeyboardButton("Выданные ДЗ", callback_data="given_homeworks")
    low_attendance_button = InlineKeyboardButton("Посещаемость", callback_data="low_attendance")
    button5 = InlineKeyboardButton("Тема урока", callback_data="empty")
    button6 = InlineKeyboardButton("Выполнение ДЗ", callback_data="empty")
    button7 = InlineKeyboardButton("Анализ успеваемости", callback_data="empty")

    markup.add(
        group_subjects_button, 
        checked_homeworks_button, 
        given_homeworks_button, 
        low_attendance_button, 
        button5, 
        button6, 
        button7
        )
    
    message_text = (
        "Выберите действие:\n\n"
        "1️⃣ *Пары группы* - количество пар группы по всем дисциплинам за неделю\n"
        "2️⃣ *Проверенные ДЗ* - проверка % проверенных заданий\n"
        "3️⃣ *Выданные ДЗ* - проверка % выданных заданий\n"
        "4️⃣ *Посещаемость* - анализ посещаемости у преподавателей\n"
        "5️⃣ *Тема урока* - проверка соответствия темы урока шаблону \"Урок №. Тема:\"\n"
        "6️⃣ *Выполнение ДЗ* - анализ выполнения ДЗ студентами\n"
        "7️⃣ *Анализ успеваемости* - информация успеваемости студентов"
    )

    bot.send_message(chat_id, message_text, reply_markup=markup, parse_mode="Markdown")
    #print(message.chat.id) Only for get admin chatID

@bot.message_handler(commands=['start'])
def start(message):
    send_menu(message.chat.id)

# 1. Number of lessons of the group
@bot.callback_query_handler(func=lambda call:call.data == "group_subjects")
def request_xlsx_file(call):
    USER_STATE[call.message.chat.id] = "group_subjects"
    bot.edit_message_text("Бот подсчитает количество проведенных пар по всем дисциплинам\nПришлите файл формата .xlsx", call.message.chat.id, call.message.message_id)

# 2. Checked homeworks
@bot.callback_query_handler(func=lambda call:call.data == "checked_homeworks")
def choose_period_checked(call):
    markup = InlineKeyboardMarkup()
    month_button = InlineKeyboardButton("Месяц", callback_data="checked_month")
    week_button = InlineKeyboardButton("Неделя", callback_data="checked_week")
    markup.row(month_button, week_button)
    bot.edit_message_text("Выберите период для анализа проверенных ДЗ:", call.message.chat.id, call.message.message_id, reply_markup=markup)

# 3. Given homeworks
@bot.callback_query_handler(func=lambda call:call.data == "given_homeworks")
def choose_period_given(call):
    markup = InlineKeyboardMarkup()
    month_button = InlineKeyboardButton("Месяц", callback_data="given_month")
    week_button = InlineKeyboardButton("Неделя", callback_data="given_week")
    markup.row(month_button, week_button)
    bot.edit_message_text("Выберите период для анализа выданных ДЗ:", call.message.chat.id, call.message.message_id, reply_markup=markup)

# 2-3. Homeworks file handler
@bot.callback_query_handler(func=lambda call: call.data in ["checked_month", "checked_week", "given_month", "given_week"])
def request_homeworks_file(call):
    USER_STATE[call.message.chat.id] = call.data
    match call.data:
        case "checked_month":
            bot.edit_message_text("Бот подсчитает % проверенных домашних заданий педагогами на группу за месяц\nПришлите файл формата .xlsx", call.message.chat.id, call.message.message_id)
        case "checked_week":
            bot.edit_message_text("Бот подсчитает % проверенных домашних заданий педагогами на группу за неделю\nПришлите файл формата .xlsx", call.message.chat.id, call.message.message_id)
        case "given_month":
            bot.edit_message_text("Бот подсчитает % выданных домашних заданий педагогами за месяц\nПришлите файл формата .xlsx", call.message.chat.id, call.message.message_id)
        case "given_week":
            bot.edit_message_text("Бот подсчитает % выданных домашних заданий педагогами за неделю\nПришлите файл формата .xlsx", call.message.chat.id, call.message.message_id)
    

# 4. Attendance below 65%
@bot.callback_query_handler(func=lambda call: call.data == "low_attendance")
def request_attendance_file(call):
    USER_STATE[call.message.chat.id] = "low_attendance"
    bot.edit_message_text("Бот выведет список преподавателей, средняя посещаемость которых ниже 65%\nПришлите отчет по посещаемости студентов в формате .xlsx", call.message.chat.id, call.message.message_id)

# Download handler .xlsx files
@bot.message_handler(content_types=['document'])
def handle_document(message):
    
    chat_id = message.chat.id
    user_state = USER_STATE.get(chat_id)
    
    if user_state not in [
            "group_subjects", # 1. Number of lessons of the group
            "checked_month", # 2.1. Checked homeworks (month)
            "checked_week", # 2.2 Checked homeworks (week)
            "given_month", # 3.1 Given homeworks (month)
            "given_week", # 3.2 Given homeworks (week)
            "low_attendance" # 4. Attendance below 65%
            ]:
        bot.reply_to(message, "Пожалуйста, выберите действие из меню.")
        send_menu(message.chat.id)
        return

    file_name = message.document.file_name
    _, file_extension = os.path.splitext(file_name)
    file_extension = file_extension.lower()

    # Catch incorrect file type
    if not file_extension == '.xlsx':
        bot.reply_to(message, "Пожалуйста, отправьте файл в формате .xlsx")
        send_menu(message.chat.id)
        return
    
    # Downloading file
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    file_path = os.path.join(TEMP_DIR, file_name)

    with open(file_path, "wb") as new_file:
        new_file.write(downloaded_file)

    try:
        match user_state:
            case "group_subjects":
                result = analyze_group_subjects(file_path)
            case "checked_month":
                  result = analyze_checked_homeworks(file_path, "month")
            case "checked_week":
                  result = analyze_checked_homeworks(file_path, "week")
            case "given_month":
                  result = analyze_given_homeworks(file_path, "month")
            case "given_week":
                  result = analyze_given_homeworks(file_path, "week")
            case "low_attendance":
                result = analyze_low_attendance(file_path)
            case _:
                result = "Неизвестное действие. Попробуйте снова."

        bot.reply_to(message, result)
        send_menu(message.chat.id)
    except Exception as e:
        bot.reply_to(message, f"Error: {e}")
        send_menu(message.chat.id)
    finally:
        if os.path.exists(file_path):
            os.remove(file_path) # Clear temp files
#endregion

print("Bot started...")
bot.polling(none_stop=True)