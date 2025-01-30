import re
import telebot
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton
import pandas as pd
import os
import xlrd
import openpyxl
import json
import shutil
import atexit
import actions
import utils
from dotenv import load_dotenv

load_dotenv("config.env")

# Telebot Fields
TOKEN = os.getenv("TOKEN")
bot = telebot.TeleBot(TOKEN)

TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)

DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)

USER_STATE = {}

full_result = []

# JSON functions
def get_teachers():
    try:
        with open("data/teachers.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}
def save_teacher(username, chat_id=None, full_name=None):
    try:
        teachers = get_teachers()
        if username in teachers:
            teachers[username]["chat_id"] = chat_id
        else:
            teachers[username] = {"chat_id": chat_id, "full_name": full_name}
        
        with open("data/teachers.json", "w", encoding="utf-8") as f:
            json.dump(teachers, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        print(f"Error with teacher save: {e}")
        return False

# Bot Functions

#regionStart Menu
def send_menu(chat_id):
    markup = InlineKeyboardMarkup()
    group_subjects_button = InlineKeyboardButton("Пары группы", callback_data="group_subjects")
    checked_homeworks_button = InlineKeyboardButton("Проверенные ДЗ", callback_data="checked_homeworks")
    given_homeworks_button = InlineKeyboardButton("Выданные ДЗ", callback_data="given_homeworks")
    topic_check_button = InlineKeyboardButton("Тема урока", callback_data="topic_check")
    low_attendance_button = InlineKeyboardButton("Посещаемость", callback_data="low_attendance")
    low_homework_percentage_button = InlineKeyboardButton("Выполнение ДЗ", callback_data="low_homework_percentage")
    marks_analysis_button = InlineKeyboardButton("Анализ успеваемости", callback_data="marks_analysis")

    markup.add(
        group_subjects_button, 
        checked_homeworks_button, 
        given_homeworks_button, 
        topic_check_button, 
        low_attendance_button, 
        low_homework_percentage_button, 
        marks_analysis_button
        )
    
    message_text = (
        "Выберите действие:\n\n"
        "1️⃣ *Пары группы* - количество пар группы по всем дисциплинам за неделю\n"
        "2️⃣ *Проверенные ДЗ* - проверка % проверенных заданий\n"
        "3️⃣ *Выданные ДЗ* - проверка % выданных заданий\n"
        "4️⃣ *Тема урока* - проверка соответствия темы урока шаблону \"Урок № . Тема:\"\n"
        "5️⃣ *Посещаемость* - анализ посещаемости у преподавателей\n"
        "6️⃣ *Выполнение ДЗ* - анализ выполнения ДЗ студентами\n"
        "7️⃣ *Анализ успеваемости* - информация успеваемости студентов"
    )

    bot.send_message(chat_id, message_text, reply_markup=markup, parse_mode="Markdown")
    #print(chat_id) #Only for get admin chatID

# Admin Commands
@bot.message_handler(commands=['add_teacher'])
def add_teacher(message):
    if message.chat.id != 1129590158:
        bot.reply_to(message, "Нет доступа к команде")
        send_menu(message.chat.id)
        return
    bot.reply_to(message, "Добавление преподавателя\nНапишите в первой строчке @username, во второй ФИО преподавателя\nОтветьте на это сообщение, чтобы добавить преподавателя")
@bot.message_handler(func=lambda message: message.reply_to_message and message.reply_to_message.text.startswith("Добавление преподавателя"))
def handle_teacher_input(message):
    if message.chat.id != 1129590158:
        bot.reply_to(message, "Нет доступа к данному действию")
        send_menu(message.chat.id)
        return
    
    # Split input to 2 rows
    user_input = message.text.split("\n")
    if len(user_input) != 2:
        bot.reply_to(message, "Ошибка: формат данных неправильный\nНапишите в первой строчке @username, во второй ФИО преподавателя")
        return
    
    username, full_name = user_input
    username = username.strip()
    full_name = full_name.strip()

    # Check correct data
    if not username.startswith("@") or len(username) < 2:
        bot.reply_to(message, "Ошибка: username должен начинаться с '@' и содержать хотя бы один символ после")
        return
    
    if save_teacher(username, None, full_name):
        bot.reply_to(
            message,
            f"Преподаватель добавлен:\nUsername: {username}\nФИО: {full_name}\nchat_id: None"
        )
    else:
        bot.reply_to(message, "Ошибка: не удалось сохранить преподавателя")
@bot.message_handler(commands=['show_teachers'])
def show_teachers(message):
    if message.chat.id != 1129590158:
        bot.reply_to(message, "Нет доступа к команде")
        send_menu(message.chat.id)
        return
    
    teachers = get_teachers()
    if not teachers:
        bot.reply_to(message, "Список преподавателей пуст")
        return
    
    result = "Список преподавателей:\n"
    for username, data in teachers.items():
        full_name = data.get("full_name", "Неизвестно")
        result += f"{username} : {full_name}\n"
    bot.reply_to(message, result)

# Common commands
@bot.message_handler(commands=['start'])
def start(message):
    username = f"@{message.from_user.username}"
    teachers = get_teachers()
    if username in teachers:
        if teachers[username]["chat_id"] == message.chat.id:
            return
        
        save_teacher(username, chat_id=message.chat.id)
        bot.reply_to(message, "Ваш ID успешно зарегистрирован")

@bot.message_handler(commands=['menu'])
def menu(message):
    send_menu(message.chat.id)

# 1. Number of lessons of the group
@bot.callback_query_handler(func=lambda call:call.data == "group_subjects")
def request_xlsx_file(call):
    USER_STATE[call.message.chat.id] = "group_subjects"
    bot.edit_message_text("Бот подсчитает количество проведенных пар по всем дисциплинам\nПришлите расписание группы в формате .xls или .xlsx", call.message.chat.id, call.message.message_id)

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
    message_part = "Пришлите отчет по домашним заданиям формате .xls или .xlsx"
    match call.data:
        case "checked_month":
            bot.edit_message_text("Бот подсчитает % проверенных домашних заданий педагогами на группу за месяц\n" + message_part, call.message.chat.id, call.message.message_id)
        case "checked_week":
            bot.edit_message_text("Бот подсчитает % проверенных домашних заданий педагогами на группу за неделю\n" + message_part, call.message.chat.id, call.message.message_id)
        case "given_month":
            bot.edit_message_text("Бот подсчитает % выданных домашних заданий педагогами за месяц\n" + message_part, call.message.chat.id, call.message.message_id)
        case "given_week":
            bot.edit_message_text("Бот подсчитает % выданных домашних заданий педагогами за неделю\n" + message_part, call.message.chat.id, call.message.message_id)

# 4. Lessons topic check
@bot.callback_query_handler(func=lambda call: call.data == "topic_check")
def request_topic_file(call):
    USER_STATE[call.message.chat.id] = "topic_check"
    bot.edit_message_text("Бот выведет список преподавателей и тем, не соответствующих шаблону \"Урок № . Тема:\"\nПришлите отчет по темам уроков в формате .xls или .xlsx", call.message.chat.id, call.message.message_id)

# 5. Attendance below 65%
@bot.callback_query_handler(func=lambda call: call.data == "low_attendance")
def request_attendance_file(call):
    USER_STATE[call.message.chat.id] = "low_attendance"
    bot.edit_message_text("Бот выведет список преподавателей, средняя посещаемость которых ниже 65%\nПришлите отчет по посещаемости студентов в формате .xls или .xlsx", call.message.chat.id, call.message.message_id)

# 6. Low Homework Percentage
@bot.callback_query_handler(func=lambda call: call.data == "low_homework_percentage")
def request_attendance_file(call):
    USER_STATE[call.message.chat.id] = "low_homework_percentage"
    bot.edit_message_text("Бот выведет список студентов, процент выполнения ДЗ которых ниже 50%\nПришлите отчет по студентам в формате .xls или .xlsx", call.message.chat.id, call.message.message_id)

# 7. Marks analysis
@bot.callback_query_handler(func=lambda call: call.data == "marks_analysis")
def request_attendance_file(call):
    USER_STATE[call.message.chat.id] = "marks_analysis"
    bot.edit_message_text("Бот выведет список студентов, средняя оценка которых ниже 3\nПришлите отчет по студентам в формате .xls или .xlsx", call.message.chat.id, call.message.message_id)

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
            "topic_check", # 4. Lessons topic check
            "low_attendance", # 5. Attendance below 65%
            "low_homework_percentage", # 6. Low Homework Percentage
            "marks_analysis" # 7. Marks analysis
            ]:
        bot.reply_to(message, "Пожалуйста, выберите действие из меню")
        send_menu(chat_id)
        return

    file_name = message.document.file_name
    _, file_extension = os.path.splitext(file_name)
    file_extension = file_extension.lower()

    # Catch incorrect file type
    if file_extension == '.xls':
        file_path = utils.download_and_convert_xls(bot, message.document.file_id, TEMP_DIR, file_name)
    elif file_extension == '.xlsx':
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        file_path = os.path.join(TEMP_DIR, file_name)

        with open(file_path, "wb") as new_file:
            new_file.write(downloaded_file)
    else:
        bot.reply_to(message, "Пожалуйста, отправьте файл в формате .xls или .xlsx")
        send_menu(chat_id)
        return

    try:
        match user_state:
            case "group_subjects":
                result = actions.analyze_group_subjects(file_path)
            case "checked_month":
                  result = actions.analyze_checked_homeworks(file_path, "month")
            case "checked_week":
                  result = actions.analyze_checked_homeworks(file_path, "week")
            case "given_month":
                  result = actions.analyze_given_homeworks(file_path, "month")
            case "given_week":
                  result = actions.analyze_given_homeworks(file_path, "week")
            case "topic_check":
                result = actions.analyze_lessons_topic(file_path)
            case "low_attendance":
                result = actions.analyze_low_attendance(file_path)
            case "low_homework_percentage":
                result = actions.analyze_low_homework_percentage(file_path)
            case "marks_analysis":
                result = actions.analyze_bad_marks(file_path)
            case _:
                result = "Неизвестное действие. Попробуйте снова"

        if len(result) >= 4096:
            messages = utils.split_message(result)
            for msg in messages:
                bot.reply_to(message, msg)
        else:
            bot.reply_to(message, result)

        send_menu(chat_id)
    except Exception as e:
        bot.reply_to(message, f"Ошибка: {e}")
        send_menu(chat_id)
    finally:
        # Clear temp files
        if os.path.exists(file_path):
            os.remove(file_path) 

atexit.register(utils.clean_temp_folder)

print("Bot started...")
bot.polling(none_stop=True)