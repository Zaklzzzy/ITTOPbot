import telebot
import pandas as pd
import os
import re
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton

# Telebot Fields
TOKEN = "7774431217:AAEg7dUvK6s2EfLVJevlCsYxd-YIDYOg5oM"
bot = telebot.TeleBot(TOKEN)

TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)

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
    
# Bot Functions

#regionStart Menu
@bot.message_handler(commands=['start'])
def start(message):
    markup = InlineKeyboardMarkup()
    button = InlineKeyboardButton("Количество пар группы", callback_data="group_subjects")
    markup.add(button)
    bot.send_message(message.chat.id, "Выберите действие:", reply_markup=markup)

@bot.callback_query_handler(func=lambda call:call.data == "group_subjects")
def request_xlsx_file(call):
    bot.edit_message_text("Пришлите файл формата .xlsx", call.message.chat.id, call.message.message_id)

@bot.message_handler(content_types=['document'])
def handle_document(message):
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
        result = analyze_group_subjects(file_path)
        bot.reply_to(message, result)
    except Exception as e:
        bot.reply_to(message, f"Error: {e}")
    finally:
        if os.path.exists(file_path):
            os.remove(file_path) # Clear temp files
#endregion

print("Bot started...")
bot.polling(none_stop=True)