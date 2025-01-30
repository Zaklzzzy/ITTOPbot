import pandas as pd
import re

# 1. Number of lessons of the group
def analyze_group_subjects(file_path: str):
    """
    Анализирует таблицу и выводит список дисцплин с количеством пар за неделю для группы

    :param file_path: Путь к файлу .xlsx
    :return: Форматированная строка со списком пар за неделю
    """
    try:
        df = pd.read_excel(file_path)

        # Get all need columns
        group_name = df['Группа'].iloc[0]
        week_columns = [col for col in df.columns if re.search(r'Понедельник|Вторник|Среда|Четверг|Пятница|Суббота', col, re.IGNORECASE)]

        # Calculating lessons count
        subject_count = {}
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
    except Exception as e:
        return f"Ошибка при анализе: {e}"
# 2. Checked homeworks
def analyze_checked_homeworks(file_path: str, period):
    """
    Анализирует проверенные ДЗ на процент выполнения < 75%

    :param file_path: Путь к файлу .xlsx
    :param period: Период, за который анализируются данные
    :return: Форматированная строка со списком преподавателей и процентом проверенных ДЗ
    """
    try:
        df = pd.read_excel(file_path, header=1)

        # Get all need columns
        if period == "month":
            received_index, checked_index = 4, 5
        elif period == "week":
            received_index, checked_index = 9, 10
        else:
            return "Ошибка: Неверный период"
        
        results = []
        for _, row in df.iterrows():
            received = row.iloc[received_index]
            checked = row.iloc[checked_index]
            percentage = (checked / received * 100) if received else 0
            if percentage < 75:
                results.append(f"{row['Unnamed: 1']}: {percentage:.1f}% (Проверено {checked} из {received})")

        return "\n".join(results) if results else "Все ДЗ проверены более чем на 75%"
    except Exception as e:
        return f"Ошибка при анализе: {e}"
# 3. Given homeworks
def analyze_given_homeworks(file_path: str, period):
    """
    Анализирует выданные ДЗ на процент выполнения < 70%

    :param file_path: Путь к файлу .xlsx
    :param period: Период, за который анализируются данные
    :return: Форматированная строка со списком преподавателей и процентом выданных ДЗ
    """
    try:
        df = pd.read_excel(file_path, header=1)

        # Get all need columns
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
    except Exception as e:
        return f"Ошибка при анализе: {e}"
# 4. Lessons topic check
def analyze_lessons_topic(file_path: str):
    """
    Анализирует темы уроков на соответствие маске “Урок №.* Тема:*”

    :param file_path: Путь к файлу .xlsx
    :return: Форматированный список несоответствий маске
    """
    try:
        df = pd.read_excel(file_path)

        # Get all need columns
        topic_column = next((col for col in df.columns if "Тема" in col), None)
        teacher_column = next((col for col in df.columns if "ФИО преподавателя" in col), None)
        
        if not topic_column:
            return "Ошибка: Не найдены необходимые столбцы (Date, Тема, ФИО преподавателя)"
        
        # Checking for compliance with the mask
        incorrect_rows = []
        for _, row in df.iterrows():
            teacher = row[teacher_column]
            topic = row[topic_column]

            if not isinstance(topic, str) or not re.match(r"Урок №.* Тема:.*", topic):
                incorrect_rows.append(f"Преподаватель: {teacher}")
        
        # Make result
        if incorrect_rows:
            result = "Список преподавателей, заполнивших тему не по шаблону:\n" + "\n".join(set(incorrect_rows))
        else:
            result = "Все темы соответствуют маске Урок №. Тема:"
        return result
    except Exception as e:
        return f"Ошибка при анализе: {e}"
# 5. Attendance below 65%
def analyze_low_attendance(file_path: str):
    """
    Анализирует отчет по посещаемости и возвращает список преподавателей с посещаемостью ниже 65%

    :param file_path: Путь к файлу .xlsx
    :return: Форматированная строка со списком преподавателей
    """
    try:
        df = pd.read_excel(file_path)
        
        df = df[:-1]

        df['Средняя посещаемость'] = df['Средняя посещаемость'].str.replace('%', '').astype(float)

        low_attendance = df[df['Средняя посещаемость'] < 65]

        if low_attendance.empty:
            return "Все преподаватели имеют посещаемость выше 65%"
        
        # Make result
        result = "Список преподавателей с посещаемостью групп ниже 65%:\n"
        for _, row in low_attendance.iterrows():
            result += f"{row['ФИО преподавателя']} - {row['Средняя посещаемость']:.0f}%\n"

        return result
    except Exception as e:
        return f"Ошибка при анализе: {e}"
# 6. Low Homework Percentage
def analyze_low_homework_percentage(file_path: str):
    """
    Анализирует процент выполнения студентами выполнения домашних заданий

    Возвращает список студентов со Percentage Homework ниже 50%

    :param file_path: Путь к файлу .xlsx
    :return: Форматированная строка со списком студентов
    """
    try:
        df = pd.read_excel(file_path)

        # Get all need columns
        required_columns = ["FIO", "Percentage Homework"]
        if not all(col in df.columns for col in required_columns):
            return "Ошибка: в файле отсутствуют необходимые столбцы (FIO, Percentage Homework)"

        # Students list with low homework percentage
        low_percentage_students = []

        for _, row in df.iterrows():
            fio = row["FIO"]
            percentage = row["Percentage Homework"]

            if pd.notna(percentage) and isinstance(percentage, (int, float)) and percentage < 50:
                low_percentage_students.append(f"{fio}: {percentage:.0f}")
        
        # Make result
        if low_percentage_students:
            result = "Студенты с процентом выполненных ДЗ ниже 50%:\n" + "\n".join(low_percentage_students)
        else:
            result = "Все студенты имеют процент выполненных ДЗ 50% или выше "
        return result
    except Exception as e:
        return f"Ошибка при анализе: {e}"
# 7. Marks analysis
def analyze_bad_marks(file_path : str):
    """
    Анализирует оценки студентов за Homework и Classroom

    Возвращает список студентов со средней оценкой ниже 3

    :param file_path: Путь к файлу .xlsx
    :return: Форматированная строка со списком студентов
    """
    try:
        df = pd.read_excel(file_path)

        # Get all need columns
        required_columns = ["FIO", "Homework", "Classroom"]
        if not all(col in df.columns for col in required_columns):
            return "Ошибка: в файле отсутствуют необходимые столбцы (FIO, Homework, Classroom)"
        
        # Students list with bad marks
        bad_marks_students = []

        for _, row in df.iterrows():
            fio = row["FIO"]
            homework = row["Homework"]
            classroom = row["Classroom"]

            if pd.notna(homework) and pd.notna(classroom) and isinstance(homework, (int, float)) and isinstance(classroom, (int, float)):
                average_score = (homework + classroom) / 2
                if average_score < 3:
                    bad_marks_students.append(f"{fio}: {average_score:.1f}")

        # Make result
        if bad_marks_students:
            result = "Студенты с средней оценкой ниже 3:\n" + "\n".join(bad_marks_students)
        else:
            result = "Все студенты имеют среднюю оценку 3 или выше"
        return result
    except Exception as e:
        return f"Ошибка при анализе: {e}"