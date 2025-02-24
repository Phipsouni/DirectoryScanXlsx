import os
import re
import pandas as pd
import time
import sys

# Чтение путей из файла path.txt
with open('path.txt', 'r') as file:
    paths = file.readlines()
    source_path = paths[0].strip()  # Путь к анализируемой директории
    save_path = paths[1].strip()  # Путь к директории для сохранения файла

# Инициализация списка для хранения данных
data = []

# Регулярное выражение для извлечения номера папки
folder_number_pattern = re.compile(r'^\d+')

# Регулярное выражение для определения файлов ЭСД и GTD
esd_pattern = re.compile(r'^([\w-]+)\.pdf$', re.IGNORECASE)  # ЭСД: любой PDF файл
gtd_pattern = re.compile(r'^GTD_(\d+)_(\d+)_(\d+)\.pdf$', re.IGNORECASE)  # GTD: формат GTD_

# Проход по всем папкам в заданной директории
for folder_name in os.listdir(source_path):
    folder_path = os.path.join(source_path, folder_name)

    # Проверка на то, что это папка
    if os.path.isdir(folder_path):
        # Извлечение номера папки
        folder_number_match = folder_number_pattern.search(folder_name)
        if folder_number_match:
            folder_number = folder_number_match.group()  # Только номер папки
        else:
            continue  # Пропустить, если номер папки не найден

        esd_numbers = []  # Список для номеров ЭСД
        gtd_numbers = []  # Список для номеров GTD

        # Проход по всем файлам в папке
        for file_name in os.listdir(folder_path):
            # Проверка на ЭСД
            esd_match = esd_pattern.match(file_name)
            if esd_match and not file_name.startswith("GTD_"):  # Исключаем файлы GTD
                esd_numbers.append(esd_match.group(1))  # Добавляем имя файла без расширения
                continue

            # Проверка на GTD
            gtd_match = gtd_pattern.match(file_name)
            if gtd_match:
                gtd_formatted = f"{gtd_match.group(1)}/{gtd_match.group(2)}/{gtd_match.group(3)}"
                gtd_numbers.append(gtd_formatted)
                continue

        # Добавить данные в список
        data.append([
            int(folder_number),  # Конвертируем номер папки в int для сортировки
            ', '.join(esd_numbers) if esd_numbers else None,
            ', '.join(gtd_numbers) if gtd_numbers else None
        ])

# Создание DataFrame
df = pd.DataFrame(data, columns=["Folder Number", "ЭСД Number", "GTD Number"])

# Сортировка по возрастанию номера папки
df_sorted = df.sort_values('Folder Number')

# Сохранение в Excel файл
output_file_path = os.path.join(save_path, "ESD_DT.xlsx")
df_sorted.to_excel(output_file_path, index=False)

# Подсчёт количества уникальных GTD
dt_count = df['GTD Number'].apply(lambda x: len(x.split(', ')) if x else 0).sum()

# Печать информации
print("Номера ЭСД и ДТ выгружены")
print(f"Количество файлов GTD: {dt_count}")
print(f"Файл ESD_DT.xlsx был создан в директории: {save_path}")

# Ожидание 4 секунды перед завершением
time.sleep(4)
sys.exit()
