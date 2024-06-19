import os
import re
import pandas as pd
import time  # Для сна
import sys   # Для закрытия скрипта

# Чтение путей из файла path.txt
with open('path.txt', 'r') as file:
    paths = file.readlines()
    source_path = paths[0].strip()  # Путь к анализируемой директории
    save_path = paths[1].strip()  # Путь к директории для сохранения файла

# Инициализация списка для хранения данных
data = []

# Регулярное выражение для поиска номера папки
folder_number_pattern = re.compile(r'\d+')

# Проход по всем папкам в заданной директории
for folder_name in os.listdir(source_path):
    folder_path = os.path.join(source_path, folder_name)

    # Проверка на то, что это папка
    if os.path.isdir(folder_path):
        # Извлечение только номера папки
        folder_number_match = folder_number_pattern.search(folder_name)
        if folder_number_match:
            folder_number = folder_number_match.group()
        else:
            continue  # Пропустить, если номер папки не найден

        has_esd_or_gtd = False  # Флаг для проверки наличия файлов ЭСД или GTD

        # Проход по всем файлам в папке
        for file_name in os.listdir(folder_path):
            if file_name.startswith("ЭСД #"):
                esd_number = file_name.replace("ЭСД #", "").split('.')[0]  # Удаление "ЭСД #" и любых расширений файла
                data.append([folder_number, esd_number, None])
                has_esd_or_gtd = True
            elif file_name.startswith("GTD_"):
                gtd_number = file_name.replace("GTD_", "").replace("_", "/").split('.')[
                    0]  # Удаление "GTD_" и любого расширения файла
                data.append([folder_number, None, gtd_number])
                has_esd_or_gtd = True

        # Если файлов ЭСД и GTD не найдено, добавить запись с пустыми значениями
        if not has_esd_or_gtd:
            data.append([folder_number, None, None])

# Создание DataFrame
df = pd.DataFrame(data, columns=["Folder Number", "ЭСД Number", "GTD Number"])

# Группировка данных для объединения всех значений ЭСД и GTD для одной папки
df_grouped = df.groupby("Folder Number").agg({'ЭСД Number': lambda x: ', '.join(filter(None, x)),
                                              'GTD Number': lambda x: ', '.join(filter(None, x))}).reset_index()

# Преобразование 'Folder Number' в численный тип данных для сортировки и сортировка
df_grouped['Folder Number'] = df_grouped['Folder Number'].astype(int)
df_sorted = df_grouped.sort_values('Folder Number')

# Сохранение в Excel файл
output_file_path = os.path.join(save_path, "ESD_DT.xlsx")
df_sorted.to_excel(output_file_path, index=False)

# Вычислить количество уникальных файлов ДТ
dt_count = df['GTD Number'].dropna().nunique()

# Печать необходимых данных
print("Номера ЭСД и ДТ выгружены")
print(f"Количество файлов ДТ: {dt_count}")
print(f"Файл ESD_DT.xlsx был создан в директории: {save_path}")

# Ожидание 3 секунды перед закрытием
time.sleep(3)

# Закрытие скрипта
sys.exit()
