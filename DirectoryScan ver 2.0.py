import os
import re
import pandas as pd
import time
import sys

# Чтение путей из файла path.txt
with open('path.txt', 'r') as file:
    paths = file.readlines()
    source_path = paths[0].strip()
    save_path = paths[1].strip()

# Регулярное выражение для извлечения номера папки и названия приложения
folder_number_pattern = re.compile(r'^(\d+)')
app_name_pattern = re.compile(r'^[^,]+, [^,]+, [^,]+, ([^,]+),')

# Регулярное выражение для определения файлов ЭСД и GTD
esd_pattern = re.compile(r'^([\w-]+)\.pdf$', re.IGNORECASE)
gtd_pattern = re.compile(r'^GTD_(\d+)_(\d+)_(\d+)\.pdf$', re.IGNORECASE)

# Список для хранения данных
data = []

# Обход папок в директории
for folder_name in os.listdir(source_path):
    folder_path = os.path.join(source_path, folder_name)

    if os.path.isdir(folder_path):
        folder_number_match = folder_number_pattern.search(folder_name)
        app_name_match = app_name_pattern.search(folder_name)

        if not folder_number_match or not app_name_match:
            continue

        folder_number = int(folder_number_match.group(1))
        app_name = app_name_match.group(1).strip()

        esd_numbers = []
        gtd_numbers = []

        for file_name in os.listdir(folder_path):
            # ЭСД: PDF с ровно 4 дефисами в названии (без расширения)
            name_without_ext = file_name[:-4] if file_name.lower().endswith('.pdf') else ''
            if (esd_pattern.match(file_name) and not file_name.startswith("GTD_")
                    and name_without_ext.count('-') == 4):
                esd_numbers.append(file_name[:-4])

            gtd_match = gtd_pattern.match(file_name)
            if gtd_match:
                gtd_numbers.append(f"{gtd_match.group(1)}/{gtd_match.group(2)}/{gtd_match.group(3)}")

        data.append([app_name, folder_number, ', '.join(esd_numbers), ', '.join(gtd_numbers)])

# Создание DataFrame
df = pd.DataFrame(data, columns=["Application", "Folder Number", "ЭСД Number", "GTD Number"])

# Сортировка по названию приложения и номеру папки
df_sorted = df.sort_values(by=["Application", "Folder Number"])

# Создание группированной таблицы
output_data = []
current_app = None

for _, row in df_sorted.iterrows():
    app_name, folder_number, esd, gtd = row

    if app_name != current_app:
        output_data.append([app_name, None, None, None])  # Добавляем заголовок приложения
        current_app = app_name

    output_data.append([None, folder_number, esd, gtd])

# Создание DataFrame и сохранение в Excel
df_final = pd.DataFrame(output_data, columns=["Application", "Folder Number", "ЭСД Number", "GTD Number"])
df_final.to_excel(os.path.join(save_path, "ESD_DT.xlsx"), index=False)

print("Файл успешно создан.")
time.sleep(4)
sys.exit()
