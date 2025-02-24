# DirectoryScanXlsx
Автоматизация записи номеров ЭСД и ДТ в xlsx документ.
Что делает скрипт:
1. Производит анализ папок с pdf-файлами в директории, указанной в файле path.txt
2. Создает таблицу Excel по пути указанному в файле path.txt. В Excel-таблице формирует столбцы с данными названия приложения (которое указывается между 3 и 4-ми запятыми в названии папки с инвойсом), номера ЭСД и номера ГТД.

1. При самом первом запуске скрипта, необходимо запустить Start.bat, который установит все необходимые библиотеки из файла "requirements.txt" для работы скрипта. это необходимо сделать только в первый раз.
2. Затем создать файл "path.txt", в котором необходимо в 1 строке указать путь к директории, где будет проводиться сбор данных, а в 2 строке путь к директории, куда будет сохраняться файл "ESD_DT" с собранными данными. Пример:

C:/Files/1
C:/Files/2

3. Запустить файл main.py, подождать успешного формирования Excel-файла, пройти в директорию сохранения файла (см. 2 пункт) и открыть файл "ESD_DT.xlsx".

Даже если в папке не будет файлов с "номером ЭСД" и "номером ГТД", то он все равно запишет данные, так как в директории есть папка с числом в начале. 
Если у папки числа нету, то в таком случае она записываться в Excel-документ не будет. Точно также не попадет любой pdf-файл, который не содержит в названии "-" или "_".
