import os
from pathlib import Path
from datetime import date

# Функция поиска файла по дате
def find_file_by_partial_name(directory = 'C:\\Users\\iz02074\\Downloads', partial_name = str(date.today())):
    for file_name in os.listdir(directory):
        if file_name.find(partial_name) != -1 and file_name.endswith('.xls'):
            return Path(directory) / f'{file_name}'
    raise FileNotFoundError(f"Файл, содержащий '{partial_name}' в названии, не найден.")

