import msvcrt
import subprocess
from colorama import Fore, Back, Style, init
import pandas as pd
import camelot
from tabula import read_pdf
import pypdfium2 as pdfium
import jpype
import time
from sys import exit
import os
from pathlib import Path
from datetime import datetime
from questionary import select
from questionary import Style 

def pause():
    print("Нажмите любую клавишу для продолжения...")
    msvcrt.getch()

def ask_yes_no(question, flag_answer):
    while True:
        if flag_answer:
            response = input(question + " [Д/н]: ").strip().lower()  # Получаем ввод и приводим к нижнему регистру
            if response in ['y', 'yes', 'д', 'да', '']:  # Проверяем на "Y" или "yes"
                 return True
            elif response in ['n', 'no', 'н', 'нет']:  # Проверяем на "N" или "no", пустой ввод считается как "no"
                return False
        else:
            response = input(question + " [д/Н]: ").strip().lower()  # Получаем ввод и приводим к нижнему регистру
            if response in ['y', 'yes', 'д', 'да']:  # Проверяем на "Y" или "yes"
                return True
            elif response in ['n', 'no', 'н', 'нет', '']:  # Проверяем на "N" или "no", пустой ввод считается как "no"
                return False


# Инициализация colorama
init(autoreset=True)  # Это автоматически сбрасывает стиль после каждого print

print(Fore.GREEN + r"   _____       __    __                                                         __   ")
print(Fore.GREEN + r"  / ___/____  / /___/ /__  _____(_)___  ____ _(_)________  ____     _________  / /_  ")
print(Fore.GREEN + r"  \__ \/ __ \/ / __  / _ \/ ___/ / __ \/ __ `/ / ___/ __ \/ __ \   / ___/ __ \/ __ \ ")
print(Fore.GREEN + r" ___/ / /_/ / / /_/ /  __/ /  / / / / / /_/ / / /  / /_/ / / / /  (__  ) /_/ / /_/ / ")
print(Fore.GREEN + r"/____/\____/_/\__,_/\___/_/  /_/_/ /_/\__, /_/_/   \____/_/ /_/  /____/ .___/_.___/  ")
print(Fore.GREEN + r"                                     /____/             " + Fore.GREEN + r"  ___   ____ " + Fore.GREEN + r"/" + Fore.GREEN + r"_" + Fore.GREEN + r"/" + Fore.GREEN + r"   ______    ")
print(Fore.GREEN + r"                                                         |__ \ / __ \__ \ / ____/    ")
print(Fore.GREEN + r"                                                         __/ // / / /_/ //___ \      ")
print(Fore.GREEN + r"                                                        / __// /_/ / __/____/ /      ")
print(Fore.GREEN + r"                                                       /____/\____/____/_____/       ")

print(Fore.GREEN + r"[+] " + Fore.YELLOW + r"Name       : " + Fore.WHITE + r"Парсер *.pdf файлов")
print(Fore.GREEN + r"[+] " + Fore.YELLOW + r"Version    : " + Fore.WHITE + r"1.0.1")
print(Fore.GREEN + r"[+] " + Fore.YELLOW + r"Date       : " + Fore.WHITE + r"01.02.2025")
print(Fore.GREEN + r"[+] " + Fore.YELLOW + r"Created by : " + Fore.WHITE + r"Олег Волков")
print(Fore.GREEN + r"   ↪ " + Fore.YELLOW + r"VK        : " + Fore.WHITE + "https://vk.com/solderingiron.stm32")
print(Fore.GREEN + r"   ↪ " + Fore.YELLOW + r"GitHub    : " + Fore.WHITE + "https://github.com/Solderingironspb")
print(Fore.GREEN + r"   ↪ " + Fore.YELLOW + r"email     : " + Fore.WHITE + "solderingiron.info@yandex.ru")

time.sleep(1)

# Укажите путь к директории
directory_path = Path(r"PDF_files")

# Получаем список файлов с расширением *.pdf и их датой изменения
pdf_files = [(file.name, file.stat().st_mtime) for file in directory_path.glob("*.pdf")]

# Проверяем, есть ли файлы
if pdf_files:
    # Сортируем файлы по дате изменения (по убыванию, чтобы новые были вверху)
    pdf_files_sorted = sorted(pdf_files, key=lambda x: x[1], reverse=True)
    # Формируем список для выбора (только имена файлов)
    choices = [file for file, mtime in pdf_files_sorted]
    choices.append("Выход")
    # Выводим список файлов с датами
    print(Fore.GREEN + f"\n[+] " + Fore.WHITE + f"В директории {directory_path} найдено {len(pdf_files_sorted)} *.pdf файла(ов):")
    print(Fore.GREEN + "    " + "-" * 82)
    print(Fore.GREEN + "    {:<60} {:<20}".format("Имя файла", "| Дата изменения"))
    print(Fore.GREEN + "    " + "-" * 82)
    for file, mtime in pdf_files_sorted:
        mtime_readable = datetime.fromtimestamp(mtime).strftime('| %Y-%m-%d %H:%M:%S')
        print(Fore.GREEN + "    {:<60} {:<20}".format(file, mtime_readable))

    print(Fore.GREEN + "    " + "-" * 82 + '\n')
    
 
    # Создаем кастомный стиль
    custom_style = Style([
        ('selected', 'fg:#ffffff bold'),  
        ('pointer', 'fg:#ffff00 bold'),   
        ('question', 'fg:#ffffff bold'),  
        ('instruction', 'fg:#ffffff'),   
        ('answered', 'fg:#00ff00 bold'),  
    ])
    
    # Запрашиваем выбор пользователя
    selected = select(
        "Выберите файл, который требуется распарсить:",
        choices=choices,
        instruction=' ',    # Убираем подсказку (Use arrow keys)
        style=custom_style  # Применяем кастомный стиль
    ).ask()

    # Обрабатываем выбор
    if selected == "Выход":
        exit()
    else:
        # Имя файла уже извлечено
        print(Fore.GREEN + f"Вы выбрали файл: {selected}")
        # Здесь можно добавить логику для парсинга выбранного файла
else:
    print(f"В директории {directory_path} нет файлов с расширением *.pdf.")
    pause()
    exit()
 

file_pdf = selected #input('\nВведите название файла без типа:')

# Укажите путь к вашему PDF файлу
pdf_file_path = f'PDF_files/{file_pdf}'
output_xlsx_file_path = file_pdf.rsplit(".pdf", 1)[0]
# Укажите путь для сохранения Excel файла
excel_file_path = f'Excel_files/{output_xlsx_file_path}.xlsx'
parsing_method = 'steam'
parsing_method_tabula = False

 

# Запрашиваем выбор пользователя
selected = select(
    "Выберите библиотеку для парсинга:",
     choices=[
        "Camelot",
        "Tabula",
        "Выход"
        ],
    instruction=' ',  # Убираем подсказку (Use arrow keys)
    style=custom_style  # Применяем кастомный стиль
).ask()

method_lib = 0

# Обрабатываем выбор
if selected == "Выход":
    exit()
else:
    # Имя файла уже извлечено
    print(Fore.GREEN + f"Вы выбрали библиотеку: {selected}")
    # Здесь можно добавить логику для парсинга выбранного файла
    if selected == "Camelot":
        method_lib = 0 #Camelot
    elif selected == "Tabula":
        method_lib = 1 #Tabula    


# Запрашиваем выбор пользователя
selected = select(
    "Выберите метод парсинга:",
     choices=[
        "lattice",
        "stream",
        "Выход"
        ],
    instruction=' ',  # Убираем подсказку (Use arrow keys)
    style=custom_style  # Применяем кастомный стиль
).ask()

# Обрабатываем выбор
if selected == "Выход":
    exit()
else:
    # Имя файла уже извлечено
    print(Fore.GREEN + f"Вы выбрали метод: {selected}")
    # Здесь можно добавить логику для парсинга выбранного файла
    if selected == "lattice":
        #print(Fore.GREEN + 'Используем метод парсинга lattice')
        parsing_method = "lattice"
        parsing_method_tabula = True
    elif selected == "stream":
        #print(Fore.GREEN + 'Используем метод парсинга stream')
        parsing_method = "stream"
        parsing_method_tabula = False  


try:
    # Извлечение таблицы из PDF
    print(Fore.WHITE + r"Извлечение таблиц из PDF")    
    if method_lib == 0:
        #print(Fore.WHITE + r"Camelot: Извлечение таблиц из PDF")    
        tables = camelot.read_pdf(pdf_file_path, pages='all', flavor=parsing_method)
    else:
        #print(Fore.WHITE + r"Tabula: Извлечение таблиц из PDF")    
        tables = read_pdf(pdf_file_path, pages='all', multiple_tables=True, lattice=parsing_method_tabula)
    print(Fore.WHITE + r"Проверяем, были ли найдены таблицы")
    # Проверяем, были ли найдены таблицы
    if tables:
        print(Fore.WHITE + r"Создаем Excel файл и записываем таблицы")
        # Создаем Excel файл и записываем таблицы
        with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
            for i, table in enumerate(tables):
                print(Fore.WHITE + r"Записываем каждую распознанную таблицу на отдельный лист")
                # Записываем каждую таблицу на отдельный лист
                if method_lib == 0:
                    table.df.to_excel(writer, sheet_name=f'part_{i + 1}', index=False)
                else:
                    table.to_excel(writer, sheet_name=f'part_{i + 1}', index=False)
        print(Fore.WHITE +f"Таблицы успешно сохранены в {excel_file_path}")
        subprocess.Popen(['start', '', excel_file_path], shell=True)
        pause()
    else:
        print(Fore.RED + "Таблицы не найдены в PDF файле.")
        pause()
except Exception as e:
    print(Fore.RED + f"Ошибка при чтении PDF-файла: {e}")
    print(Fore.RED + "Попробуйте другую библиотеку...")
    pause()
