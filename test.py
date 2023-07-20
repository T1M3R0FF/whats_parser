import time
import pandas as pd
from pywinauto import Application
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import pywinauto.keyboard as keyboard
from tqdm.auto import tqdm


def excel():
    print('Добро пожаловать в WhatsApp парсер!\nДля работы требуется файл Excel.')
    time.sleep(3)
    excel_file = input('Введите название файла Excel без расширения: ') + '.xlsx'
    print('Принято')
    time.sleep(1)
    numbers = open('numbers.txt', 'w')
    print('Создаю файл found_in_Wapp.txt для вывода')
    found = open('found_in_Wapp.txt', 'w')
    data_frame = pd.read_excel(excel_file)
    data_frame_dict = data_frame.to_dict()
    print('Перезаписываю формат номеров')
    for key, _ in data_frame_dict.items():
        for num in data_frame_dict[key].values():
            num = str(num) + '\n'
            numbers.write(num)
    return found


def get_window_text(number, found):
    app = Application(backend="uia").connect(title_re="WhatsApp")
    main_window = app.window(title_re="WhatsApp")
    texts = main_window.descendants(control_type="Text")

    if texts[0].window_text() != 'Чаты':
        print(f'Номер не найден: +{number}')
        keyboard.send_keys('{ENTER}')
    else:
        print(f'Номер найден: +{number}')
        found.write(number)


def wapp_opener(found):
    time.sleep(1)
    print('Запускаю модуль парсинга...')
    time.sleep(1)
    print(3)
    time.sleep(1)
    print(2)
    time.sleep(1)
    print(1)
    driver = webdriver.Chrome()
    numbers = open('numbers.txt', 'r')
    batch_size = 20  # Размер пакета номеров для обработки
    number_batch = []  # Хранит номера текущего пакета

    with open('numbers.txt', 'r') as numbers_file:
        number_count = sum(1 for _ in numbers_file)
        for i, number in enumerate(numbers, 1):
            number_batch.append(number)
            if i % batch_size == 0 or i == number_count:
                process_batch(driver, number_batch, found)
                number_batch = []  # Очистить пакет номеров
            print(i)

    numbers.close()


def process_batch(driver, number_batch, found):
    driver.switch_to.window(driver.window_handles[0])
    for number in tqdm(number_batch):
        driver.get(f'https://wa.me/{number}')
        wait = WebDriverWait(driver, 10)
        button = wait.until(ec.visibility_of_element_located((By.ID, 'action-button')))
        button.click()
        time.sleep(0.3)  # Для того, чтобы успело прогрузиться main_window.descendants(control_type="Text")
        get_window_text(number, found)


def main():
    found = excel()
    wapp_opener(found)


main()
