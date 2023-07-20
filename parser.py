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
    print('WELCUM to the WhatsApp parser!\nДля работы понадобится файл Excel.')
    time.sleep(3)
    excel_file = input('Введите название файла excel без расширения: ') + '.xlsx'
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
        print(f'didnt find number: +{number}')
        keyboard.send_keys('{ENTER}')
    else:
        print(f'found number: +{number}')
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
    i = 1
    for number in tqdm(numbers):
        driver.get(f'https://wa.me/{number}')
        wait = WebDriverWait(driver, 10)
        button = wait.until(ec.visibility_of_element_located((By.ID, 'action-button')))
        button.click()
        time.sleep(0.3)  # чтобы успел прогрузиться main_window.descendants(control_type="Text")
        get_window_text(number, found)
        driver.switch_to.window(driver.window_handles[0])
        print(i)
        i += 1
    numbers.close()


def main():
    found = excel()
    wapp_opener(found)


main()
