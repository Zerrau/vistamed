import os
import sys
import time

import pyautogui
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

# Удаление нстроек и подключения
if os.path.exists(r'C:\Users\Xe\AppData\Roaming\vista-med\connections.ini'):
    os.remove(r'C:\Users\Xe\AppData\Roaming\vista-med\connections.ini')
if os.path.exists(r'C:\Users\Xe\AppData\Roaming\vista-med\S11App.ini'):
    os.remove(r'C:\Users\Xe\AppData\Roaming\vista-med\S11App.ini')

# Подключаемся к webdriver и запускаем Висту
driver = webdriver.Remote(
    command_executor='http://localhost:9999',
    desired_capabilities={
        "debugConnectToRunningApp": 'false',
        "app": r"D:\VistaMed\winclient\vista-med.exe"
    })

# неявное ожидания 
driver.implicitly_wait(15)


def find_by_name(name):
    response = driver.find_element_by_name(name)
    return response


def findel_click(name):
    driver.find_element_by_name(name).click()


def tap_tab():
    pyautogui.keyDown(key='tab')


def tap_enter():
    pyautogui.keyDown('enter')


# Закрытие окна не найденной бд
actions = ActionChains(driver)
# иногда окно не находится сразу
try:
    login_failed = driver.find_element_by_name("Ошибка открытия базы данных")
except exceptions.NoSuchElementException:
    while True:
        login_failed = driver.find_element_by_name("Ошибка открытия базы данных")

login_failed.find_element_by_name("Закрыть").click()

# Открытые настроек бд
findel_click("Настройки")
findel_click("База данных")

# Устновка настроек
settings_db = find_by_name('Настройки базы данных')
settings_db.find_element_by_name("Сервер").click()
tap_tab()
settings_db.send_keys('p17_testspb')
tap_tab()
tap_tab()
settings_db.send_keys('192.168.0.3')
tap_tab()
settings_db.send_keys('3306')
tap_tab()
settings_db.send_keys('p17_testspb')
tap_enter()

# Подключние
findel_click('Сессия')
findel_click('Подключиться к базе данных')

# Вход
login = find_by_name('Регистрация')
for i in range(4):
    tap_tab()
login.send_keys('Виста')
tap_enter()

# Проверка успешной авторизации
try:
    findel_click('Настройки')
    findel_click('Умолчания')
    check = 0
except exceptions.NoSuchElementException:
    print("Oops!. NoSuchElement Настройки or Умолчания")
    check = 1
assert check == 0, 'Test 1 Entry error'

# Закрытие Висты
os.system("TASKKILL /F /IM vista-med.exe")
