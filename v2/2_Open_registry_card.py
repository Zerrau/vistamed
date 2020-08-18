import os
import shutil
import sys
import time

import pyautogui
from mysql import connector
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

# Удаление нстроек и подключения
if os.path.exists(r'C:\Users\Xe\AppData\Roaming\vista-med\connections.ini'):
    os.remove(r'C:\Users\Xe\AppData\Roaming\vista-med\connections.ini')
if os.path.exists(r'C:\Users\Xe\AppData\Roaming\vista-med\S11App.ini'):
    os.remove(r'C:\Users\Xe\AppData\Roaming\vista-med\S11App.ini')

# Копирование нстроек и подключений
original1 = r'D:\VistaMed\settings\connections.ini'
original2 = r'D:\VistaMed\settings\S11App.ini'

target1 = r'C:\Users\Xe\AppData\Roaming\vista-med\connections.ini'
target2 = r'C:\Users\Xe\AppData\Roaming\vista-med\S11App.ini'

shutil.copyfile(original1, target1)
shutil.copyfile(original2, target2)

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


# Вход
login = find_by_name('Регистрация')
for i in range(4):
    tap_tab()
login.send_keys('Виста')
tap_enter()

# Вход в обслуживание пациентов
findel_click('Работа')
findel_click('Обслуживание пациентов')

# Открытие Рег. карты и проверка
findel_click('Редактировать (F4)')
try:
    find_by_name('Регистрационная карточка ')
    check = 0
except exceptions.NoSuchElementException:
    print("Oops!. NoSuchElement Регистрационная карточка")
    check = 1
assert check == 0, 'Test 2 error finde registry window'

# Закрытие Висты
os.system("TASKKILL /F /IM vista-med.exe")

# Удаление из AppLock после закрытия висты
db = connector.connect(host='192.168.0.3', user='dbuser', passwd='dbpassword', database='p17_testspb', port='3306')
db_cursor = db.cursor()

add_client_stmt = u"""
delete
from AppLock
where addr = 'CENTRAL'"""
db_cursor.execute(add_client_stmt)
db.commit()
