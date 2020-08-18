# !/usr/bin/python
# -*- coding: utf-8 -*-

import os
import time
from tkinter import *

import pyautogui
from mysql import connector
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys

# Добавление пациента в базу
HOST = '192.168.0.3'
DB = 'p17_testspb'
PORT = '3306'
USER = 'dbuser'
PASSWORD = 'dbpassword'

# DB-CONNECTION:
db = connector.connect(host=HOST, user=USER, passwd=PASSWORD, database=DB, port=PORT)
db_cursor = db.cursor()

# Добавляем пациента
add_client = u"""
INSERT INTO Client(`createDatetime`, `createPerson_id`, `modifyDatetime`, `modifyPerson_id`, `lastName`, `firstName`,
                   `patrName`, `birthDate`, `birthTime`, `sex`, `SNILS`, `bloodNotes`, `growth`, `weight`,
                   `embryonalPeriodWeek`, `birthPlace`, `diagNames`, `chartBeginDate`, `notes`, `IIN`, `isUnconscious`,
                   chronicalMKB)
VALUES ('2020-02-12T18:29:40', 614, '2020-02-12T18:29:40', 614, 'Тест', 'Тест', 'Тест', '1999-07-25', '00:00:00',
        1,
        '11111111145', '', '0', '0', '0', '', '', '2020-02-12', '', '', 0, '')
        """

# Находим id добавленного пациента
get_client_id = u"""
select id
from Client
where lastName = 'Тест'
  and firstName = 'Тест'
  and patrName = 'Тест'
  and birthDate = '1999-07-25'
  and birthTime = '00:00:00'
  and sex = 1
  and SNILS = '11111111145'
  and deleted = 0
order by id desc
limit 1
"""
db_cursor.execute(add_client)  # Добавляется клиент
db.commit()  # Зафиксировать текущую транзакцию

client_id = db_cursor.execute(get_client_id)  # Получаем id клиента
client_id = db_cursor.fetchone()[0]  # Достаем id из кортежа

# Добавляем паспорт
add_ClientDocument = """
insert into ClientDocument (createDatetime, createPerson_id, modifyDatetime, modifyPerson_id, deleted, client_id,
                            documentType_id, serial, number, date, origin, endDate, version)
VALUES (now(), 614, now(), 614, 0, {}, 1, '11 11', '111111', now(),
        'СПБ МВД', null, null)""".format(client_id)

db_cursor.execute(add_ClientDocument)  # Добавляется клиент
db.commit()

# Добавляеться полис
add_ClientPolicy = """
INSERT INTO ClientPolicy (createDatetime, createPerson_id, modifyDatetime, modifyPerson_id, deleted, client_id,
                          insurer_id, policyType_id, policyKind_id, serial, number, begDate, endDate, dischargeDate,
                          name, note, insuranceArea, isSearchPolicy, franchisePercent)
VALUES (now(), 614, now(), 614, 0, {}, 3913, 1, 3, 'ЕП',
        '1111111111111111', now(), now(), null, 'Название', 'Примечание', '', 0, 0)""".format(client_id)
db_cursor.execute(add_ClientPolicy)  # Добавляется клиент
db.commit()

if os.path.exists(r'C:\Users\Xe\AppData\Roaming\vista-med\connections.ini'):
    os.remove(r'C:\Users\Xe\AppData\Roaming\vista-med\connections.ini')
if os.path.exists(r'C:\Users\Xe\AppData\Roaming\vista-med\S11App.ini'):
    os.remove(r'C:\Users\Xe\AppData\Roaming\vista-med\S11App.ini')

driver = webdriver.Remote(
    command_executor='http://localhost:9999',
    desired_capabilities={
        "debugConnectToRunningApp": 'false',
        "app": r"D:\VistaMed\winclient\vista-med.exe"
    })

driver.implicitly_wait(15)

try:
    actions = ActionChains(driver)
    login_failed = driver.find_element_by_name("Ошибка открытия базы данных")
    login_failed.find_element_by_name('Закрыть').click()

    # Установка настроек БД
    driver.find_element_by_name("Настройки").click()
    driver.find_element_by_name("База данных").click()

    driver.find_element_by_name('Сервер').click()
    pyautogui.moveRel(None, -5)
    pyautogui.click()
    database_settings = driver.find_element_by_name("Настройки базы данных")
    database_settings.send_keys(Keys.CONTROL + '192.168.0.3')

    driver.find_element_by_name('Сервер').click()
    pyautogui.moveRel(None, 5)
    pyautogui.click()
    database_settings.send_keys(Keys.CONTROL + '3306')

    driver.find_element_by_name('База').click()
    pyautogui.moveRel(None, -5)
    pyautogui.click()
    database_settings.send_keys(Keys.CONTROL + 'p17_testspb')
    database_settings.find_element_by_name('ОК').click()

    # Подключение
    driver.find_element_by_name('Сессия').click()
    driver.find_element_by_name('Подключиться к базе данных').click()

    # Вход
    registration = driver.find_element_by_name('Регистрация')
    registration.find_element_by_name('Имя').click()
    pyautogui.moveRel(None, -5)
    pyautogui.click()
    registration.send_keys(Keys.CONTROL + 'Виста')
    registration.find_element_by_name('ОК').click()

    # Установка умолчаний
    driver.find_element_by_name("Настройки").click()
    driver.find_element_by_name('Умолчания').click()

    driver.find_element_by_name("...").click()

    driver.find_element_by_name('Код ИНФИС').click()
    organization_choice = driver.find_element_by_name('Выбор организации ')
    pyautogui.moveRel(100, None)
    pyautogui.click()
    organization_choice.send_keys('о25')
    organization_choice.find_element_by_name('ОК').click()

    defaults = driver.find_element_by_name('Умолчания')
    defaults.find_element_by_name('ОК').click()

    # Переподключение
    driver.find_element_by_name('Сессия').click()
    driver.find_element_by_name('Сменить пользователя').click()
    driver.find_element_by_name('ОК').click()

    # Убрать график
    driver.find_element_by_name("Настройки").click()
    driver.find_element_by_name('График').click()
    driver.find_element_by_name("Настройки").click()
    driver.find_element_by_name('ЛУД').click()
    driver.find_element_by_name("Настройки").click()
    driver.find_element_by_name('Открытые обращения').click()

    # Вход в обслуживание пациентов
    driver.find_element_by_name('Работа').click()
    driver.find_element_by_name('Обслуживание пациентов').click()

    # Регистрация пациента
    pyautogui.keyDown(key='f9')

    registration_card = driver.find_element_by_name('Регистрационная карточка *')
    try:
        driver.find_element_by_name("Фамилия").click()  # реализованно для пропуска не обязательных полей
    except exceptions.NoSuchElementException:
        print("Oops!. Try again...")
    pyautogui.moveRel(100, None)  # Перемещение относительно текущей опзиции
    pyautogui.click()

    registration_card.send_keys('Тест')

    pyautogui.keyDown(key='tab')
    registration_card.send_keys('Тест')

    pyautogui.keyDown(key='tab')
    registration_card.send_keys('Тест')

    driver.find_element_by_name('Дата рождения').click()
    pyautogui.moveRel(50, None)
    pyautogui.click()
    registration_card.send_keys('25071999')

    # Пол
    reg_sex = driver.find_element_by_name('Пол')
    reg_sex.click()
    pyautogui.keyDown('down')
    pyautogui.keyDown('enter')

    driver.find_element_by_name('СНИЛС').click()
    pyautogui.moveRel(50, None)
    pyautogui.click()
    registration_card.send_keys('11111111145')

    # Заполнение Адреса проживания
    address_lives = driver.find_element_by_name("Адрес проживания")
    address_lives.find_element_by_name("Дом").click()
    pyautogui.moveRel(-50, None)
    pyautogui.click()
    registration_card.send_keys('1-й (Горелово) проезд')

    address_lives.find_element_by_name("Дом").click()
    pyautogui.moveRel(40, None)
    pyautogui.click()
    registration_card.send_keys('1')

    address_lives.find_element_by_name("Корп").click()
    pyautogui.moveRel(40, None)
    pyautogui.click()
    registration_card.send_keys('1')

    address_lives.find_element_by_name("Кв").click()
    pyautogui.moveRel(None, -25)
    pyautogui.click()
    driver.find_element_by_name("Адмиралтейский").click()

    actions = ActionChains(driver)  # MouseDoubleClick
    actions.double_click()
    actions.perform()

    registration_card.find_element_by_name("Кв").click()
    pyautogui.moveRel(40, None)
    pyautogui.click()
    registration_card.send_keys('1')

    # Подтянуть адресс регистрации
    reg_card_addres = driver.find_element_by_name("Адрес регистрации")
    reg_card_addres.find_element_by_name("...").click()

    # Ввод документа
    # Кем выдан
    driver.find_element_by_name("Документ").click()
    pyautogui.moveRel(None, 20)
    pyautogui.click()
    driver.find_element_by_name("ПАСПОРТ РФ").click()
    actions.double_click()
    actions.perform()
    # Серия
    address_lives.find_element_by_name("Кв").click()
    pyautogui.moveRel(30, None)
    pyautogui.click()
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='tab')
    registration_card.send_keys("11")
    pyautogui.keyDown(key='tab')
    registration_card.send_keys("11")
    # Номер
    pyautogui.keyDown(key='tab')
    registration_card.send_keys("111111")
    # Дата выдачи
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='space')
    # кем выдан
    pyautogui.keyDown(key='tab')
    registration_card.send_keys("СПБ МВД")

    # Полис ОМС
    # Тип Полиса ОМС
    pyautogui.keyDown(key='tab')
    registration_card.send_keys("3")
    # Серия
    pyautogui.keyDown(key='tab')
    # Номер
    pyautogui.keyDown(key='tab')
    registration_card.send_keys("1111111111111111")
    # дата с
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='space')
    # дата по
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='space')
    pyautogui.keyDown(key='tab')
    # СМО
    registration_card.send_keys("3")
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='tab')
    # Название
    registration_card.send_keys("Название")
    pyautogui.keyDown(key='tab')
    # Примечание
    registration_card.send_keys("Примечание")

    # Это все для Места рождения
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='tab')
    registration_card.send_keys("СПБ")

    # Контакты
    pyautogui.keyDown(key='tab')
    pyautogui.keyDown(key='tab')
    registration_card.send_keys("3")
    pyautogui.keyDown(key='tab')
    registration_card.send_keys("77777777777")

    # Сохранение обращения
    registration_card.find_element_by_name('ОК').click()

    # Двойник по дате рождения и СНИЛС
    driver.find_element_by_name('ОК').click()
    driver.find_element_by_name('ОК').click()
    driver.find_element_by_name('ОК').click()  # Доделать запрос
    driver.find_element_by_name('ОК').click()  # Доделать запрос

    # Получаем id созданного пациента
    pyautogui.keyDown(key='f4')
    driver.find_element_by_name("Код").click()
    pyautogui.moveRel(45, None)
    pyautogui.click()

    actions.double_click()
    actions.perform()
    pyautogui.rightClick()
    driver.find_element_by_name("Копировать	Ctrl+C").click()

    # Копирование из буфера id Client
    c = Tk()
    c.withdraw()
    copy_client_id = c.clipboard_get()
    c.update()
    c.destroy()

    # Сохраняем обращение
    driver.find_element_by_name('ОК').click()
    driver.find_element_by_name('ОК').click()
    driver.find_element_by_name('ОК').click()
    driver.find_element_by_name('ОК').click()
    driver.find_element_by_name('ОК').click()

    # Удаляем клиента
    deleted_client = u"""
     update Client
     set deleted=1
     where id = {}
     """.format(client_id)
    db_cursor.execute(deleted_client)
    db.commit()

    # Удаляем созданного клиента + проверка записи в бд
    deleted_copy_client = u"""
     update Client
     set deleted=1
     where id = {}
       and lastName = 'Тест'
       and firstName = 'Тест'
       and patrName = 'Тест'
       and modifyPerson_id = 614
       and birthDate = '1999-07-25'
       and SNILS = 11111111145
       and birthTime = '00:00:00'
       and deleted = 0
       and birthPlace = 'СПБ'
     """.format(copy_client_id)
    db_cursor.execute(deleted_copy_client)
    db.commit()

    deleted_ClientDocument = """update ClientDocument
    set deleted = 1
    where client_id = {}""".format(client_id)
    db_cursor.execute(deleted_ClientDocument)
    db.commit()

    deleted_ClientPolicy = """
    update ClientPolicy
    set deleted = 1
    where client_id = {}""".format(client_id)
    db_cursor.execute(deleted_ClientPolicy)
    db.commit()
finally:
    # подумать что должно происходить после подения теста
    print('hello')
    # subprocess
