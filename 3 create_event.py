# # !/usr/bin/python
# # -*- coding: utf-8 -*-

import os
import shutil
import time
from tkinter import *

import pyautogui
from mysql import connector
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys

# # delete settings for vista
# if os.path.exists(r'C:\Users\Xe\AppData\Roaming\vista-med\connections.ini'):
#     os.remove(r'C:\Users\Xe\AppData\Roaming\vista-med\connections.ini')
# if os.path.exists(r'C:\Users\Xe\AppData\Roaming\vista-med\S11App.ini'):
#     os.remove(r'C:\Users\Xe\AppData\Roaming\vista-med\S11App.ini')
#
# # copy settings for vista
# shutil.copy(r'D:\VistaMed\settings\connections.ini', r'C:\Users\XE\AppData\Roaming\vista-med') # поправить путь
# shutil.copy(r'D:\VistaMed\settings\S11App.ini', r'C:\Users\XE\AppData\Roaming\vista-med')


# DATA-BASE SETTINGS
db = connector.connect(host='192.168.0.3', user='dbuser', passwd='dbpassword', database='p17_testspb', port='3306')
db_cursor = db.cursor()


def insert_stmt(stmt):
    db_cursor.execute(stmt)
    db.commit()
    return db_cursor.lastrowid


def select_stmt(stmt):
    db_cursor.execute(stmt)
    fetch_all = db_cursor.fetchall()
    return list(i for i in fetch_all)


# create client
def get_client_id():
    add_client_stmt = u"""
INSERT INTO Client(`createDatetime`, `createPerson_id`, `modifyDatetime`, `modifyPerson_id`, 
                   `firstName`, `lastName`, `patrName`,
                    `birthDate`, `birthTime`, `sex`, `SNILS`, `bloodNotes`, `growth`, `weight`,
                   `embryonalPeriodWeek`, `birthPlace`, `diagNames`, `chartBeginDate`, `notes`, 
                   `IIN`, `isUnconscious`, `chronicalMKB`)
VALUES (NOW(), 1, NOW(), 1, 
        'Тест', 'Тест', 'Тест',
        '1999-07-25', '00:00:00', 1, '11111111145', '', '0', '0',
        '0', 'СПБ', '', '2020-02-12', '', 
        '', 0, '')"""
    result = insert_stmt(add_client_stmt)
    return result


def get_document_type():
    document_type_stmt = u"""
SELECT rbDocumentType.id 
FROM rbDocumentType 
WHERE rbDocumentType.name LIKE '%паспорт%'
ORDER BY id
LIMIT 1"""
    result = select_stmt(document_type_stmt)
    return result[0][0]


def get_client_document(client_id):
    document_type = get_document_type()
    add_document_stmt = u"""
INSERT INTO ClientDocument(`createDatetime`, `createPerson_id`, `modifyDatetime`, `modifyPerson_id`, `deleted`, 
`client_id`, `documentType_id`, `serial`, `number`, `date`, `origin`) 
VALUES (NOW(), 1, NOW(), 1, 0, 
{client_id}, {document_type}, '1234', '123456', 'NOW()', 'УФМС России')""".format(
        client_id=client_id, document_type=document_type)
    result = insert_stmt(add_document_stmt)
    return result


def get_format_policy():
    policyTypeStmt = u"""
SELECT id
FROM rbPolicyType
WHERE rbPolicyType.name LIKE '%ОМС%'
ORDER BY id
LIMIT 1"""
    result = select_stmt(policyTypeStmt)
    return result[0][0]


def add_client_policy(client_id):
    policyType = get_format_policy()
    client_policy_stmt = u"""
INSERT INTO ClientPolicy(`createDatetime`, `createPerson_id`, `modifyDatetime`, `modifyPerson_id`, `deleted`,
                         `client_id`, `insurer_id`, policyType_id, `policyKind_id`, `serial`, `number`, `begDate`,
                         `endDate`, `name`, `note`, `insuranceArea`)
VALUES (NOW(), 1, NOW(), 1, 0, {client_id}, 3307, {policyType}, 3, 'ЕП', 
'111111', '2020-02-17', '2200-01-01', 'РОСНО', 'СПБ', '7800000000000')""".format(
        client_id=client_id, policyType=policyType, )
    result = insert_stmt(client_policy_stmt)
    return result


def add_client_contact(client_id):
    client_contact_stmt = u"""
INSERT INTO ClientContact(`createDatetime`, `createPerson_id`, `modifyDatetime`, `modifyPerson_id`, `deleted`,
                          `client_id`, `contactType_id`, `isPrimary`, `contact`, `notes`)
VALUES (NOW(), 1, NOW(), 1, 0,
        {client_id}, 3, 1, '5259471', 'тест')""".format(
        client_id=client_id)
    result = insert_stmt(client_contact_stmt)
    return result


def add_client_address(client_id, address_id):
    client_address_stmt = u"""
INSERT INTO ClientAddress(`createDatetime`, `createPerson_id`, `modifyDatetime`, `modifyPerson_id`, `client_id`, 
`type`, `address_id`, `district_id`, `isVillager`,`freeInput`)
VALUES (NOW(), 1, NOW(), 1, {client_id}, 0, {address_id}, 1, 0,'')""".format(
        client_id=client_id, address_id=address_id)
    result = insert_stmt(client_address_stmt)
    return result


def get_address_house_id():
    address_house_stmt = u"""
INSERT INTO AddressHouse(`createDatetime`, `createPerson_id`, `modifyDatetime`, `modifyPerson_id`, `KLADRCode`,
                         `KLADRStreetCode`, `number`, `corpus`)
VALUES (NOW(), 1, NOW(), 1, '7800000000000', 
        '78000000000227000', '1', '1')"""
    result = insert_stmt(address_house_stmt)
    return result


def get_address_id(address_house_id):
    add_address_stmt = u"""
INSERT INTO Address(`createDatetime`, `createPerson_id`, `modifyDatetime`, `modifyPerson_id`, `house_id`, `flat`)
VALUES (NOW(), 1, NOW(), 1, {address_house_id}, '1')""".format(
        address_house_id=address_house_id)
    result = insert_stmt(add_address_stmt)
    return result


# create patient:
client_id = get_client_id()
client_document = get_client_document(client_id)
client_policy = add_client_policy(client_id)
client_contact = add_client_contact(client_id)
address_house_id = get_address_house_id()
address_id = get_address_id(address_house_id)
client_address = add_client_address(client_id, address_id)
printed_client = u'Client.id = {client_id}'.format(client_id=client_id)
printed_event = u'Event.id = {event_id}'
print(printed_client)

driver = webdriver.Remote(
    command_executor='http://localhost:9999',
    desired_capabilities={
        "debugConnectToRunningApp": 'false',
        "app": r"D:\VistaMed\winclient\vista-med.exe"
    })

driver.implicitly_wait(15)

actions = ActionChains(driver)


# Найти имя родительского элемента
def findel_parent(name):
    something_response = driver.find_element_by_name(name)
    return something_response.parent


# Найти имя  элемента
def findel(name):
    something_response = driver.find_element_by_name(name)
    return something_response


# Нажатие на элемент внутри родителя
def findel_sub_click(sub_element, sub_name1):
    sub_element.find_element_by_name(sub_name1).click()


# Нажатие на элемент
def findel_click(name):
    driver.find_element_by_name(name).click()


# Перемещение курсора относительно текущей позиции
def move_cursor_click(x, y):
    pyautogui.moveRel(x, y)
    pyautogui.click()


def enter_words(fun, word):
    fun.send_keys(Keys.CONTROL + word)


try:
    # Подключение
    def into():
        findel_click('Сессия')
        findel_click('Подключиться к базе данных')


    into()


    # Вход
    def log_in():
        registration = findel('Регистрация')
        findel_sub_click(registration, 'Имя')
        move_cursor_click(None, -5)
        pyautogui.click()
        enter_words(registration, 'Виста')
        findel_sub_click(registration, 'ОК')


    log_in()


    # Обслуживание пациентов
    def patient_care():
        findel_click("Работа")
        findel_click("Обслуживание пациентов")


    patient_care()


finally:
    print('hello')
