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


    # Не найдена база
    def error_open_base():
        asdasd = driver.find_element_by_name("Ошибка открытия базы данных")
        asdasd.find_element_by_name("Закрыть").click()
        # main_window = findel_parent('Ошибка открытия базы данных')
        # findel_sub_click(main_window, 'Закрыть')


    try:
        error_open_base()  # обрабатываеться исключение в функции
    except Exception as e:
        pass
        print('Не найденно окно, Ошибка открытия базы данных')


    # Установка настроек БД
    def set_setings_database():
        findel_click('Настройки')
        findel_click('База данных')
        findel_click('Сервер')
        move_cursor_click(None, -5)
        database_settings = findel("Настройки базы данных")
        enter_words(database_settings, '192.168.0.3')


    try:
        set_setings_database()
    except Exception as e:
        pass
        print('Не удалось настроить подключение к БД')

    # Установка настроек БД
    into = findel("Настройки базы данных")
    into.send_keys(Keys.CONTROL + '192.168.0.3')
    pyautogui.keyDown(key='tab')
    into.send_keys(Keys.CONTROL + '3306')
    pyautogui.keyDown(key='tab')
    into.send_keys(Keys.CONTROL + 'p17_testspb')
    findel_click('ОК')

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


    # Открытие Работы
    # Обслуживание пациентов
    def patient_care():
        driver.find_element_by_name("Работа").click()
        driver.find_element_by_name("Обслуживание пациентов").click()
        driver.find_element_by_name("Регистрация (F9)")


    try:
        patient_care()
    except Exception as e:
        pass
        print('Не удалось открыть Обслуживание пациентов')


    def opening_time_tracking():
        # Открытие Учета рабочего времени
        driver.find_element_by_name("Работа").click()
        driver.find_element_by_name("Учёт рабочего времени").click()
        driver.find_element_by_name("График ")
        pyautogui.keyDown('esc')


    try:
        opening_time_tracking()
    except Exception as e:
        pass
        print('Не удалось открыть Учет рабочего времени')


    def deferred_log():
        # Журнал отложенного спроса
        driver.find_element_by_name("Работа").click()
        driver.find_element_by_name("Журнал отложенного спроса").click()  # пдумать как определить что ЖОС открылся


    try:
        deferred_log()
    except Exception as e:
        pass
        print('Не удалось открыть Журнал отложенного спроса')


    def home_doctor_call_log():
        # Журнал вызова врача на дом
        driver.find_element_by_name("Работа").click()
        driver.find_element_by_name(
            "Журнал вызова врача на дом").click()  # пдумать как определить что Журнал вызова врача на дом открылся


    try:
        home_doctor_call_log()
    except Exception as e:
        pass
        print('Не удалось открыть Журнал вызова врача на дом')


    def treatment_room_journal():
        # Журнал процедурного кабинета
        driver.find_element_by_name("Работа").click()
        driver.find_element_by_name("Журнал процедурного кабинета").click()
        driver.find_element_by_name("Журнал процедурного кабинета ")
        pyautogui.keyDown('esc')


    try:
        treatment_room_journal()
    except Exception as e:
        pass
        print('Не удалось открыть Журнал процедурного кабинета')


    def form_accounting():
        # Учёт бланков
        driver.find_element_by_name("Работа").click()
        driver.find_element_by_name("Учёт бланков").click()
        driver.find_element_by_name("Учет бланков ")
        pyautogui.keyDown('esc')


    try:
        form_accounting()
    except Exception as e:
        pass
        print('Не удалось открыть Учёт бланков')


    def stationary_monitor():
        # Стационарный монитор
        driver.find_element_by_name("Работа").click()
        driver.find_element_by_name("Стационарный монитор").click()
        stationary_monitor = driver.find_element_by_name("Стационарный монитор ")
        stationary_monitor.find_element_by_name("Закрыть").click()


    try:
        stationary_monitor()
    except Exception as e:
        pass
        print('Не удалось открыть Стационарный монитор')


    def work_planning():
        # Планирование работ
        driver.find_element_by_name("Работа").click()
        driver.find_element_by_name("Планирование работ").click()
        driver.find_element_by_name("Планирование ресурсов ")
        pyautogui.keyDown('esc')


    try:
        work_planning()
    except Exception as e:
        pass
        print('Не удалось открыть Планирование работ')


    def performance_of_work():
        # Выполнение работ
        driver.find_element_by_name("Работа").click()
        driver.find_element_by_name("Выполнение работ").click()
        driver.find_element_by_name("Выполнение работ ")
        pyautogui.keyDown('esc')


    try:
        performance_of_work()
    except Exception as e:
        pass
        print('Не удалось открыть Выполнение работ')


    def inventory_control():
        # Складской учёт
        driver.find_element_by_name("Работа").click()
        driver.find_element_by_name("Складской учёт").click()
        driver.find_element_by_name("Склад ЛСиИМН ")
        pyautogui.keyDown('esc')


    try:
        inventory_control()
    except Exception as e:
        pass
        print('Не удалось открыть Складской учёт')


    def quotas():
        # Квотирование
        driver.find_element_by_name("Работа").click()
        driver.find_element_by_name("Квотирование").click()
        driver.find_element_by_name("Квотирование ")
        pyautogui.keyDown('esc')


    try:
        quotas()
    except Exception as e:
        pass
        print('Не удалось открыть Квотирование')


    def laboratory_journal():
        # Лабораторный журнал
        driver.find_element_by_name("Работа").click()
        driver.find_element_by_name("Лабораторный журнал").click()
        driver.find_element_by_name("Лабораторный журнал ")
        pyautogui.keyDown('esc')


    try:
        laboratory_journal()
    except Exception as e:
        pass
        print('Не удалось открыть Лабораторный журнал')


    def warehouse_pharmacy():
        # Склад-Аптека
        driver.find_element_by_name("Работа").click()
        driver.find_element_by_name("Склад-Аптека").click()
        warehouse_pharmacy = driver.find_element_by_name("Склад-Аптека ")
        warehouse_pharmacy.find_element_by_name("Закрыть").click()  # подумать над реализацией через except


    try:
        warehouse_pharmacy()
    except Exception as e:
        pass
        print('Не удалось открыть Склад-Аптека')


    def accounts():
        # Расчет
        # Счета
        driver.find_element_by_name("Расчёт").click()
        driver.find_element_by_name("Счета").click()
        driver.find_element_by_name("Расчеты ")
        pyautogui.keyDown('esc')


    try:
        accounts()
    except Exception as e:
        pass
        print('Не удалось открыть Счета')


    def cash_register():
        # Журнал кассовых операций
        driver.find_element_by_name("Расчёт").click()
        driver.find_element_by_name("Журнал кассовых операций").click()
        driver.find_element_by_name("Журнал кассовых операций ")
        pyautogui.keyDown('esc')


    try:
        cash_register()
    except Exception as e:
        pass
        print('Не удалось открыть Журнал кассовых операций')


    def treaties():
        # Договоры
        driver.find_element_by_name("Расчёт").click()
        driver.find_element_by_name("Договоры").click()
        driver.find_element_by_name("Договоры ")
        pyautogui.keyDown('esc')


    try:
        treaties()
    except Exception as e:
        pass
        print('Не удалось открыть Договоры')


    def work_with_the_cash_register():
        # Работа с кассовым аппаратом
        driver.find_element_by_name("Расчёт").click()
        driver.find_element_by_name("Работа с кассовым аппаратом").click()
        driver.find_element_by_name("Менеджер ККМ")
        pyautogui.keyDown('esc')


    try:
        work_with_the_cash_register()
    except Exception as e:
        pass
        print('Не удалось открыть Работа с кассовым аппаратом')


    def import_population_list():
        # Обмен
        # Испорт
        # Импорт списка населения
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт списка населения").click()
        driver.find_element_by_name("Загрузка населения (ДД) из ТФОМС")
        pyautogui.keyDown('esc')


    try:
        import_population_list()
    except Exception as e:
        pass
        print('Не удалось открыть Импорт списка населения')


    def import_form_131_from_dbf():
        # Импорт формы 131
        # Импорт формы 131 из dbf
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт формы 131").click()
        driver.find_element_by_name("Импорт формы 131 из dbf").click()
        driver.find_element_by_name("Импорт реестра форм 131")
        pyautogui.keyDown('esc')


    try:
        import_form_131_from_dbf()
    except Exception as e:
        pass
        print('Не удалось открыть Импорт формы 131 из dbf')


    def import_an_XML_forms_registry_131():
        # Импорт реестра форм 131 XML
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт формы 131").click()
        driver.find_element_by_name("Импорт формы 131 из XML").click()
        driver.find_element_by_name("Импорт реестра форм 131 XML")
        pyautogui.keyDown('esc')


    try:
        import_an_XML_forms_registry_131()
    except Exception as e:
        pass
        print('Не удалось открыть Импорт реестра форм 131 XML')


    def import_error_log():
        # Импорт протокола ошибок
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт формы 131").click()
        driver.find_element_by_name("Импорт протокола ошибок").click()
        driver.find_element_by_name("Импорт протокола ошибок")
        pyautogui.keyDown('esc')


    try:
        import_error_log()
    except Exception as e:
        pass
        print('Не удалось открыть Импорт протокола ошибок')


    def from_INFIS():
        # Импорт профелей
        # из ИНФИС
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт профилей").click()
        driver.find_element_by_name("из ИНФИС").click()
        driver.find_element_by_name("Импорт профилей из ИНФИС")
        pyautogui.keyDown('esc')


    try:
        from_INFIS()
    except Exception as e:
        pass
        print('Не удалось открыть из ИНФИС')


    def import_beneficiaries():
        # Импорт льготников
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт льготников").click()
        driver.find_element_by_name("Импорт льготников")
        pyautogui.keyDown('esc')


    try:
        import_beneficiaries()
    except Exception as e:
        pass
        print('Не удалось открыть Импорт льготников')


    def import_quotas_from_VTMP():
        # Импорт квот из ВТМП
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт квот из ВТМП").click()
        driver.find_element_by_name("Загрузка квот из ВТМП").click()
        pyautogui.keyDown('esc')


    try:
        import_quotas_from_VTMP()
    except Exception as e:
        pass
        print('Не удалось открыть Импорт квот из ВТМП')


    def import_from_DBF_with_full_data_loading():
        # Импорт из DBF с полной загрузкой данных
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт из DBF с полной загрузкой данных").click()
        driver.find_element_by_name("Dialog").click()
        pyautogui.keyDown('esc')


    try:
        import_from_DBF_with_full_data_loading()
    except Exception as e:
        pass
        print('Не удалось открыть Импорт из DBF с полной загрузкой данных')


    def import_doctors():
        # Импорт врачей
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт врачей").click()
        driver.find_element_by_name("Импорт врачей").click()
        pyautogui.keyDown('esc')


    try:
        import_doctors()
    except Exception as e:
        pass
        print('Не удалось открыть Импорт врачей')


    def import_customers():
        # Импорт клиентов
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт клиентов").click()
        driver.find_element_by_name("Импорт пациентов").click()
        pyautogui.keyDown('esc')


    try:
        import_customers()
    except Exception as e:
        pass
        print('Не удалось открыть Импорт клиентов')


    def organizations_of_LPY_from_INFIS():
        # Импорт справочников
        # Организации ЛПУ из ИНФИС
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Организации ЛПУ из ИНФИС").click()
        driver.find_element_by_name("Импорт данных об ЛПУ из справочника ИНФИС")
        pyautogui.keyDown('esc')


    try:
        organizations_of_LPY_from_INFIS()
    except Exception as e:
        pass
        print('Не удалось открыть Организации ЛПУ из ИНФИС')


    def organizations_of_INFIS():
        # Организации из ЛПУ
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Организации из ЛПУ").click()
        driver.find_element_by_name("Импорт справочников.")
        pyautogui.keyDown('esc')


    try:
        organizations_of_INFIS()
    except Exception as e:
        pass
        print('Не удалось открыть Организации из ЛПУ')


    def action_types():
        # Типы действий
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Типы действий").click()
        driver.find_element_by_name("Импорт типов действий")
        pyautogui.keyDown('esc')


    try:
        action_types()
    except Exception as e:
        pass
        print('Не удалось открыть Типы действий')


    def event_types():
        # Типы событий
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Типы событий").click()
        driver.find_element_by_name("Импорт типов событий")
        pyautogui.keyDown('esc')


    try:
        event_types()
    except Exception as e:
        pass
        print('Не удалось открыть Типы событий')


    def action_templates():
        # Шаблоны действий
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Шаблоны действий").click()
        driver.find_element_by_name("Импорт шаблонов действий")
        pyautogui.keyDown('esc')


    try:
        action_templates()
    except Exception as e:
        pass
        print('Не удалось открыть Шаблоны действий')


    def event_results():
        # Результаты события
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Результаты события").click()
        driver.find_element_by_name('Импорт справочника "Результаты События"')
        pyautogui.keyDown('esc')


    try:
        event_results()
    except Exception as e:
        pass
        print('Не удалось открыть Результаты события')


    def units():
        # Единицы измерения
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Единицы измерения").click()
        driver.find_element_by_name("Отмена").click()


    try:
        units()
    except Exception as e:
        pass
        print('Не удалось открыть Единицы измерения')


    def thesaurus():
        # Тезаурус
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Тезаурус").click()
        driver.find_element_by_name("Импорт Тезауруса")
        pyautogui.keyDown('esc')


    try:
        thesaurus()
    except Exception as e:
        pass
        print('Не удалось открыть Тезаурус')


    def complaints():
        # Жалобы
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Жалобы").click()
        driver.find_element_by_name("Закрыть").click()


    try:
        complaints()
    except Exception as e:
        pass
        print('Не удалось открыть Жалобы')


    def services():
        # Услуги
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Услуги").click()
        driver.find_element_by_name("Закрыть").click()


    try:
        services()
    except Exception as e:
        pass
        print('Не удалось открыть Услуги')


    def services_from_the_nomenclature():
        # Услуги из номенклатуры
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Услуги из номенклатуры").click()
        driver.find_element_by_name("Загрузка услуг из номенкулатуры")
        pyautogui.keyDown('esc')


    try:
        services_from_the_nomenclature()
    except Exception as e:
        pass
        print('Не удалось открыть Услуги из номенклатуры')


    def services_from_the_MES_nomenclature():
        # Услуги из номенклатуры МЭС
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Услуги из номенклатуры МЭС").click()
        driver.find_element_by_name("Загрузка услуг из МЭС")
        pyautogui.keyDown('esc')


    try:
        services_from_the_MES_nomenclature()
    except Exception as e:
        pass
        print('Не удалось открыть Услуги из номенклатуры МЭС')


    def OMS_policy_series():
        # Серии полисов ОМС
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Серии полисов ОМС").click()
        driver.find_element_by_name("Загрузка серий полисов из DBF")
        pyautogui.keyDown('esc')


    try:
        OMS_policy_series()
    except Exception as e:
        pass
        print('Не удалось открыть Серии полисов ОМС')


    def doctors_vista_med():
        # Врачи (Виста-Мед)
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Врачи (Виста-Мед)").click()
        driver.find_element_by_name("Импорт справочников.")
        pyautogui.keyDown('esc')


    try:
        doctors_vista_med()
    except Exception as e:
        pass
        print('Не удалось открыть Врачи (Виста-Мед)')


    def diagnosis_results():
        # Результаты диагнозов
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Результаты диагнозов").click()
        driver.find_element_by_name("Импорт справочников.")
        pyautogui.keyDown('esc')


    try:
        diagnosis_results()
    except Exception as e:
        pass
        print('Не удалось открыть Результаты диагнозов')


    def directory_import_wizard():
        # Мастер импорта справочников
        driver.find_element_by_name("Обмен").click()
        driver.find_element_by_name("Импорт").click()
        driver.find_element_by_name("Импорт справочников").click()
        driver.find_element_by_name("Мастер импорта справочников").click()
        driver.find_element_by_name("Импорт справочников.")
        pyautogui.keyDown('esc')


    try:
        directory_import_wizard()
    except Exception as e:
        pass
        print('Не удалось открыть Мастер импорта справочников')


    def import_SMP():
        # Импорт СМП
        findel_click("Обмен")
        findel_click("Импорт")
        findel_click("Импорт СМП")
        findel("Импорт СМП")
        pyautogui.keyDown('esc')


    try:
        import_SMP()
    except Exception as e:
        pass
        print('Не удалось открыть Импорт СМП')


    def import_PACS():
        # Импорт ПАКС
        findel_click("Обмен")
        findel_click("Импорт")
        findel_click("Импорт ПАКС")
        findel("Импорт снимков в ПАКС-хранилище ")
        pyautogui.keyDown('esc')


    try:
        import_PACS()
    except Exception as e:
        pass
        print('Не удалось открыть Импорт ПАКС')


    def import_hospitalization_list():
        # Импорт списка госпитализации
        findel_click("Обмен")
        findel_click("Импорт")
        findel_click("Импорт списка госпитализации")
        findel("Загрузка сведений о госпитализации из ЕИС")
        pyautogui.keyDown('esc')


    try:
        import_hospitalization_list()
    except Exception as e:
        pass
        print('Не удалось открыть Импорт списка госпитализации')


    def export_forms_131():
        # Экспорт форм 131
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт форм 131")
        findel("Экспорт формы 131")
        pyautogui.keyDown('esc')


    try:
        export_forms_131()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт форм 131')


    def HL7_v2_51():
        # HL7 v2.5
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт первичных документов")
        findel_click("HL7 v2.5")
        findel("Экспорт сообщения о визите пациента в формате ISO/HL7")
        pyautogui.keyDown('esc')


    try:
        HL7_v2_51()
    except Exception as e:
        pass
        print('Не удалось открыть HL7_v2_51')


    def DBF():
        # DBF
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт первичных документов")
        findel_click("DBF")
        findel("Экспорт первичных документов ")
        pyautogui.keyDown('esc')


    try:
        DBF()
    except Exception as e:
        pass
        print('Не удалось открыть DBF')


    def XML():
        # XML
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт первичных документов")
        findel_click("XML")
        findel("Экспорт первичных документов в XML")
        pyautogui.keyDown('esc')


    try:
        XML()
    except Exception as e:
        pass
        print('Не удалось открыть XML')


    def export_action_results():
        # Экспорт результатов действий
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт результатов действий")
        findel("Экспорт результатов действий")
        pyautogui.keyDown('esc')


    try:
        export_action_results()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт результатов действий')


    def export_action_types():
        # Экспорт результатов действий
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт справочников")
        findel_click("Типы действий")
        findel("Экспорт типов действий")
        pyautogui.keyDown('esc')


    try:
        export_action_types()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт Типы действий')


    def export_event_type():
        # Типы событий
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт справочников")
        findel_click("Типы событий")
        findel("Экспорт типов событий")
        pyautogui.keyDown('esc')


    try:
        export_event_type()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт Типы событий')


    def action_templates():
        # Шаблоны действий
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт справочников")
        findel_click("Шаблоны действий")
        findel("Экспорт шаблонов свойств действий")
        pyautogui.keyDown('esc')


    try:
        action_templates()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт Шаблоны действий')


    def action_templates():
        # Шаблоны действий
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт справочников")
        findel_click("Шаблоны действий")
        findel("Экспорт шаблонов свойств действий")
        pyautogui.keyDown('esc')


    try:
        action_templates()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт Шаблоны действий')


    def event_results():
        # Результаты события
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт справочников")
        findel_click("Результаты события")
        findel('Экспорт справочника "Результаты События"')
        pyautogui.keyDown('esc')


    try:
        event_results()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт Результаты события')


    def export_units():
        # Единицы измерения
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт справочников")
        findel_click("Единицы измерения")
        findel('Экспорт справочника "Единицы измерения"')
        pyautogui.keyDown('esc')


    try:
        export_units()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт Единицы измерения')


    def export_thesaurus():
        # Тезаурус
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт справочников")
        findel_click("Тезаурус")
        findel('Экспорт Тезауруса')
        pyautogui.keyDown('esc')


    try:
        export_thesaurus()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт Тезаурус')


    def complaints():
        # Жалобы
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт справочников")
        findel_click("Жалобы")
        findel('Экспорт справочника "Жалобы"')
        pyautogui.keyDown('esc')


    try:
        complaints()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт Жалобы')


    def services():
        # Услуги
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт справочников")
        findel_click("Услуги")
        findel('Экспорт справочника "Услуги"')
        pyautogui.keyDown('esc')


    try:
        services()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт Услуги')


    def organization():
        # Организации
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт справочников")
        findel_click("Организации")
        findel('Мастер экспорта справочника организаций')
        pyautogui.keyDown('esc')


    try:
        organization()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт Организации')


    def doctors():
        # Врачи
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт справочников")
        findel_click("Врачи")
        findel('Мастер экспорта справочника врачей')
        pyautogui.keyDown('esc')


    try:
        doctors()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт Врачи')


    def diagnosis_results():
        # Результаты диагнозов
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт справочников")
        findel_click("Результаты диагнозов")
        findel('Мастер экспорта справочника результатов диагнозов')
        pyautogui.keyDown('esc')


    try:
        diagnosis_results()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт Результаты диагнозов')


    def diagnosis_results():
        # Результаты диагнозов
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт справочников")
        findel_click("Мастер экспорта справочников")
        findel('Мастер экспорта справочников')
        pyautogui.keyDown('esc')


    try:
        diagnosis_results()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт Мастер экспорта справочников')

        ##########################


    def IEMK_Export():
        # Экспорт ИЭМК
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт ИЭМК")
        findel('Отправка ИЭМК ')
        pyautogui.keyDown('esc')


    try:
        IEMK_Export()
    except Exception as e:
        pass
        print('Не удалось открыть  Экспорт ИЭМК')


    def export_to_FRTB():
        # Экспорт в ФРТБ
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт в ФРТБ")
        findel('Экспорт в Федеральный регистр туберкулезных больных ')
        pyautogui.keyDown('esc')


    try:
        export_to_FRTB()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт в ФРТБ')


    def export_operations():
        # Экспорт операций
        findel_click("Обмен")
        findel_click("Экспорт")
        findel_click("Экспорт операций")
        findel('Экспорт в DBF ')
        pyautogui.keyDown('esc')


    try:
        export_operations()
    except Exception as e:
        pass
        print('Не удалось открыть Экспорт операций')


finally:
    # pyautogui.hotkey('alt', 'f4') подумать как закрыть програму
    # или через крестик
    print('Hello')

# Придумать метод распознования
"""
 try:
        vrachi()
    except EXCEPTION as e:
        pass
        print(e)

    def findel(name):
        driver.find_element_by_name(name).click()


    findel('Настройки')

    sprav_list = ["Обмен", "Импорт", "Импорт справочников"]

    for sprav in sprav_list:
        findel(sprav)

    try:
        vrachi()
    except EXCEPTION as e:
        pass
        print(e)


            # im1 = pyautogui.screenshot()
            # im1.save('my_screenshot.png')
"""
