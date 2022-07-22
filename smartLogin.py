import json
import os
from time import sleep

import requests
from selenium.webdriver.common.by import By

from seleniumDriver import CreateEdgeDriverService


def SmartLogin_With(URL, username, password, TargetTitle):
    with CreateEdgeDriverService() as driver:
        SmartLogin(driver, URL, username, password, TargetTitle)
    return LoadCookies_requests()


def SmartLogin(driver, URL, username, password, TargetTitle):
    CookiesLogin(URL, driver)
    driver.maximize_window()
    while True:
        if '登录' in driver.title or '智慧校园' in driver.title:
            Login(URL, driver, username, password)
            SaveCookies(driver.get_cookies())
        elif driver.title == TargetTitle:
            break
        else:
            raise
    return LoadCookies_requests()


def Login(URL, driver, Username, Password):
    driver.get(URL)
    driver.implicitly_wait(5)
    while '登录' in driver.title:
        login = driver.find_element(By.ID, 'username')
        login.clear()
        login.send_keys(Username)
        password = driver.find_element(By.ID, 'password'
                                       )
        password.clear()
        password.send_keys(Password)
        driver.find_element(By.CLASS_NAME, 'el-checkbox__inner').click()
        driver.find_element(By.NAME, 'submit').click()
        sleep(5)


def CookiesLogin(URL, driver):
    driver.get(URL)
    cookies = LoadCookies()
    if cookies:
        for cookie in cookies:
            driver.add_cookie(cookie)
        driver.get(URL)
    else:
        driver.get(URL)
    # driver.implicitly_wait(3)


def SaveCookies(cookies):
    with open('cookies.json', 'w') as f:
        json.dump(cookies, f)


def LoadCookies():
    if os.path.exists('cookies.json'):
        with open('cookies.json', 'r') as f:
            cookies = json.load(f)
        return cookies
    else:
        return False


def LoadCookies_requests():
    cookies = LoadCookies()
    cookiesJar = requests.cookies.RequestsCookieJar()
    for cookie in cookies:
        cookiesJar.set(cookie['name'], cookie['value'],
                       domain=cookie['domain'], path=cookie['path'])
    # Dict Type
    # cookies = '; '.join(item for item in [
    #     item["name"] + "=" + item["value"] for item in LoadCookies()])
    return cookiesJar
