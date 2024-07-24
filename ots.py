import time
import schedule

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# options = webdriver.ChromeOptions()
# driver = webdriver.Chrome(options=options)

def get_xls_file_from_ots():
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=options)
    try:
        # добавить БД с учеткой локального компа и данными пользователя
        #
        # user_login = input("Введите ваш логин в OTS: ")
        # user_password = input("Введите ваш пароль в OTS: ")
        #

        url_1 = 'https://piellardj.github.io/totp-generator/?secret=3EJNNR2TXDVKXAD5&digits=6&period=30&algorithm=SHA-1'
        driver.get(url_1)

        wait_1 = WebDriverWait(driver, 10)
        time.sleep(3)

        element = wait_1.until(EC.presence_of_element_located((By.ID, 'generated-code')))
        element_value = element.text
        'This code will expire in 24'
        time_pin = wait_1.until(EC.presence_of_element_located((By.ID, 'seconds-left')))
        time_pin_value = time_pin.text

        print(element_value)

        if int(time_pin_value[25:27]) < 7:
            time.sleep(8)
            element_value = element.text

        print(element_value)

        user_login = "elizaveta.chagelova@autofinancebank.ru"
        user_password =

        url = 'https://ots.cft.ru/'
        # Откройте сайт
        driver.get(url)

        # Явное ожидание, чтобы кнопка стала кликабельной (поиск по ID, классу или другому селектору)
        wait = WebDriverWait(driver, 10)
        time.sleep(1)

        form_login = wait.until(EC.presence_of_element_located((By.ID, 'login-form-username')))
        form_login.send_keys(user_login)

        form_password = wait.until(EC.presence_of_element_located((By.ID, 'login-form-password')))
        form_password.send_keys(user_password)

        submit_button = wait.until(EC.presence_of_element_located((By.ID, 'login-form-submit')))  # Используем ID элемента
        submit_button.click()


        time.sleep(1)

        form_pin = wait.until(EC.presence_of_element_located((By.ID, '2fpin')))
        form_pin.send_keys(element_value)

        pin_button = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'aui-button')))
        pin_button.click()


        time.sleep(5)

        tasks_button = wait.until(EC.presence_of_element_located((By.ID, 'find_link')))
        tasks_button.click()
        time.sleep(3)

        open_tasks_button = wait.until(EC.presence_of_element_located((By.ID, 'filter_lnk_59401_lnk')))
        open_tasks_button.click()
        time.sleep(5)

        export_button = wait.until(EC.presence_of_element_located((By.ID, 'AJS_DROPDOWN__67')))
        export_button.click()
        time.sleep(5)

        excel_current_fields_button = wait.until(EC.presence_of_element_located((By.ID, 'currentExcelFields')))
        excel_current_fields_button.click()
        time.sleep(3)

    finally:
        # Закройте браузер
        driver.quit()
