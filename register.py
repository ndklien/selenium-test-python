from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

import openpyxl

path = "./White box Testing.xlsx"

lab2 = openpyxl.load_workbook(path)
register_test = lab2["Function1"]
columns = register_test.max_column

exception_case = [45, 46]

# Tạo ra danh sách các test case
test_case = []
for c in range(6, columns):
    unit = []
    for r in range(13, 47):
        checked = register_test.cell(row=r, column=c).value
        if checked == "O":
            # print("row number ", r)
            unit.append(r)

    if len(unit) == 0:
        pass
    else:
        test_case.append(unit)

# Config Driver
driver = webdriver.Chrome(executable_path=r"D:\UNIVERSITY\Nam_3\Nam3_2\SE113\chromedriver_win32\chromedriver.exe")
driver.maximize_window()
driver.get("http://127.0.0.1:8000/register/")

# Automation Test Cases
i = 0
untested_cases = 0
failed = 0
succeed = 0
for test in test_case:
    i += 1
    if test[7] in exception_case:
        untested_cases += 1
        continue
    else:
        username = driver.find_element_by_id("id_username")
        username.clear()
        # username.send_keys(test[3])
        username.send_keys(register_test.cell(row=test[3], column=4).value)

        first_name = driver.find_element_by_id("id_first_name")
        first_name.clear()
        first_name.send_keys(register_test.cell(row=test[0], column=4).value)

        last_name = driver.find_element_by_id("id_last_name")
        last_name.clear()
        last_name.send_keys(register_test.cell(row=test[1], column=4).value)

        email = driver.find_element_by_id("id_email")
        email.clear()
        email.send_keys(register_test.cell(row=test[2], column=4).value)

        psw_1 = driver.find_element_by_id("id_password1")
        psw_1.clear()
        psw_1.send_keys(register_test.cell(row=test[4], column=4).value)

        psw_2 = driver.find_element_by_id("id_password2")
        psw_2.clear()
        psw_2.send_keys(register_test.cell(row=test[5], column=4).value)

        submit_btn = driver.find_element_by_class_name("submitButton")
        submit_btn.click()
        if test[7] not in exception_case:
            alert_msg = driver.switch_to.alert
            WebDriverWait(driver, 10).until(EC.alert_is_present())

            if alert_msg.text == register_test.cell(row=test[7], column=4).value:
                print("Test case", i, "succeed!")
                succeed += 1
            else:
                print("Test case", i, "failed!")
                failed += 1

            alert_msg.accept()
        else:
            continue
driver.close()

print("--------------------------")
print("Report: Register Testing")
print("Tested cases:", len(test_case) - untested_cases)
print("Untested cases:", untested_cases)
print("Failed tests:", failed)
print("Succeed tests:", succeed)