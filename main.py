from data_extract import pub_data_dict
from data_extract import pub_valid_flag

# Generated by Selenium IDE
import pytest
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoSuchElementException

import sys

jss_url = "https://jamf.ulinkedu.com:8443/"

class Main():

    def setup_method(self, method = None):
        self.driver = webdriver.Chrome()
        self.vars = {}

    def teardown_method(self, method = None):
        self.driver.quit()

    def wait_for_window(self, timeout = 2):
        time.sleep(round(timeout / 1000))
        wh_now = self.driver.window_handles
        wh_then = self.vars["window_handles"]
        if len(wh_now) > len(wh_then):
            return set(wh_now).difference(set(wh_then)).pop()
    def importUser(self):
        self.driver.get(jss_url)
        self.driver.set_window_size(1800, 1061)
        time.sleep(10)
        self.driver.find_element(By.LINK_TEXT, "Users").click()
        self.driver.switch_to.frame(0)
        self.driver.find_element(By.CSS_SELECTOR, ".jamf-search > .jamf-button").click()
        self.driver.find_element(By.LINK_TEXT, "New").click()
        for i in range(len(pub_data_dict['CN_Name'])):
            row_data = {header: values[i] for header, values in pub_data_dict.items()}
            
            stu_loginID = row_data['Login_ID']
            stu_relName = row_data['CN_Name']
            stu_email_email = row_data['StuEmail']
            stu_form = row_data['Form']
            
            # tea_loginID = row_data['Login_Name']
            # tea_relName = row_data['CN_Name']
            # tea_email_email = row_data['Email']
            # tea_dept = row_data['Department']
            
            
            self.driver.find_element(By.ID, "FIELD_USERNAME").click()
            self.driver.find_element(By.ID, "FIELD_USERNAME").clear()
            self.driver.find_element(By.ID, "FIELD_USERNAME").send_keys(stu_loginID)
            # self.driver.find_element(By.ID, "FIELD_USERNAME").send_keys(tea_loginID)
            
            self.driver.find_element(By.ID, "FIELD_REAL_NAME").click()
            self.driver.find_element(By.ID, "FIELD_REAL_NAME").clear()
            self.driver.find_element(By.ID, "FIELD_REAL_NAME").send_keys(stu_relName)
            # self.driver.find_element(By.ID, "FIELD_REAL_NAME").send_keys(tea_relName)
            
            self.driver.find_element(By.ID, "FIELD_EMAIL_ADDRESS").click()
            self.driver.find_element(By.ID, "FIELD_EMAIL_ADDRESS").clear()
            self.driver.find_element(By.ID, "FIELD_EMAIL_ADDRESS").send_keys(stu_email_email)
            # self.driver.find_element(By.ID, "FIELD_EMAIL_ADDRESS").send_keys(tea_email_email)
            
            self.driver.find_element(By.ID, "FIELD_POSITION").click()
            self.driver.find_element(By.ID, "FIELD_POSITION").clear()
            self.driver.find_element(By.ID, "FIELD_POSITION").send_keys(stu_form)
            # self.driver.find_element(By.ID, "FIELD_POSITION").send_keys(tea_dept)
            
            self.driver.find_element(By.ID, "save-button").click()
            self.driver.find_element(By.ID, "clone-button").click()

if __name__ == "__main__":
    if pub_valid_flag:
        test = Main()
        while True:
            print("1. Add User")
            print("2. Exit")
            choice = input("Please enter your choice code: ")
            if choice == "1":
                test.setup_method()
                print("Start Automation")
                start_time = time.time()
                test.importUser()
                time.sleep(2)
                test.teardown_method()
                # 计算执行时间
                end_time = time.time()
                execution_time = end_time - start_time
                print(f"Excution Time：{execution_time}秒")
                print("End Automation")
            elif choice == "2":
                sys.exit()
            elif choice == "-1":
                continue
            else:
                print("Invalid input, please re-enter！")
                sys.exit()
    else:
        print("Please check the excel format！")