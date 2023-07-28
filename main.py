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

class Main():

    def setup_method(self, method = None):
        self.driver = webdriver.Chrome()
        self.vars = {}
        self.driver.implicitly_wait(5)

    def teardown_method(self, method = None):
        self.driver.quit()

    def wait_for_window(self, timeout = 2):
        time.sleep(round(timeout / 1000))
        wh_now = self.driver.window_handles
        wh_then = self.vars["window_handles"]
        if len(wh_now) > len(wh_then):
            return set(wh_now).difference(set(wh_then)).pop()
    def importUser(self):
        self.driver.get("https://jamf.ulinkedu.com:8443/")
        self.driver.set_window_size(1800, 1061)
        time.sleep(10)
        self.driver.find_element(By.LINK_TEXT, "Users").click()
        self.driver.switch_to.frame(0)
        for i in range(len(pub_data_dict['CN_Name'])):
            row_data = {header: values[i] for header, values in pub_data_dict.items()}
            stu_cnName = row_data['CN_Name']
            stu_email_email = row_data['StuEmail']
            
            self.driver.find_element(By.CSS_SELECTOR, ".jamf-search > .jamf-button").click()
            self.driver.find_element(By.LINK_TEXT, "New").click()
            self.driver.find_element(By.ID, "FIELD_USERNAME").click()
            self.driver.find_element(By.ID, "FIELD_USERNAME").send_keys(stu_cnName)
            self.driver.find_element(By.ID, "FIELD_EMAIL_ADDRESS").click()
            self.driver.find_element(By.ID, "FIELD_EMAIL_ADDRESS").send_keys(stu_email_email)
            self.driver.find_element(By.ID, "save-button").click()
            self.driver.switch_to.default_content()
            self.driver.find_element(By.ID, "done-button").click()
            self.driver.switch_to.frame(0)
            time.sleep(1)

if __name__ == "__main__":
    if pub_valid_flag:
        test = Main()
        while True:
            print("1. 新增学生")
            print("2. 程序退出")
            choice = input("请输入操作序号: ")
            if choice == "1":
                test.setup_method()
                print("自动化开始")
                start_time = time.time()
                test.importUser()
                time.sleep(2)
                test.teardown_method()
                # 计算执行时间
                end_time = time.time()
                execution_time = end_time - start_time
                print(f"程序执行时间：{execution_time}秒")
                print("自动化结束")
            elif choice == "2":
                sys.exit()
            elif choice == "-1":
                continue
            else:
                print("无效序号，程序退出")
                sys.exit()
    else:
        print("请检查excel是否正确！")