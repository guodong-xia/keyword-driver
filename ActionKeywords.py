#coding:utf-8

from DriverScript import *
from selenium import webdriver

class ActionKeywords():
    driver = webdriver.Firefox()

    @staticmethod
    def openBrowser(data):
        if data == 'Mozilla':
            ActionKeywords.driver = webdriver.Firefox()
        else:
            DriverScript.bResult = False
            return
        ActionKeywords.driver.maximize_window()

    @staticmethod
    def navigate(data):
        ActionKeywords.driver.get(data)

    @staticmethod
    def input(find_by_type, page_element_code, data):
        ActionKeywords.high_light_element(find_by_type, page_element_code)
        ActionKeywords.driver.find_element(by=find_by_type, value=page_element_code).clear()
        ActionKeywords.driver.find_element(by=find_by_type, value=page_element_code).send_keys(data)

    @staticmethod
    def click(find_by_type, page_element_code):
        ActionKeywords.high_light_element(find_by_type, page_element_code)
        ActionKeywords.driver.find_element(by=find_by_type, value=page_element_code).click()

    @staticmethod
    def high_light_element(find_by_type, page_element_code):
        ActionKeywords.driver.execute_script("element = arguments[0];" +
                                             "original_style = element.getAttribute('style');" +
                                             "element.setAttribute('style', original_style + \";" +
                                             "background: yellow; border: 2px solid red;\");" +
                                             "setTimeout(function(){element.setAttribute('style', original_style);}, 1000);",
                                             ActionKeywords.driver.find_element(by=find_by_type,
                                                                                value=page_element_code))

    @staticmethod
    def capture():
        ActionKeywords.driver.save_screenshot("codingpy.png")
