# coding:utf-8

from selenium import webdriver

from selenium.webdriver.common.keys import Keys

import time, os, xlrd

from ActionKeywords import *

class DriverScript():
    bResult = False
    step = ''
    descript = ''
    page = ''
    find_by_type= ''
    page_element=''
    page_element_code=''
    operation = ''
    data = ''

    @staticmethod
    def task_scheduling():
        test_plan_path = os.getcwd()+"/TestSuite.xlsx"
        test_plan_table = xlrd.open_workbook(test_plan_path)
        test_plan_data = test_plan_table.sheet_by_name(u'任务执行计划')
        last_num = test_plan_data.nrows
        for i in range(1,last_num):
            test_plan_name = test_plan_data.row(i)[0].value
            test_plan_status = test_plan_data.row(i)[1].value
            if test_plan_status == 'Yes':
                DriverScript.run_test_suite(test_plan_name)

    @staticmethod
    def run_test_suite(test_suite_yes):
        test_plan_path = os.getcwd() + "/TestSuite.xlsx"
        test_plan_table = xlrd.open_workbook(test_plan_path)
        test_plan_data = test_plan_table.sheet_by_name(u'测试套件 ')
        last_num = test_plan_data.nrows
        for i in range(1,last_num):
            test_suite_name = test_plan_data.cell(i,0).value
            test_suite_case = test_plan_data.cell(i,2).value
            if test_suite_name == test_suite_yes:
                test_data_index = 0
                DriverScript.run_test_case(test_suite_case,test_data_index)

    @staticmethod
    def run_test_case(test_suite_case, test_data_index):
        test_case_path = os.getcwd() + "/TestCase.xlsx"
        test_case_table = xlrd.open_workbook(test_case_path)
        test_case_data = test_case_table.sheet_by_name(test_suite_case)
        last_num = test_case_data.nrows
        for i in range(1, last_num):
            DriverScript.step = test_case_data.cell(i, 0).value

            DriverScript.descript = test_case_data.cell(i, 1).value

            DriverScript.page = test_case_data.cell(i, 2).value

            DriverScript.find_by_type = test_case_data.cell(i, 3).value

            DriverScript.page_element = test_case_data.cell(i, 4).value

            DriverScript.page_element_code = test_case_data.cell(i, 5).value

            DriverScript.operation = test_case_data.cell(i, 6).value

            DriverScript.data = test_case_data.cell(i, 7).value

            print(DriverScript.step, DriverScript.descript, DriverScript.page, DriverScript.find_by_type, \
                  DriverScript.page_element, DriverScript.page_element_code, DriverScript.operation, DriverScript.data,)

            DriverScript.execute_actions()

    @staticmethod
    def execute_actions():
        if DriverScript.operation == 'openBrowser':
            ActionKeywords.openBrowser(DriverScript.data)
        elif DriverScript.operation == 'navigate':
            ActionKeywords.navigate(DriverScript.data)
        elif DriverScript.operation == 'input':
            ActionKeywords.input(DriverScript.find_by_type, DriverScript.page_element_code, DriverScript.data)
        elif DriverScript.operation == 'click':
            ActionKeywords.click(DriverScript.find_by_type, DriverScript.page_element_code)

if __name__ == "__main__":
    DriverScript.task_scheduling()
