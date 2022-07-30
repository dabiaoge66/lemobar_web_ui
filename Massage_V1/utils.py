import sys
from enum import Enum
import xlrd
from selenium import webdriver
from time import sleep


class CaseElement(Enum):
    # 用例要素
    ID = '编号'
    STEP = '执行步骤'
    PLACE = '元素位置'
    WAY = '操作方式'
    PARAMETER = '参数'
    RESULT = '结果'
    # 数据键名
    SHEETS = 'sheets'
    SHEETKEY = 'sheet_key'
    CASE = 'cases'
    CASEKEY = 'case_key'
    STEPS = 'steps'
    STEPKEY = 'step_key'
    # 操作方式
    INPUT = 'input'
    CLICK = 'click'
    TEXT = 'text'
    SLEEP = 'sleep'


class GetCases(object):
    """
    用例管理类
            提供处理用例文件的类方法：
            read_excel:读取用例文件
            processing：对read_excel读取的数据进行预处理
            get_step_data：获取用例步骤中各要素的参数集合
    """

    @classmethod
    def read_excel(cls, fire_name, assign_sheet=None):
        """
        读取用例文件（excel）
        :param fire_name: 必传(str)：传入需要读取的excel文件路径
        :param assign_sheet: 非必传(str)，传参数时读取整个excel文件
                            传入需要读取的指定画布名称，则读取单
                            画布内的用例数据。
        :return: 未传assign_sheet时，返回整个excel全部用例数据。
                传入assign_sheet时，返回指定画布下的用例数据
        """
        try:
            book = xlrd.open_workbook(fire_name)
            sheet_names = book.sheet_names()
            suites = []
            for i in range(len(sheet_names)):
                if assign_sheet is not None:
                    sheet_name = assign_sheet
                    sheet = book.sheet_by_name(sheet_name)
                else:
                    sheet_name = sheet_names[i]
                    sheet = book.sheet_by_name(sheet_name)

                suite = {}
                case_list = []
                case_name = ''
                flag_list = []
                for line in range(sheet.nrows):  # 遍历行
                    dataset = {}
                    col = []
                    flag = False
                    for row in range(sheet.ncols):  # 遍历列
                        if sheet.cell(line, 0).ctype == 2:
                            if sheet.cell(line, 0).value == -1 and row == 0:
                                case_name = sheet.cell(line, 1).value
                                flag_list = []
                                dataset[case_name] = flag_list
                                case_list.append(dataset)
                                # col.append(sheet.cell(line, 5).value)
                            elif sheet.cell(line, 0).value > 0:
                                case_name = sheet.cell(line, 1).value
                                col.append(sheet.cell(line, row).value)
                                flag = True
                    if flag:
                        dataset[case_name] = col
                        flag_list.append(dataset)
                if assign_sheet is not None:
                    return case_list
                suite[sheet_name] = case_list
                suites.append(suite)
            return suites

        except xlrd.XLRDError as e:
            print('Read excel error:%s' % e)
            sys.exit(1)

    @classmethod
    def processing(cls, data, a=None):
        """
        对excel读取出来的数据进行分解，简化数据处理过程
        :param data: 必传(str)，传入excel读取出来的数据
        :param a: 必传(str)，传入你需要获取的数据的list名称
                    例如：
                    'sheets'-------excel画布对象
                    'sheet_key'----excel画布名称（从画布对象中get value）
                    'cases'--------用例对象
                    'case_key'-----用例名称（从用例对象中get value）
                    'steps'--------步骤对象
                    'step_key'-----步骤名称（从步骤对象中get value）
        :return:入参key对应值集合
        """
        sheets = []  # 画布对象(dict)
        sheet_key = []  # 画布key--获取用例集(str)
        suites = []  # 套件(list[用例对象dict])
        cases = []  # 用例对象(dict)
        case_key = []  # 用例key--获取步骤集(str)
        step_list = []  # 套件(list[步骤对象dict])
        steps = []  # 步骤对象(dict)
        step_key = []  # 步骤key--获取参数集(str)
        for sheet in data:
            sheets.append(sheet)  # sheets对象
            for sheet_name in sheet:
                sheet_key.append(sheet_name)  # sheet_names

        for i in range(len(sheets)):
            suites.append(sheets[i].get(sheet_key[i]))  # cases集合(sheet_values)

        for i in range(len(suites)):
            for suite in suites[i]:
                cases.append(suite)  # cases对象

        for case in cases:
            for case_name in case:
                case_key.append(case_name)  # case_names

        for i in range(len(cases)):
            step_list.append(cases[i].get(case_key[i]))

        for i in range(len(step_list)):
            for suite in step_list[i]:
                steps.append(suite)

        for step in steps:
            for step_name in step:
                step_key.append(step_name)

        if a == CaseElement.SHEETS.value:
            return sheets
        elif a == CaseElement.SHEETKEY.value:
            return sheet_key
        elif a == CaseElement.CASE.value:
            return cases
        elif a == CaseElement.CASEKEY.value:
            return case_key
        elif a == CaseElement.STEPS.value:
            return steps
        elif a == CaseElement.STEPKEY.value:
            return step_key
        else:
            raise Exception('入参错误喔，检查一下吧~')

    @classmethod
    def get_step_data(cls, fire_name, element=None, assign_sheet=None):
        """
        获取用例文件的要素中的值集合
        :param assign_sheet: 非必传(str):指定画布名
        :param fire_name:必传(str)：用例文件名
        :param element:必传(str)：根据要素名称获取指定值的集合
        :return:入参key对应值集合
        """
        data = cls.read_excel(fire_name, assign_sheet)
        if element == CaseElement.ID.value:
            index = 0
        elif element == CaseElement.STEP.value:
            index = 1
        elif element == CaseElement.PLACE.value:
            index = 2
        elif element == CaseElement.WAY.value:
            index = 3
        elif element == CaseElement.PARAMETER.value:
            index = 4
        elif element == CaseElement.RESULT.value:
            index = 5
        else:
            raise Exception('入参错误喔，检查一下吧~')

        steps = cls.processing(data, CaseElement.STEPS.value)
        step_key = cls.processing(data, CaseElement.STEPKEY.value)
        data_list = []
        for i in range(len(steps)):
            value = steps[i].get(step_key[i])[index]
            data_list.append(value)
        return data_list


class DriverUtil(object):
    """浏览器对象管理类"""
    __driver = None

    @classmethod
    def get_driver(cls, website):
        """"获取浏览器对象方法"""
        if cls.__driver is None:
            cls.__driver = webdriver.Chrome()
            cls.__driver.get(website)
            cls.__driver.maximize_window()
            cls.__driver.implicitly_wait(5)
        return cls.__driver

    @classmethod
    def quit_driver(cls):
        """"退出浏览器对象方法"""
        sleep(2)
        cls.__driver.quit()
        cls.__driver = None
