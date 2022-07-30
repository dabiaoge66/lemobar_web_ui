from time import sleep
from selenium.webdriver.common.by import By
from Massage_V1 import utils
from Massage_V1.utils import CaseElement


class ReadyCase(object):
    """用例执行类"""
    driver = None  # 获取浏览器对象
    fire_name = '../Cases/用例.xls'  # 用例文件路径
    case_id = utils.GetCases.get_step_data(fire_name, element=CaseElement.ID.value)  # 用例编号list
    case_step = utils.GetCases.get_step_data(fire_name, element=CaseElement.STEP.value)  # 用例步骤list
    element_place = utils.GetCases.get_step_data(fire_name, element=CaseElement.PLACE.value)  # 元素位置list
    handle_way = utils.GetCases.get_step_data(fire_name, element=CaseElement.WAY.value)  # 操作方式list
    parameter = utils.GetCases.get_step_data(fire_name, element=CaseElement.PARAMETER.value)  # 参数list
    case_result = utils.GetCases.get_step_data(fire_name, element=CaseElement.RESULT.value)  # 执行结果list

    @classmethod
    def handle_elements(cls):
        for i in range(len(cls.handle_way)):
            if cls.element_place[i] != '':
                elements = (By.XPATH, cls.element_place[i])  # 定位元素（写死XPATH）--牺牲定位效率，简化用例编写
                operate = cls.driver.find_element(elements[0], elements[1])
                if cls.handle_way[i] == CaseElement.INPUT.value:  # 键入
                    operate.send_keys(cls.parameter[i])
                elif cls.handle_way[i] == CaseElement.CLICK.value:  # 点击
                    operate.click()
                elif cls.handle_way[i] == CaseElement.TEXT.value:  # 获取文本
                    return operate.text
            elif cls.handle_way[i] == CaseElement.SLEEP.value:  # 显式等待
                sleep(cls.parameter[i])
                print('睡眠中~~~还再等 %d 秒~~~' % cls.parameter[i])
            else:
                raise Exception('元素位置为空哦，检查一下吧~')
