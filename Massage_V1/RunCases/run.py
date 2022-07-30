import pytest
from Massage_V1.utils import DriverUtil
from Massage_V1.Cases.case import ReadyCase


class TestRun(object):
    """登录测试类"""

    def setup_class(self):
        self.ready_case = ReadyCase
        self.ready_case.driver = DriverUtil.get_driver('http://lemo-saas.lemobar.cn/')  # 获取浏览器对象
        # self.ready_case.fire_name = '../Cases/用例.xls'

    def teardown_class(self):
        pass
        # DriverUtil.quit_driver()  # 退出浏览器对象

    def setup(self):
        pass
        # self.driver.get('http://lemo-saas.lemobar.cn/')  # 打开测试地址

    def teardown(self):
        pass
        DriverUtil.quit_driver()

    def test_add(self):
        self.ready_case.handle_elements()


if __name__ == '__main__':
    pytest.main(['-s', 'run.py'])
