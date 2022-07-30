# import sys
# from time import sleep
# import xlrd
# from selenium import webdriver
# from selenium.webdriver.common.by import By
#
#
# def read_excel(fire_name, assign_sheet=None):
#     """
#     读取用例文件（excel）
#     :param fire_name: 必传：传入需要读取的excel文件路径
#     :param assign_sheet: 非必传，传参数时读取整个excel文件
#                         传入需要读取的指定画布名称，则读取单
#                         画布内的用例数据。
#     :return: 未传assign_sheet时，返回整个excel全部用例数据。
#             传入assign_sheet时，返回指定画布下的用例数据
#     """
#     try:
#         book = xlrd.open_workbook(fire_name)
#         sheet_names = book.sheet_names()
#         suites = []
#         for i in range(len(sheet_names)):
#             if assign_sheet is not None:
#                 sheet_name = assign_sheet
#                 sheet = book.sheet_by_name(sheet_name)
#             else:
#                 sheet_name = sheet_names[i]
#                 sheet = book.sheet_by_name(sheet_name)
#
#             suite = {}
#             case_list = []
#             case_name = ''
#             flag_list = []
#             for line in range(sheet.nrows):  # 遍历行
#                 dataset = {}
#                 col = []
#                 flag = False
#                 for row in range(sheet.ncols):  # 遍历列
#                     if sheet.cell(line, 0).ctype == 2:
#                         if sheet.cell(line, 0).value == -1 and row == 0:
#                             case_name = sheet.cell(line, 1).value
#                             flag_list = []
#                             dataset[case_name] = flag_list
#                             case_list.append(dataset)
#                             # col.append(sheet.cell(line, 5).value)
#                         elif sheet.cell(line, 0).value > 0:
#                             case_name = sheet.cell(line, 1).value
#                             col.append(sheet.cell(line, row).value)
#                             flag = True
#                 if flag:
#                     dataset[case_name] = col
#                     flag_list.append(dataset)
#             if assign_sheet is not None:
#                 return case_list
#             suite[sheet_name] = case_list
#             suites.append(suite)
#         return suites
#
#     except xlrd.XLRDError as e:
#         print('Read excel error:%s' % e)
#         sys.exit(1)
#
#
# def processing(data, a=None):
#     """
#     对excel读取出来的数据进行分解，简化数据处理过程
#     :param data: 必传，传入excel读取出来的数据
#     :param a: 必传，传入你需要获取的数据的list名称
#                 例如：
#                 'sheets'-------excel画布对象
#                 'sheet_key'----excel画布名称（从画布对象中get value）
#                 'cases'--------用例对象
#                 'case_key'-----用例名称（从用例对象中get value）
#                 'steps'--------步骤对象
#                 'step_key'-----步骤名称（从步骤对象中get value）
#     :return:
#     """
#     sheets = []
#     sheet_key = []
#     suites = []
#     cases = []
#     case_key = []
#     step_list = []
#     steps = []
#     step_key = []
#     for sheet in data:
#         sheets.append(sheet)  # sheets对象
#         for sheet_name in sheet:
#             sheet_key.append(sheet_name)  # sheet_names
#
#     for i in range(len(sheets)):
#         suites.append(sheets[i].get(sheet_key[i]))  # cases集合(sheet_values)
#
#     for i in range(len(suites)):
#         for suite in suites[i]:
#             cases.append(suite)  # cases对象
#
#     for case in cases:
#         for case_name in case:
#             case_key.append(case_name)  # case_names
#
#     for i in range(len(cases)):
#         step_list.append(cases[i].get(case_key[i]))
#
#     for i in range(len(step_list)):
#         for suite in step_list[i]:
#             steps.append(suite)
#
#     for step in steps:
#         for step_name in step:
#             step_key.append(step_name)
#
#     if a == 'sheets':
#         return sheets
#     elif a == 'sheet_key':
#         return sheet_key
#     elif a == 'cases':
#         return cases
#     elif a == 'case_key':
#         return case_key
#     elif a == 'steps':
#         return steps
#     elif a == 'step_key':
#         return step_key
#     else:
#         return '参数错误'
#
#
# def get_step_data(fire_name, element=None):
#     data = read_excel(fire_name)
#     if element == '步骤ID':
#         index = 0
#     elif element == '执行步骤':
#         index = 1
#     elif element == '元素位置':
#         index = 2
#     elif element == '操作方式':
#         index = 3
#     elif element == '参数':
#         index = 4
#     elif element == '结果':
#         index = 5
#     else:
#         raise Exception('参数传递有误，检查一下吧~')
#
#     # 这里还没优化!!!!!
#     steps = processing(data, 'steps')
#     step_key = processing(data, 'step_key')
#     data_list = []
#     for i in range(len(steps)):
#         value = steps[i].get(step_key[i])[index]
#         data_list.append(value)
#     return data_list
#
#
# def run_demo(fire_name):
#     driver = webdriver.Chrome()
#     driver.get('http://lemo-saas.lemobar.cn/')
#     driver.maximize_window()
#     driver.implicitly_wait(10)
#
#     # 这里可以用枚举实现！！！！！
#
#
#
#
#
#
#
#
# # data = read_excel('../Cases/用例.xls')
# # initialize_data(data)
# # print(data)
# # with open
# # 'http://lemo-saas.lemobar.cn/'
# run_demo('../../../Massage_V1/Cases/用例.xls')
