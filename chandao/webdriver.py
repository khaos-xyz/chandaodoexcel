# -*-coding: Utf-8 -*-
# @File : webdriver .py
"""
# author: zhangyuxiong66
# Time：2023/12/27 10:28
# @Desc:
"""
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time  # 定时

from selenium.webdriver.support.select import Select  # 下拉框选择

import datetime  # 当前日期
from time import strftime  # 当前日期
from selenium.webdriver.common.keys import Keys  # 为了在webdriver中按ESC键

import pandas as pd
import os

option = webdriver.ChromeOptions()
option.binary_location = r'C:\Program Files\Google\Chrome\Application\chrome.exe'  # 这里添加edge的启动文件=>chrome的话添加chrome.exe的绝对路径

diy_prefs = {'profile.default_content_settings.popups': 0,
             'download.default_directory': r'C:\Users\Administrator\Desktop\\'}
# 设置下载路径到selenium配置中
option.add_experimental_option('prefs', diy_prefs)

driver = webdriver.Edge(service=Service(r'G:\Python38\chromedriver.exe'), options=option)  # 这里添加的是driver的绝对路径，驱动去微软官网下载

driver.get('http://zentao.jiuqingsk.com/zentao/user-login.html')  # 打开禅道登录页
driver.maximize_window()

# 输入账号
driver.find_element(By.ID, 'account').send_keys('zhangyuxiong')
# driver.find_element 的新语法样式。

# 输入密码
driver.find_element(By.NAME, 'password').send_keys('Zyx123456')
# driver.find_element(By.Name, 'loginUsername')


# 点击登录
time.sleep(3)
driver.find_element(By.ID, 'submit').click()
time.sleep(3)

# 跳转到全部BUG列表 菜单路径：测试>bug>所有
driver.get("http://zentao.jiuqingsk.com/zentao/bug-browse-4-all-all.html")
time.sleep(5)

# # 点击搜索   放大镜图标那个显示查询条件
# driver.find_element(By.ID, 'bysearchTab').click()
# time.sleep(5)
#
# # 选择解决日期 大于等于是上周三
# driver.find_element(By.LINK_TEXT, 'Bug标题').click()
# time.sleep(1)
# driver.find_element(By.CSS_SELECTOR, '.active-result:nth-child(31)').click()
#
# # 选择解决时间条件
# Select(driver.find_element(By.ID, 'operator1')).select_by_value(">=")
#
# # 计算上周三日期
# today = datetime.date.today()
# if today.weekday() + 1 > 3:
#     last_wednesday = today - datetime.timedelta(7 + ((today.weekday() + 1) - 3))
# elif today.weekday() + 1 < 3:
#     last_wednesday = today - datetime.timedelta(7 - (3 - (today.weekday() + 1)))
# else:
#     last_wednesday = today - datetime.timedelta(7)
# print('上周三是', last_wednesday)
#
# driver.find_element(By.CSS_SELECTOR, '#value1').send_keys(last_wednesday.strftime("%Y%m%d"))
# # 这个地方日期组合控件有坑，原思路是点击录入框在日期控件选日期，后来发现可以直接给录入框赋值。
# time.sleep(1)
# driver.find_element(By.CSS_SELECTOR, '#value1').send_keys(Keys.ESCAPE)
# # 这个是为了解决赋值后，弹出日期控件的问题，不过应该可以省略掉，在下一步操作日期控件会消失。
#
#
# # 选择解决日期 结束小于本周三
# driver.find_element(By.LINK_TEXT, '重现步骤').click()
# time.sleep(1)
# driver.find_element(By.CSS_SELECTOR, '#field4_chosen > div > ul > li:nth-child(31)').click()
#
# # 选择解决时间条件
# Select(driver.find_element(By.ID, 'operator4')).select_by_value("<")
#
# # 计算本周三日期
# if today.weekday() + 1 > 3:
#     now_wednesday = today - datetime.timedelta((today.weekday() + 1) - 3)
# elif today.weekday() + 1 < 3:
#     now_wednesday = today + datetime.timedelta(3 - (today.weekday() + 1))
# else:
#     now_wednesday = today
# print('本周三是', now_wednesday)
#
# driver.find_element(By.CSS_SELECTOR, '#value4').send_keys(now_wednesday.strftime("%Y%m%d"))
# time.sleep(1)
# driver.find_element(By.CSS_SELECTOR, '#value4').send_keys(Keys.ESCAPE)
#
# # 点击蓝色搜索
# driver.find_element(By.ID, 'submit').click()
# time.sleep(1)

# 点击导出
# time.sleep(5)
# driver.find_element(By.CSS_SELECTOR, ".btn-group:nth-child(2) > .btn-link").click()
# 等待父元素加载完成
print(1)
# 显式等待直到按钮可点击
button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@class="btn btn-link dropdown-toggle"]/div[3]/div[1]'))
)
button.click()
<button type="button" class="" data-toggle="dropdown" style="border-radius: 4px;">
        <i class="icon icon-export muted"></i> <span class="text"> 导出</span> <span class="caret"></span></button>
print(2)
# parent_element.click()
print(3)

driver.find_element(By.LINK_TEXT, "导出数据").click()
print(4)

print(5)
time.sleep(10)

# 点击导出数据
# driver.find_element(By.LINK_TEXT, "导出").click()
# time.sleep(5)
# driver.find_element(By.LINK_TEXT, "导出数据").click()
# time.sleep(10)

filename = datetime.datetime.now().strftime('%Y年%m月%d日')
filename = filename + "解决BUG清单周报"
# 输入文件名
driver.switch_to.frame(1)  # 1.用frame的index来定位，第一个是0
# driver.switch_to.frame("frame1")  # 2.用id来定位
# driver.switch_to.frame("myframe")  # 3.用name来定位
# driver.switch_to.frame(driver.find_element_by_tag_name("iframe"))  # 4.用WebElement对象来定位

driver.find_element(By.CSS_SELECTOR, '#fileName').click()
driver.find_element(By.CSS_SELECTOR, '#fileName').send_keys(filename)

# 点击导出
driver.find_element(By.ID, 'submit').click()
time.sleep(5)
#
# CSV_fileName = r'C:\Users\Administrator\Desktop\\' + filename + r'.csv'
# file_name, old_extension = os.path.splitext(filename)
#
# EXCLE_fileName = r'C:\Users\Administrator\Desktop\\' + file_name + r'.xlsx';
# # 读取 CSV 文件并转换为 DataFrame
# df = pd.read_csv(CSV_fileName, index_col=False, encoding='utf-8')
# # 这个地方也是个坑原先写的是index_col=None 会导致转换到EXCEL丢失BUG编号内容
#
#
# # 将 DataFrame 保存为 XLSX 文件
# df.to_excel(EXCLE_fileName, index=False)

time.sleep(20)
driver.quit()