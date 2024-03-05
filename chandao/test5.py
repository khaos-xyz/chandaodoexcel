# -*-coding: Utf-8 -*-
# @File : test5 .py
"""
# author: zhangyuxiong66
# Time：2024/1/31 16:16
# @Desc:
"""
from colorama import win32


def create_save_img(file_path, sheetname, img_name):

    # 读取excel内容转换为图片
    from PIL import ImageGrab
    import xlwings as xw
    import pythoncom
    import os
    import time

    # 强制终止所有名为"EXCEL.exe"的Excel进程
    os.system('taskkill /IM EXCEL.exe /F')
    pythoncom.CoInitialize()
    # 使用xlwings的app启动
    app = xw.App(visible=False, add_book=False)

    # 打开Excel应用程序
    # app = win32.Dispatch("Excel.Application")

    # 打开文件
    file_path = os.path.abspath(file_path)
    wb = app.books.open(file_path)
    # 选定sheet
    sheet = wb.sheets(sheetname)
    # 获取有内容的区域
    all = sheet.used_range
    # 复制图片区域
    all.api.CopyPicture()
    # 粘贴
    sheet.api.Paste()
    # 当前图片
    pic = sheet.pictures[-1]
    # 复制图片
    pic.api.Copy()
    time.sleep(3)# 延迟一下操作，不然获取不到图片
    # 获取剪贴板的图片数据
    img = ImageGrab.grabclipboard()
    # 保存图片
    img.save(img_name)
    # 删除sheet上的图片
    pic.delete()
    # 不保存，直接关闭
    wb.close()
    # 退出xlwings的app启动
    app.quit()
    pythoncom.CoUninitialize()  # 关闭多线程
    # os.remove(file_path)


#路径

filename = r"G:\PycharmProjects\pythonProject\chandao\BUG统计日报--20240130.xlsx"
# filename = r"C:\Users\admin\Desktop\有表头3列  UTF-8.csv"

create_save_img(filename, 'Bug统计', 'Excel复制成图片的例子.png')