# -*-coding: Utf-8 -*-
# @File : test2 .py
"""
# author: zhangyuxiong66
# Time：2024/1/30 17:03
# @Desc:
"""
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.table import Table
from PIL import Image, ImageDraw, ImageFont

def generate_table_image():
    # 创建一个空白图像
    img = Image.new('RGB', (800, 600), color='white')
    draw = ImageDraw.Draw(img)

    # 设置字体
    font = ImageFont.truetype('arial.ttf', size=12)

    # 读取数据
    data = pd.read_excel('BUG统计日报--20240130.xlsx')

    # 获取表格的行数和列数
    num_rows, num_cols = data.shape

    # 设置表格的大小和位置
    table_width = 700
    table_height = 400
    table_x = (800 - table_width) // 2
    table_y = (600 - table_height) // 2

    # 绘制表格边框
    draw.rectangle([table_x, table_y, table_x + table_width, table_y + table_height], outline='black')

    # 计算每个单元格的宽度和高度
    cell_width = table_width / num_cols
    cell_height = table_height / num_rows

    # 设置表头的背景色
    header_bg_color = '#C0C0C0'

    # 设置表格内容的背景色
    content_bg_color = '#FFFFFF'

    # 设置表头和内容的字体颜色
    header_text_color = '#000000'
    content_text_color = '#000000'

    # 绘制表头
    for j, header in enumerate(data.columns):
        cell_x = table_x + j * cell_width
        cell_y = table_y
        cell_width_actual = cell_width
        cell_height_actual = cell_height

        # 合并单元格
        if j == 0:
            cell_width_actual *= 2

        # 绘制背景色
        draw.rectangle([cell_x, cell_y, cell_x + cell_width_actual, cell_y + cell_height_actual],
                       fill=header_bg_color)

        # 绘制文字
        text_x = cell_x + cell_width_actual / 2
        text_y = cell_y + cell_height_actual / 2 - font.size / 2
        draw.text((text_x, text_y), str(header), fill=header_text_color, font=font, anchor='mm')

    # 绘制表格内容
    for i in range(num_rows):
        for j in range(num_cols):
            cell_x = table_x + j * cell_width
            cell_y = table_y + (i + 1) * cell_height
            cell_width_actual = cell_width
            cell_height_actual = cell_height

            # 合并单元格
            if j == 0:
                cell_width_actual *= 2

            # 绘制背景色
            draw.rectangle([cell_x, cell_y, cell_x + cell_width_actual, cell_y + cell_height_actual],
                           fill=content_bg_color)

            # 绘制文字
            text_x = cell_x + cell_width_actual / 2
            text_y = cell_y + cell_height_actual / 2 - font.size / 2
            draw.text((text_x, text_y), str(data.iloc[i, j]), fill=content_text_color, font=font, anchor='mm')

    # 保存图像
    img.save('table_image.png')
    print("表格图片生成完成！")

# 调用生成表格图片的函数
generate_table_image()


# if __name__ == '__main__':
#     # 使用示例
#     excel_to_image_and_send('ceshi@jiuqingsk.com', 'jss3mPZnw7wFSigR', 'zhangyx@jiuqingsk.com',
#                             'BUG统计日报--20240130.xlsx')