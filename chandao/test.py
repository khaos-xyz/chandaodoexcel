import pandas as pd
import matplotlib.pyplot as plt
from pandas.plotting import table
from PIL import Image
from io import BytesIO

from pylab import mpl

mpl.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # 指定默认字体：解决plot不能显示中文问题
mpl.rcParams['axes.unicode_minus'] = False


xls = pd.ExcelFile('BUG统计日报--20240130.xlsx')
j = 0
for sheet_name in xls.sheet_names:
    # 读取工作表数据
    df = pd.read_excel(xls, sheet_name=sheet_name)
    # 将缺失值转为空字符串，避免表格中出现NaN
    df = df.applymap(lambda x: str(x) if pd.notna(x) else '')
    # 创建一个样式对象

    print("正在处理 %s 表格" % sheet_name)
    j = j + 1

    if df.shape[0] > 150:  # 假设以100行作为一个分割点

        df_split = [df[i:i + 150] for i in range(0, df.shape[0], 150)]

        for i, data in enumerate(df_split):
            fig = plt.figure(figsize=(9, 26), dpi=200)
            ax = fig.add_subplot(111, frame_on=False)

            # 隐藏x轴 y轴
            ax.xaxis.set_visible(False)  # hide the x axis
            ax.yaxis.set_visible(False)  # hide the y axis

            fig.tight_layout()
            # 如果数据是图片，则保存为图片，否则执行表格处理
            if isinstance(data.iloc[0, 0], Image.Image):  # 检查数据是否为图片
                img = data.iloc[:, 0].apply(lambda x: x.tobytes())  # 将图片转换为字节流
                img_buffer = BytesIO()
                img_buffer.write(img.values.tobytes())  # 将字节流转为BytesIO对象
                img = Image.open(img_buffer)  # 将BytesIO对象转为图片
                img.save(f'{sheet_name}.png')  # 保存图片
            else:

                ax.table(cellText=data.values, colLabels=data.columns, cellLoc='center', loc='center')
                ax.set_title(sheet_name, fontsize=12, loc='left')
                #ax.text(0, 1, sheet_name, ha='left', va='center', fontsize=12, transform=ax.transAxes)
                plt.savefig(f'{sheet_name}.png')
            plt.close(fig)
            print("已完成 %s 表格处理" % sheet_name)
    elif df.shape[0] > 50:
        if isinstance(df.iloc[0, 0], Image.Image):  # 检查数据是否为图片
            img = df.iloc[:, 0].apply(lambda x: x.tobytes())  # 将图片转换为字节流
            img_buffer = BytesIO()
            img_buffer.write(img.values.tobytes())  # 将字节流转为BytesIO对象
            img = Image.open(img_buffer)  # 将BytesIO对象转为图片
            img.save(f'{sheet_name}.png')  # 保存图片
        else:

            fig = plt.figure(figsize=(9, 10), dpi=500)
            ax = fig.add_subplot(111, frame_on=False)
            ax.set_title(sheet_name, fontsize=12, loc='left')
            #ax.text(0, 0.8, sheet_name, ha='left', va='center', fontsize=12, transform=ax.transAxes)
            table1=ax.table(cellText=df.values, colLabels=df.columns, cellLoc='center', loc='center')
            table1.auto_set_font_size(True)  # 禁用自动设置字体大小
            table1.set_fontsize(14)  # 设置字体大小为12
            ax.xaxis.set_visible(False)  # hide the x axis
            ax.yaxis.set_visible(False)  # hide the y axis
            fig.tight_layout()
            plt.savefig(f'{sheet_name}.png')
        plt.close(fig)
        print("已完成 %s 表格处理" % sheet_name)


    else:
        if isinstance(df.iloc[0, 0], Image.Image):  # 检查数据是否为图片
            img = df.iloc[:, 0].apply(lambda x: x.tobytes())  # 将图片转换为字节流
            img_buffer = BytesIO()
            img_buffer.write(img.values.tobytes())  # 将字节流转为BytesIO对象
            img = Image.open(img_buffer)  # 将BytesIO对象转为图片
            img.save(f'{sheet_name}.png')  # 保存图片
        else:

            fig = plt.figure(figsize=(9, 6), dpi=500)
            ax = fig.add_subplot(111, frame_on=False)
            ax.set_title(sheet_name, fontsize=12, loc='left')
            #设置标题在x,y轴上位置，字体
            #ax.text(0, 0.8, sheet_name, ha='left', va='center', fontsize=12, transform=ax.transAxes)
            table1=ax.table(cellText=df.values, colLabels=df.columns, cellLoc='center', loc='center')
            table1.auto_set_font_size(True)  # 禁用自动设置字体大小
            table1.set_fontsize(14)  # 设置字体大小为12

            ax.xaxis.set_visible(False)  # hide the x axis
            ax.yaxis.set_visible(False)  # hide the y axis

            #fig.tight_layout()
            # plt.savefig('./output/' + str(j) + '!' + sheet_name + '.jpg')
            plt.savefig(f'{sheet_name}.png')
        plt.close(fig)
        print("已完成 %s 表格处理" % sheet_name)
