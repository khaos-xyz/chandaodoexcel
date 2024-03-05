# -*-coding: Utf-8 -*-
# @File : caogao .py
"""
# author: zhangyuxiong66
# Time：2024/1/10 14:48
# @Desc:
"""
import re
import xlsxwriter
import time
import analysisbug
# print(time.time())
# print(time.strftime("%Y-%m-%d %H:%M:%S"))
# print(time.strftime("%m%d%H%M%S"))
today = time.strftime("%Y%m%d")


def doexcle():

        # 创建一个 Excel 文件对象
        workbook = xlsxwriter.Workbook(f"BUG统计日报--{today}.xlsx")
        # workbook = xlsxwriter.Workbook(f"BUG统计日报--{today}.xls")

        # 创建第一个工作表，并设置其名称为 'Bug统计'
        worksheet1 = workbook.add_worksheet('Bug统计')

        # 标题样式 border设置边框线 1表示实线 2表示虚线 4表示点线 8表示双实线
        title_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 20, 'border': 1})
        title_format.set_bg_color('D9D9D9')

        # 二级标题样式
        # 红色 FCB8B1
        # 草绿色 92D050
        # 绿色 00B050
        # 字体蓝色 4A99D2
        # 红色p1 C00000
        # 红色p2 D1443E
        # 红色p3 FF6D6D
        # 红色p4 F8AEB8
        erji_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 14, 'border': 1})
        # erji_format.set_bg_color('D9D9D9')
        pone_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 14, 'font_color': 'C00000', 'border': 1})
        ptwo_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 14, 'font_color': 'D1443E', 'border': 1})
        pthree_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 14, 'font_color': 'FF6D6D', 'border': 1})
        pfour_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 14, 'font_color': 'F8AEB8', 'border': 1})
        pfive_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 14, 'font_color': '4A99D2', 'border': 1})
        blue14_left_format = workbook.add_format({'bold': False, 'align': 'left', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 14, 'font_color': '4A99D2', 'border': 1})

        # 内容样式
        content_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 12, 'border': 1})
        content16left_format = workbook.add_format({'bold': False, 'align': 'left', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 16, 'border': 1})
        content20center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 20, 'border': 1})
        left_align_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 12, 'border': 1, 'text_wrap': True})
        center_align_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 12, 'border': 1})

        # 数据内容样式
        data_22bule_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 22, 'font_color': '4A99D2', 'border': 1})
        datasub_24bulebg_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 24, 'font_color': '5A61ED', 'border': 1})
        datasub_24bulebg_format.set_bg_color('FCB8B1')
        datasub_fix_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 24, 'font_color': '5A61ED', 'border': 1})
        datasub_fix_format.set_bg_color('92D050')
        datasub_off_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 24, 'font_color': '5A61ED', 'border': 1})
        datasub_off_format.set_bg_color('00B050')
        datasub_22black_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 22, 'border': 1})
        datasubp1_16redbg_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 16, 'font_color': 'C00000', 'border': 1})
        datasubp2_16redbg_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 16, 'font_color': 'D1443E', 'border': 1})
        datasubp3_16redbg_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 16, 'font_color': 'FF6D6D', 'border': 1})
        datasubp4_16redbg_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 16, 'font_color': 'F8AEB8', 'border': 1})
        datasubp1_18redbg_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 18, 'font_color': 'C00000', 'border': 1})
        datasubp2_18redbg_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 18, 'font_color': 'D1443E', 'border': 1})
        datasubp3_18redbg_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 18, 'font_color': 'FF6D6D', 'border': 1})
        datasubp4_18redbg_format = workbook.add_format({'bold': False, 'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑', 'font_size': 18, 'font_color': 'F8AEB8', 'border': 1})


        # 设置单元格样式 合并
        worksheet1.merge_range('A1:F1', '今日Bug统计', title_format)

        # 写入数据
        worksheet1.write('B2', '紧急（P1）', pone_format)
        worksheet1.write('C2', '高（P2）', ptwo_format)
        worksheet1.write('D2', '中（P3）', pthree_format)
        worksheet1.write('E2', '低（P4）', pfour_format)
        worksheet1.write('F2', '总计', pfive_format)

        worksheet1.write('A3', '今日新增bug', content_format)
        worksheet1.write('A4', '今日已解决bug', content_format)
        worksheet1.write('A5', '今天已关闭bug', content_format)

        # 设置B3到E5范围内单元格的样式
        # for row in range(2, 5):
        #     for col in range(1, 5):
        #         worksheet1.write(row, col, '', data_format)
        worksheet1.write('B3', analysisbug.new_p1_count, data_22bule_format)
        worksheet1.write('B4', analysisbug.fixed_p1_count, data_22bule_format)
        worksheet1.write('B5', analysisbug.closed_p1_count, data_22bule_format)
        worksheet1.write('C3', analysisbug.new_p2_count, data_22bule_format)
        worksheet1.write('C4', analysisbug.fixed_p2_count, data_22bule_format)
        worksheet1.write('C5', analysisbug.closed_p2_count, data_22bule_format)
        worksheet1.write('D3', analysisbug.new_p3_count, data_22bule_format)
        worksheet1.write('D4', analysisbug.fixed_p3_count, data_22bule_format)
        worksheet1.write('D5', analysisbug.closed_p3_count, data_22bule_format)
        worksheet1.write('E3', analysisbug.new_p4_count, data_22bule_format)
        worksheet1.write('E4', analysisbug.fixed_p4_count, data_22bule_format)
        worksheet1.write('E5', analysisbug.closed_p4_count, data_22bule_format)
        worksheet1.write('F3', '=SUM(B3:E3)', datasub_24bulebg_format)
        worksheet1.write('F4', '=SUM(B4:E4)',  datasub_fix_format)
        worksheet1.write('F5', '=SUM(B5:E5)',  datasub_off_format)
        # 列
        worksheet1.merge_range('A7:F7', '待处理Bug统计', title_format)
        worksheet1.write('A8', '所属模块', content16left_format)
        worksheet1.write('A9', '数据连接', content_format)
        worksheet1.write('A10', '任务管理', content_format)
        worksheet1.write('A11', '任务配置', content_format)
        worksheet1.write('A12', '映射配置', content_format)
        worksheet1.write('A13', '文件同步', content_format)
        worksheet1.write('A14', '运维中心', content_format)
        worksheet1.write('A15', '配置管理', content_format)
        worksheet1.write('A16', '引擎-任务执行', content_format)
        worksheet1.write('A17', '引擎-主备', content_format)
        worksheet1.write('A18', '引擎-数据读取', content_format)
        worksheet1.write('A19', '引擎-数据写入', content_format)
        worksheet1.write('A20', '引擎-类型转换', content_format)
        worksheet1.write('A21', '引擎-运行支撑', content_format)
        worksheet1.write('A22', '引擎-序列化异常', content_format)
        worksheet1.write('A23', '资料', content_format)
        worksheet1.write('A24', '挂起评审', content_format)
        worksheet1.write('A25', 'BUG剩余数', blue14_left_format)
        # 行
        worksheet1.write('B8', '紧急（P1）', pone_format)
        worksheet1.write('C8', '高（P2）', ptwo_format)
        worksheet1.write('D8', '中（P3）', pthree_format)
        worksheet1.write('E8', '低（P4）', pfour_format)
        worksheet1.write('F8', '总计', pfive_format)
        # 竖
        worksheet1.write('F9', '=SUM(B9:E9)',  datasub_22black_format)
        worksheet1.write('F10', '=SUM(B10:E10)',  datasub_22black_format)
        worksheet1.write('F11', '=SUM(B11:E11)',  datasub_22black_format)
        worksheet1.write('F12', '=SUM(B12:E12)',  datasub_22black_format)
        worksheet1.write('F13', '=SUM(B13:E13)',  datasub_22black_format)
        worksheet1.write('F14', '=SUM(B14:E14)',  datasub_22black_format)
        worksheet1.write('F15', '=SUM(B15:15)',  datasub_22black_format)
        worksheet1.write('F16', '=SUM(B16:E16)',  datasub_22black_format)
        worksheet1.write('F17', '=SUM(B17:E17)',  datasub_22black_format)
        worksheet1.write('F18', '=SUM(B18:E18)',  datasub_22black_format)
        worksheet1.write('F19', '=SUM(B19:E19)',  datasub_22black_format)
        worksheet1.write('F20', '=SUM(B20:E20)',  datasub_22black_format)
        worksheet1.write('F21', '=SUM(B21:E21)',  datasub_22black_format)
        worksheet1.write('F22', '=SUM(B22:E22)',  datasub_22black_format)
        worksheet1.write('F23', '=SUM(B23:E23)',  datasub_22black_format)
        worksheet1.write('F24', '=SUM(B24:E24)',  datasub_22black_format)
        worksheet1.write('B25', '=SUM(B9:B24)',  datasubp1_18redbg_format)
        worksheet1.write('C25', '=SUM(C9:C24)',  datasubp2_18redbg_format)
        worksheet1.write('D25', '=SUM(D9:D24)',  datasubp3_18redbg_format)
        worksheet1.write('E25', '=SUM(E9:E24)',  datasubp4_18redbg_format)
        worksheet1.write('F25', '=SUM(B25:E25)',  datasub_24bulebg_format)

        # 插入真实数据
        worksheet1.write('B9', analysisbug.bug_count_dict['active']['180']['1'], datasubp1_16redbg_format)
        worksheet1.write('B10', analysisbug.bug_count_dict['active']['181']['1'], datasubp1_16redbg_format)
        worksheet1.write('B11', analysisbug.bug_count_dict['active']['182']['1'], datasubp1_16redbg_format)
        worksheet1.write('B12', analysisbug.bug_count_dict['active']['183']['1'], datasubp1_16redbg_format)
        worksheet1.write('B13', analysisbug.bug_count_dict['active']['184']['1'], datasubp1_16redbg_format)
        worksheet1.write('B14', analysisbug.bug_count_dict['active']['185']['1']+analysisbug.bug_count_dict['active']['193']['1']+analysisbug.bug_count_dict['active']['194']['1']+analysisbug.bug_count_dict['active']['195']['1']+analysisbug.bug_count_dict['active']['196']['1']+analysisbug.bug_count_dict['active']['197']['1']+analysisbug.bug_count_dict['active']['198']['1']+analysisbug.bug_count_dict['active']['240']['1']+analysisbug.bug_count_dict['active']['244']['1'], datasubp1_16redbg_format)
        worksheet1.write('B15', analysisbug.bug_count_dict['active']['186']['1']+analysisbug.bug_count_dict['active']['208']['1']+analysisbug.bug_count_dict['active']['209']['1']+analysisbug.bug_count_dict['active']['210']['1']+analysisbug.bug_count_dict['active']['200']['1']+analysisbug.bug_count_dict['active']['202']['1']+analysisbug.bug_count_dict['active']['203']['1'], datasubp1_16redbg_format)
        worksheet1.write('B16', analysisbug.bug_count_dict['active']['187']['1']+analysisbug.bug_count_dict['active']['204']['1']+analysisbug.bug_count_dict['active']['205']['1']+analysisbug.bug_count_dict['active']['206']['1']+analysisbug.bug_count_dict['active']['207']['1'], datasubp1_16redbg_format)
        worksheet1.write('B17', analysisbug.bug_count_dict['active']['188']['1'], datasubp1_16redbg_format)
        worksheet1.write('B18', analysisbug.bug_count_dict['active']['189']['1'], datasubp1_16redbg_format)
        worksheet1.write('B19', analysisbug.bug_count_dict['active']['190']['1'], datasubp1_16redbg_format)
        worksheet1.write('B20', analysisbug.bug_count_dict['active']['191']['1'], datasubp1_16redbg_format)
        worksheet1.write('B21', analysisbug.bug_count_dict['active']['192']['1'], datasubp1_16redbg_format)
        worksheet1.write('B22', analysisbug.bug_count_dict['active']['219']['1'], datasubp1_16redbg_format)
        worksheet1.write('B23', analysisbug.bug_count_dict['active']['243']['1'], datasubp1_16redbg_format)
        worksheet1.write('B24', analysisbug.bug_count_dict['active']['148']['1'], datasubp1_16redbg_format)
        worksheet1.write('B9', analysisbug.bug_count_dict['active']['180']['1'], datasubp1_16redbg_format)

        worksheet1.write('C9', analysisbug.bug_count_dict['active']['180']['2'], datasubp2_16redbg_format)
        worksheet1.write('C10', analysisbug.bug_count_dict['active']['181']['2'], datasubp2_16redbg_format)
        worksheet1.write('C11', analysisbug.bug_count_dict['active']['182']['2'], datasubp2_16redbg_format)
        worksheet1.write('C12', analysisbug.bug_count_dict['active']['183']['2'], datasubp2_16redbg_format)
        worksheet1.write('C13', analysisbug.bug_count_dict['active']['184']['2'], datasubp2_16redbg_format)
        worksheet1.write('C14', analysisbug.bug_count_dict['active']['185']['2']+analysisbug.bug_count_dict['active']['193']['2']+analysisbug.bug_count_dict['active']['194']['2']+analysisbug.bug_count_dict['active']['195']['2']+analysisbug.bug_count_dict['active']['196']['2']+analysisbug.bug_count_dict['active']['197']['2']+analysisbug.bug_count_dict['active']['198']['2']+analysisbug.bug_count_dict['active']['240']['2']+analysisbug.bug_count_dict['active']['244']['2'], datasubp2_16redbg_format)
        worksheet1.write('C15', analysisbug.bug_count_dict['active']['186']['2']+analysisbug.bug_count_dict['active']['208']['2']+analysisbug.bug_count_dict['active']['209']['2']+analysisbug.bug_count_dict['active']['210']['2']+analysisbug.bug_count_dict['active']['200']['2']+analysisbug.bug_count_dict['active']['202']['2']+analysisbug.bug_count_dict['active']['203']['2'], datasubp2_16redbg_format)
        worksheet1.write('C16', analysisbug.bug_count_dict['active']['187']['2']+analysisbug.bug_count_dict['active']['204']['2']+analysisbug.bug_count_dict['active']['205']['2']+analysisbug.bug_count_dict['active']['206']['2']+analysisbug.bug_count_dict['active']['207']['2'], datasubp2_16redbg_format)
        worksheet1.write('C17', analysisbug.bug_count_dict['active']['188']['2'], datasubp2_16redbg_format)
        worksheet1.write('C18', analysisbug.bug_count_dict['active']['189']['2'], datasubp2_16redbg_format)
        worksheet1.write('C19', analysisbug.bug_count_dict['active']['190']['2'], datasubp2_16redbg_format)
        worksheet1.write('C20', analysisbug.bug_count_dict['active']['191']['2'], datasubp2_16redbg_format)
        worksheet1.write('C21', analysisbug.bug_count_dict['active']['192']['2'], datasubp2_16redbg_format)
        worksheet1.write('C22', analysisbug.bug_count_dict['active']['219']['2'], datasubp2_16redbg_format)
        worksheet1.write('C23', analysisbug.bug_count_dict['active']['243']['2'], datasubp2_16redbg_format)
        worksheet1.write('C24', analysisbug.bug_count_dict['active']['148']['2'], datasubp2_16redbg_format)

        worksheet1.write('D9', analysisbug.bug_count_dict['active']['180']['3'], datasubp3_16redbg_format)
        worksheet1.write('D10', analysisbug.bug_count_dict['active']['181']['3'], datasubp3_16redbg_format)
        worksheet1.write('D11', analysisbug.bug_count_dict['active']['182']['3'], datasubp3_16redbg_format)
        worksheet1.write('D12', analysisbug.bug_count_dict['active']['183']['3'], datasubp3_16redbg_format)
        worksheet1.write('D13', analysisbug.bug_count_dict['active']['184']['3'], datasubp3_16redbg_format)
        worksheet1.write('D14', analysisbug.bug_count_dict['active']['185']['3']+analysisbug.bug_count_dict['active']['193']['3']+analysisbug.bug_count_dict['active']['194']['3']+analysisbug.bug_count_dict['active']['195']['3']+analysisbug.bug_count_dict['active']['196']['3']+analysisbug.bug_count_dict['active']['197']['3']+analysisbug.bug_count_dict['active']['198']['3']+analysisbug.bug_count_dict['active']['240']['3']+analysisbug.bug_count_dict['active']['244']['3'], datasubp3_16redbg_format)
        worksheet1.write('D15', analysisbug.bug_count_dict['active']['186']['3']+analysisbug.bug_count_dict['active']['208']['3']+analysisbug.bug_count_dict['active']['209']['3']+analysisbug.bug_count_dict['active']['210']['3']+analysisbug.bug_count_dict['active']['200']['3']+analysisbug.bug_count_dict['active']['202']['3']+analysisbug.bug_count_dict['active']['203']['3'], datasubp3_16redbg_format)
        worksheet1.write('D16', analysisbug.bug_count_dict['active']['187']['3']+analysisbug.bug_count_dict['active']['204']['3']+analysisbug.bug_count_dict['active']['205']['3']+analysisbug.bug_count_dict['active']['206']['3']+analysisbug.bug_count_dict['active']['207']['3'], datasubp3_16redbg_format)
        worksheet1.write('D17', analysisbug.bug_count_dict['active']['188']['3'], datasubp3_16redbg_format)
        worksheet1.write('D18', analysisbug.bug_count_dict['active']['189']['3'], datasubp3_16redbg_format)
        worksheet1.write('D19', analysisbug.bug_count_dict['active']['190']['3'], datasubp3_16redbg_format)
        worksheet1.write('D20', analysisbug.bug_count_dict['active']['191']['3'], datasubp3_16redbg_format)
        worksheet1.write('D21', analysisbug.bug_count_dict['active']['192']['3'], datasubp3_16redbg_format)
        worksheet1.write('D22', analysisbug.bug_count_dict['active']['219']['3'], datasubp3_16redbg_format)
        worksheet1.write('D23', analysisbug.bug_count_dict['active']['243']['3'], datasubp3_16redbg_format)
        worksheet1.write('D24', analysisbug.bug_count_dict['active']['148']['3'], datasubp3_16redbg_format)

        worksheet1.write('E9', analysisbug.bug_count_dict['active']['180']['4'], datasubp4_16redbg_format)
        worksheet1.write('E10', analysisbug.bug_count_dict['active']['181']['4'], datasubp4_16redbg_format)
        worksheet1.write('E11', analysisbug.bug_count_dict['active']['182']['4'], datasubp4_16redbg_format)
        worksheet1.write('E12', analysisbug.bug_count_dict['active']['183']['4'], datasubp4_16redbg_format)
        worksheet1.write('E13', analysisbug.bug_count_dict['active']['184']['4'], datasubp4_16redbg_format)
        worksheet1.write('E14', analysisbug.bug_count_dict['active']['185']['4']+analysisbug.bug_count_dict['active']['193']['4']+analysisbug.bug_count_dict['active']['194']['4']+analysisbug.bug_count_dict['active']['195']['4']+analysisbug.bug_count_dict['active']['196']['4']+analysisbug.bug_count_dict['active']['197']['4']+analysisbug.bug_count_dict['active']['198']['4']+analysisbug.bug_count_dict['active']['240']['4']+analysisbug.bug_count_dict['active']['244']['4'], datasubp4_16redbg_format)
        worksheet1.write('E15', analysisbug.bug_count_dict['active']['186']['4']+analysisbug.bug_count_dict['active']['208']['4']+analysisbug.bug_count_dict['active']['209']['4']+analysisbug.bug_count_dict['active']['210']['4']+analysisbug.bug_count_dict['active']['200']['4']+analysisbug.bug_count_dict['active']['202']['4']+analysisbug.bug_count_dict['active']['203']['4'], datasubp4_16redbg_format)
        worksheet1.write('E16', analysisbug.bug_count_dict['active']['187']['4']+analysisbug.bug_count_dict['active']['204']['4']+analysisbug.bug_count_dict['active']['205']['4']+analysisbug.bug_count_dict['active']['206']['4']+analysisbug.bug_count_dict['active']['207']['4'], datasubp4_16redbg_format)
        worksheet1.write('E17', analysisbug.bug_count_dict['active']['188']['4'], datasubp4_16redbg_format)
        worksheet1.write('E18', analysisbug.bug_count_dict['active']['189']['4'], datasubp4_16redbg_format)
        worksheet1.write('E19', analysisbug.bug_count_dict['active']['190']['4'], datasubp4_16redbg_format)
        worksheet1.write('E20', analysisbug.bug_count_dict['active']['191']['4'], datasubp4_16redbg_format)
        worksheet1.write('E21', analysisbug.bug_count_dict['active']['192']['4'], datasubp4_16redbg_format)
        worksheet1.write('E22', analysisbug.bug_count_dict['active']['219']['4'], datasubp4_16redbg_format)
        worksheet1.write('E23', analysisbug.bug_count_dict['active']['243']['4'], datasubp4_16redbg_format)
        worksheet1.write('E24', analysisbug.bug_count_dict['active']['148']['4'], datasubp4_16redbg_format)


        # 设置列的宽度 索引方式
        # worksheet1.set_column(0, 0, 15)
        # worksheet1.set_column(1, 5, 10)
        # 设置列宽
        worksheet1.set_column('A:A', 30)
        worksheet1.set_column('B:F', 20)

        # 设置全局行高为30
        worksheet1.set_default_row(27.75)

        # 设置行高
        worksheet1.set_row(0, 32.75)

        # ==========================================================================================
        # ==========================================================================================
        # 创建第二个工作表，并设置其名称为 '待处理P1和P2的Bug列表'
        worksheet2 = workbook.add_worksheet('待处理P1和P2的Bug列表')
        worksheet2.merge_range('A1:E1', '待处理P1和P2的Bug列表', title_format)
        worksheet2.set_column('A:A', 12)
        worksheet2.set_column('B:B', 115)
        worksheet2.set_column('C:C', 24)
        worksheet2.set_column('D:D', 9)
        worksheet2.set_column('E:E', 25)
        worksheet2.set_default_row(35.25)

        headers = ['Bug编号', 'BUG标题', '指派人', '优先级', '备注']
        for col, header in enumerate(headers):
                worksheet2.write(1, col, header, erji_format)  # 表头从2行开始

        for row, bug in enumerate(analysisbug.prionetwo_bugs, start=2):  # 注意从第三行开始写入数据（因为第一行是标题行第二行表头行）
                for col, field in enumerate(headers):
                        if field == '备注':
                                module_value = int(bug['module'])
                                module_text = analysisbug.module_mapping.get(module_value, '')
                                worksheet2.write(row, 4, module_text, center_align_format)
                        else:
                                worksheet2.write(row, 0, bug['id'], center_align_format)
                                worksheet2.write(row, 1, bug['title'], left_align_format)
                                worksheet2.write(row, 2, bug['assignedTo'], center_align_format)
                                worksheet2.write(row, 3, bug['pri'], center_align_format)



        # ==========================================================================================
        # ==========================================================================================
        # 创建第三个工作表，并设置其名称为 '回归不通过bug'
        worksheet3 = workbook.add_worksheet("回归不通过的bug列表")
        worksheet3.merge_range('A1:G1', '回归不通过的bug列表', title_format)
        worksheet3.set_column('A:A', 12)
        worksheet3.set_column('B:B', 100)
        worksheet3.set_column('C:D', 10)
        worksheet3.set_column('E:F', 15)
        worksheet3.set_column('G:G', 18)
        worksheet3.set_default_row(35.25)
        headers = ['Bug编号', 'BUG标题', '优先级', '激活次数', '修复人', '审核人', '备注']
        for col, header in enumerate(headers):
                worksheet3.write(1, col, header, erji_format)
        for row, bug in enumerate(analysisbug.returned_bugs, start=2):
                worksheet3.write(row, 0, bug['id'], center_align_format)
                worksheet3.write(row, 1, bug['title'], left_align_format)
                worksheet3.write(row, 2, bug['pri'], center_align_format)
                worksheet3.write(row, 3, bug['activatedCount'], center_align_format)
                worksheet3.write(row, 4, bug['assignedTo'], center_align_format)
                worksheet3.write(row, 5, bug['lastEditedBy'], center_align_format)
                worksheet3.write(row, 6, '', center_align_format)


        # 保存 Excel 文件
        workbook.close()

        print("表格创建完成！")

if __name__ ==  "__main__":

        doexcle()