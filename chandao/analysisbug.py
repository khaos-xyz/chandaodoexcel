# -*-coding: Utf-8 -*-
# @File : 解析bug .py
"""
# author: zhangyuxiong66
# Time：2024/1/19 15:39
# @Desc:
"""
import getallbug
import datetime
from collections import defaultdict

all_bug = getallbug.get_bugs()
prionetwo_bugs = []
returned_bugs = []
# 定义一个 defaultdict，用于记录每个模块和优先级的 bug 个数
# 这里的 bug_count_dict 是一个嵌套的 defaultdict 对象，它的第一层键是模块名（例如 A、B、C），它的第二层键是优先级（例如 P1、P2、P3、P4）。这样我们就可以通过 bug_count_dict[module][priority] 来访问每个模块下的不同优先级的 bug 个数。
bug_count_dict = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
# module_A_priority_P1_count = bug_count_dict['A']['P1']
# 变量表示模块 A 下优先级为 P1 的 bug 个数，module_B_priority_P3_count 变量表示模块 B 下优先级为 P3 的 bug 个数
module_mapping = {
    0: '/',
    180: '数据连接',
    181: '任务管理',
    182: '任务配置',
    183: '映射配置',
    184: '文件同步',
    185: '运维中心',
    193: "/运维中心/驾驶舱",
    194: "/运维中心/运维大盘",
    195: "/运维中心/实例运维",
    196: "/运维中心/智能监控",
    197: "/运维中心/审计日志",
    198: "/运维中心/消息中心",
    240: "/运维中心/引擎运维",
    244: "/运维中心/数据对账",
    186: '配置管理',
    208: "/配置管理/项目",
    209: "/配置管理/脱敏规则",
    210: "/配置管理/通知推送管理",
    200: "/配置管理/用户管理",
    202: "/配置管理/用户管理/角色管理",
    203: "/配置管理/用户管理/用户管理",
    187: '引擎-任务执行',
    204: '引擎-任务执行/任务提交',
    205: '引擎-任务执行/任务状态',
    206: '引擎-任务执行/任务指标',
    207: '引擎-任务执行/异常日志',
    188: '引擎-主备',
    189: '引擎-数据读取',
    190: '引擎-数据写入',
    191: '引擎-类型转换',
    192: '引擎-运行支撑',
    219: '引擎-序列化异常',
    243: '资料',
    148: '挂起评审'
}
# print(all_bug)
# today = datetime.datetime.today().strftime("%Y-%m-%d")

# 获取今天的日期
today = str(datetime.datetime.today().date())
print(today)
# today = "2024-01-18"

# 初始化计数器
new_p1_count = 0
new_p2_count = 0
new_p3_count = 0
new_p4_count = 0
closed_p1_count = 0
closed_p2_count = 0
closed_p3_count = 0
closed_p4_count = 0
fixed_p1_count = 0
fixed_p2_count = 0
fixed_p3_count = 0
fixed_p4_count = 0

# 遍历字典列表，计数满足条件的数量
for i in all_bug:

    title = i["title"][0].split(" ")[0]
    priority = i["pri"][0]
    # print(priority)
    # print(type(priority))
    status = i["status"][0]
    module = i["module"][0]
    activatedCount = i["activatedCount"][0]


    openedDate = i["openedDate"][0].split(" ")[0]
    # print(openedDate)
    if openedDate == today and priority == "1":
        new_p1_count += 1
        # print(title)
    elif openedDate == today and priority == "2":
        # print(i)
        new_p2_count += 1
    elif openedDate == today and priority == "3":
        # print(i)
        new_p3_count += 1
    elif openedDate == today and priority == "4":
        new_p4_count += 1

    closedDate = i["closedDate"][0].split(" ")[0]
    # print(closedDate)
    if closedDate == today and priority == "1":
        closed_p1_count += 1
    elif closedDate == today and priority == "2":
        # print(i)
        closed_p2_count += 1
    elif closedDate == today and priority == "3":
        # print(i)
        closed_p3_count += 1
    elif closedDate == today and priority == "4":
        closed_p4_count += 1

    resolvedDate = i["resolvedDate"][0].split(" ")[0]
    # print(resolvedDate)
    if resolvedDate == today and priority == "1":
        fixed_p1_count += 1
    elif resolvedDate == today and priority == "2":
        fixed_p2_count += 1
    elif resolvedDate == today and priority == "3":
        fixed_p3_count += 1
    elif resolvedDate == today and priority == "4":
        fixed_p4_count += 1

    if int(priority) <= 2 and status == "active":
        # print("-------------+++++++++++++----------\n", i)
        prionetwo = {
            "id": i["id"][0],
            "title": i["title"][0],
            "assignedTo": i["assignedTo"][0],
            "pri": i["pri"][0],
            "module": i["module"][0]
        }
        prionetwo_bugs.append(prionetwo)

    if int(activatedCount) >= 1 and status == "active":
        # print("-------------+++++++++++++----------\n", i)
        returnedbug = {
            "id": i["id"][0],
            "title": i["title"][0],
            "pri": i["pri"][0],
            "activatedCount": i["activatedCount"][0],
            "assignedTo": i["assignedTo"][0],
            "resolvedBy": i["resolvedBy"][0],
            "lastEditedBy": i["lastEditedBy"][0],
            "module": i["module"][0]
        }
        returned_bugs.append(returnedbug)

    # 下面这个很重要! 不指定获取数据为空
    bug_count_dict[status][module][priority] += 1

    # for status, module_count_dict in bug_count_dict.items():
    #     print(f"状态 {status}:")
    #     for module, priority_count_dict in module_count_dict.items():
    #         print(f"    模块 {module}:")
    #         for priority, count in priority_count_dict.items():
    #             print(f"    优先级 {priority}: {count} 个 bug")


# 打印结果
# print("今天日期新增 p1 的 bug 数量：", new_p1_count)
# print("今天日期新增 p2 的 bug 数量：", new_p2_count)
# print("今天日期新增 p3 的 bug 数量：", new_p3_count)
# print("今天日期新增 p4 的 bug 数量：", new_p4_count)
# print("今天日期关闭 p1 的 bug 数量：", closed_p1_count)
# print("今天日期关闭 p2 的 bug 数量：", closed_p2_count)
# print("今天日期关闭 p3 的 bug 数量：", closed_p3_count)
# print("今天日期关闭 p4 的 bug 数量：", closed_p4_count)
# print(f'今天日期关闭 p1 的 bug 数量:  {fixed_p1_count}个')
# print(f'今天日期关闭 p2 的 bug 数量:  {fixed_p2_count}个')
# print(f'今天日期关闭 p3 的 bug 数量:  {fixed_p3_count}个')
# print(f'今天日期关闭 p4 的 bug 数量:  {fixed_p4_count}个')
#
# print(prionetwo_bugs)
# print(returned_bugs)

