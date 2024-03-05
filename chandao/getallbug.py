# -*-coding: Utf-8 -*-
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from _datetime import datetime
import json
import chardet
import re
import ast

curr_time = datetime.now()
timestamp = datetime.strftime(curr_time, '%Y%m%d-%H%M%S')

# url = 'http://123.249.89.129:8888/zentao/api.php/v1'
url = 'http://zentao.jiuqingsk.com/zentao/api.php/v1'
# 页面上是html 改成json
# getbug_url = 'http://zentao.jiuqingsk.com/zentao/bug-browse-4-all-all.json'
getbug_url = 'http://zentao.jiuqingsk.com/zentao/bug-browse-4-all-all-0--1434-2000-1.json'
# getbug_url = 'http://192.168.2.21/zentao/bug-browse-4-all-bySearch-myQueryID.json'

def get_token():

    token_url = '/tokens'
    # 登录接口
    headers = {'Content-Type': 'application/json'}
    # login_params = {'_': '1699433580680'}
    data = {"account": "zhangyuxiong", "password": "Zyx123456"}
    token_response = requests.post(f"{url}{token_url}", headers=headers, json=data, verify=False)
    # print(token_response.text)
    token = token_response.json().get('token')
    # print(token)
    return token

def get_products():

    products_url = '/products'
    token = get_token()
    headers = {'Token': f'{token}', 'Content-Type': 'application/json'}
    products_response = requests.post(f"{url}{products_url}", headers=headers, verify=False)
    # print(products_response.text)
    products = products_response.json().get('products')
    # print(products)
    return products


def get_bugs():

    token = get_token()
    headers = {'Token': f'{token}', 'Content-Type': 'application/json'}
    # print(headers)
    response = requests.get(f"{getbug_url}", headers=headers, verify=False)
    # a = response.text.encode('ascii').decode('unicode_escape').encode('utf-8').decode('unicode_escape')
    # print(type(a))
    # print(a)
    b = response.text.encode('ascii').decode('unicode_escape').encode('utf-8').decode('unicode_escape').replace('\\/', '/')
    # print(type(b))
    # print(b)
    pattern = r'"bugs":\[(.*?)\],\s*"users"'
    # matches = re.findall(pattern, a, re.DOTALL)
    matches = re.findall(pattern, b, re.DOTALL)
    # print("-------------------\n", matches[0])
    # print(type(matches[0]))
    data = matches[0]
    patterns = {
        'id': r'"id":\s*"(\d+)"',
        'module': r'"module":\s*"(\d+)"',
        'title': r'"title":\s*"([^"]+)"',
        'severity': r'"severity":\s*"(\d+)"',
        'pri': r'"pri":\s*"(\d+)"',
        'activatedCount': r'"activatedCount":\s*"(\d+)"',
        'status': r'"status":\s*"([^"]+)"',
        'assignedTo': r'"assignedTo":\s*"([^"]+)"',
        'openedDate': r'"openedDate":\s*"([^"]+)"',
        'assignedDate': r'"assignedDate":\s*"([^"]+)"',
        'resolvedDate': r'"resolvedDate":\s*"([^"]+)"',
        'resolvedBy': r'"resolvedBy":\s*"([^"]+)"',
        'lastEditedBy': r'"lastEditedBy":\s*"([^"]+)"',
        'closedDate': r'"closedDate":\s*"([^"]+)"',
        'feedback': r'"feedback":\s*"(\d+)"'
    }
    # 使用re.findall()获取所有匹配结果
    # matches = re.findall(r'{(.*?)}', data, re.DOTALL) # 如果匹配所有以 { 开头、以 } 结尾的字符串，并且字符串中间不包含 } 字符的子串。这个表达式从数据中提取了多个 JSON 对象，但由于每个 JSON 对象的键值对数量和内容可能不同，因此在提取字段时可能会出现一些缺省值或错误值。

    matches = re.findall(r'{(.*?"needconfirm":false)}', data, re.DOTALL)
    # 过滤掉没有"id"字段的数据
    filtered_matches = [match for match in matches if re.search(patterns['id'], match)]
    # 逐个匹配并打印结果
    all_result = []
    for match in filtered_matches:
        result = {}
        for field, pattern in patterns.items():
            match_fields = re.findall(pattern, match, re.DOTALL)
            if match_fields:
                result[field] = match_fields
            else:
                result[field] = "Not found"
        all_result.append(result)
    # all_result = []
    # for match in filtered_matches:
    #     result = {}
    #     for field, pattern in patterns.items():
    #         match_field = re.search(pattern, match)
    #         if match_field:
    #             result[field] = match_field.group(1)
    #         else:
    #             result[field] = "Not found"
    #     all_result.append(result)
    #     print(result)
    # print(all_result)
    return all_result


if __name__ == "__main__":

    # get_products()
    get_bugs()