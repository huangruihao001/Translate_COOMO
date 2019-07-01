import requests
import json
import time
from translate import Translator

import os
import re
import shutil
import openpyxl
from openpyxl import Workbook

start=time.time()


# 翻译函数
def live(a):
    content=a
    chrome={"user-agent":"Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36"}
    url="https://fanyi.youdao.com/openapi.do?keyfrom=123licheng&key=1933182090&type=data&doctype=json&version=1.1&q={}".format(content)
    html=requests.get(url,headers=chrome)
    resqer=html.content
    load=json.loads(resqer)
    # print(load["translation"][0])
    return load["translation"][0] #返回翻译结果

# live()
# exit=time.time()
# print('[+]该程序耗时',exit-start)


def en_to_zh(a):
    translator = Translator(to_lang="chinese")
    translation = translator.translate(str(a))
    # print(translation)
    return str(translation)


def zh_to_en(a):
    translator = Translator(from_lang="chinese", to_lang="english")
    translation = translator.translate(str(a))
    # print(translation)
    return str(translation)


if __name__ == '__main__':
    # a = zh_to_en("测试")
    # print("最终结果：" + a)




    row = int(input("输入从第几行开始翻译："))
    # row_end = re_sheet.max_row # excel最大行数
    row_end = int(input("输入从第几行停止翻译："))

    while row <= row_end:
        re_workbook = openpyxl.load_workbook(".\楷模衣柜翻译内容-2019-06-13.xlsx")
        re_sheet = re_workbook["Sheet1"]
        zh_cell = "C" + str(row) # 中文字符所在单元格
        en_cell = "D" + str(row) # 英文字符所在单元格
        a = str(re_sheet[str(zh_cell)].value) # 待翻译的中文字符
        print("中文：" + str(a))
        b = zh_to_en(str(a)) # 翻译后的英文字符
        print("英文：" + b)
        re_sheet[str(en_cell)] = str(b)
        row += 1
        print("准备翻译第" + str(row) + "行")
        if "MYMEMORY WARNING:" in str(b):
            print("异常")
            re_workbook.save(".\楷模衣柜翻译内容-2019-06-13.xlsx")
            break
        else:
            re_workbook.save(".\楷模衣柜翻译内容-2019-06-13.xlsx")
    print("翻译完成")
