# -*- coding:utf-8 -*-
import os
from openpyxl import load_workbook
from openpyxl import Workbook

# 创建一个新的工作簿
wb = Workbook()
# 获取默认的工作表
sheet = wb.active
# 设置表头
sheet.append(['学号', '姓名'])
classnum = 3
# 文件夹路径
folder_path = f'./{classnum}/1'

# 遍历文件夹中的所有文件
for filename in os.listdir(folder_path):
    # 仅处理文件名以'.txt'结尾的文件
    if filename.endswith('.docx'):
        # 解析文件名，获取学号和姓名
        parts = filename[:-18].split('+')  # 去掉'.txt'后缀并使用下划线分隔
        student_id = parts[0]
        student_name = parts[1]

        # 向工作表中写入数据
        sheet.append([student_id, student_name])

# 保存工作簿
wb.save(f'{classnum}ban.xlsx')