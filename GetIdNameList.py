import os
import re

# 指定文件夹路径
folder_path = 'go语言实验报告1'

# 获取文件列表
file_names = os.listdir(folder_path)

# 定义正则表达式模式
pattern = re.compile(r'(\d+)([^\d《]+)《')

# 遍历文件列表并提取学号和名字
data = []

# 遍历文件名列表并提取学号和名字
for file_name in file_names:
    match = pattern.search(file_name)
    if match:
        student_id, name = match.groups()
        name = name.strip().replace('+', "").replace('-', "")
        data.append([student_id, name])

# 格式化成适合粘贴到Excel的文本
# excel_text = "学号\t名字\n"
excel_text = ""
for entry in data:
    excel_text += f"{entry[0]}\t{entry[1]}\n"

# 输出结果
print(excel_text)
