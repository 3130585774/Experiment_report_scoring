import os

import pandas as pd

file_path = '2class.xlsx'


def excel_to_dict(file_path, sheet_name="Sheet"):
    # 读取 Excel 文件
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # 将 DataFrame 转换为字典
    data_dict = df.to_dict(orient='records')
    new_dict = {}
    for item in data_dict:
        new_dict[item["学号"]] = item["姓名"]
    return new_dict


def find_key_value_pairs(my_string, my_dict):
    result = {}
    for key, value in my_dict.items():
        if str(key) in my_string:
            result[key] = value
    if result == {}:
        print(my_string)
        return None
    key, value = next(iter(result.items()))
    return key, value


def list_subdirectories(directory):
    subdirectories = [d for d in os.listdir(directory) if os.path.isdir(os.path.join(directory, d))]
    return subdirectories


def list_doc_files(directory):
    doc_files = [f for f in os.listdir(directory) if f.endswith('.docx') and os.path.isfile(os.path.join(directory, f))]
    return doc_files


def renamefile(old, new):
    os.rename(old, new)


className = "2班"

Name_dict = excel_to_dict(file_path)

lessons = list_subdirectories(className)

for lesson in lessons:
    lessonNumbers = list_subdirectories(className + "/" + lesson)
    for little_lesson in lessonNumbers:
        Nowdir = className + "/" + lesson + "/" + little_lesson
        file_list = list_doc_files(Nowdir)
        for file_name in file_list:
            # reName
            id_and_name = find_key_value_pairs(file_name, Name_dict)
            if id_and_name is None:
                print(id_and_name[0], id_and_name[1])
                continue
            newName = f"{id_and_name[0]}+{id_and_name[1]}+《{little_lesson[:-1]}》实验报告{little_lesson[-1:]}.docx"
            renamefile(Nowdir + "/" + file_name, Nowdir + "/" + newName)
            # print(file_name, "\nv\n", newName, "\n")

for lesson in lessons:
    lessonNumbers = list_subdirectories(className + "/" + lesson)
    for little_lesson in lessonNumbers:
        Nowdir = className + "/" + lesson + "/" + little_lesson
        file_list = list_doc_files(Nowdir)
        for file_name in file_list:
            pass
