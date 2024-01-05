from ChanageName import *
from SetScore import *

data_dict = getScoreDict(f"{className}.xlsx")

# for lesson in list_subdirectories(className):
#
#     lessonNumbers = list_subdirectories(className + "/" + lesson)
#     for little_lesson in lessonNumbers:
#         Nowdir = className + "/" + lesson + "/" + little_lesson
#         file_list = list_doc_files(Nowdir)
#         for file in file_list:
#             s = getScoreByFileName(data_dict, little_lesson, file)
#             setScore(f"{Nowdir}/{file}", s)

for i in range(1, 8):
    dir_name = f"go语言实验报告{i}"
    file_names = list_doc_files(dir_name)
    for file_name in file_names:
        s = getScoreByFileName(data_dict, dir_name, file_name)
        setScore(f"{dir_name}/{file_name}", s)

