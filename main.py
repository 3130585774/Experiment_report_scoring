from ChanageName import *
from SetScore import *

data_dict = getScoreDict(f"{className}.xlsx")

for lesson in list_subdirectories(className):

    lessonNumbers = list_subdirectories(className + "/" + lesson)
    for little_lesson in lessonNumbers:
        Nowdir = className + "/" + lesson + "/" + little_lesson
        file_list = list_doc_files(Nowdir)
        for file in file_list:
            s = getScoreByFileName(data_dict, little_lesson, file)
            setScore(f"{Nowdir}/{file}", s)
