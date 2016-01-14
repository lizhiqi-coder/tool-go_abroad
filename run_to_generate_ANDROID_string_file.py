# coding=utf-8
from go_abroad_core import *
import copy

if (__name__ == "__main__"):
    global EN_map, CN_map, CN_key_list, EN_key_list
    result = copy.deepcopy(read_xlsx_file(FILE_PATH_XLSX))
    CN_map = result[0]
    EN_map = result[1]
    CN_key_list = result[2]
    generate_android_strings_file(EN_map, CN_key_list, CN_map)
