# coding=utf-8
from strings_to_en import *

if (__name__ == "__main__"):
    read_xlsx_file(FILE_PATH_XLSX)
    generate_android_strings_file(EN_map, CN_key_list, CN_map)
