# coding=utf-8
from go_abroad_core import *

if (__name__ == "__main__"):
    global EN_map, CN_map, CN_key_list, EN_key_list
    result = read_xml_file(FILE_PATH_XML)
    CN_map = result[0]
    CN_key_list = result[1]
    generate_PM_xlsx_file(CN_map, CN_key_list)
