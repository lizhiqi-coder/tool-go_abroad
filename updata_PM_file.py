# coding=utf-8
import openpyxl
import copy
from go_abroad_core import *

OLD_FILE_PATH_XLSX = 'strings_old.xlsx'
NEW_FILE_PATH_XLSX = 'strings.xlsx'

global OLD_CN_MAP
global OLD_EN_MAP
global OLD_KEY_LIST

global NEW_CN_MAP
global NEW_KEY_LIST

global wb
global sheet

# global old_result
# global new_result

OLD_CN_MAP = {}
OLD_EN_MAP = {}
OLD_KEY_LIST = []

NEW_CN_MAP = []
NEW_KEY_LIST = []


# 新生成的PM 文件没有任何翻译成英文，旧PM 文件翻译的信息存入新文件中，使用新文件给PM 编辑
# 既然是把旧文件内容同步到新文件中，则以旧文件的key_list 为准
# 若旧文件的 一个key 在新文件中不存在，则不写入
def write_old_file_to_new(old_file, new_file):
    global OLD_CN_MAP
    global OLD_EN_MAP
    global OLD_KEY_LIST

    global NEW_CN_MAP
    global NEW_KEY_LIST

    global wb
    global sheet

    OLD_CN_MAP = {}
    OLD_EN_MAP = {}
    OLD_KEY_LIST = []

    NEW_CN_MAP = []
    NEW_KEY_LIST = []

    old_result = copy.deepcopy(read_xlsx_file(old_file))
    OLD_CN_MAP = old_result[0]
    OLD_EN_MAP = old_result[1]
    OLD_KEY_LIST = old_result[2]

    new_result = copy.deepcopy(read_xlsx_file(new_file))
    NEW_CN_MAP = new_result[0]
    NEW_KEY_LIST = new_result[2]

    wb = openpyxl.load_workbook(filename=NEW_FILE_PATH_XLSX)
    sheet = wb.worksheets[0]

    print 'OLD_KEY_LIST:', OLD_KEY_LIST
    print 'NEW_KEY_LIST:', NEW_KEY_LIST

    for key in OLD_KEY_LIST:

        if key not in NEW_KEY_LIST:
            continue

        row = get_key_row_in_xlsx(key, sheet)

        if row == None:
            continue

        EN_cell_position = EN_VALUE_COLUMN_IN_XLSX + str(row)

        sheet[EN_cell_position] = OLD_EN_MAP[key]
    # write_to_EN_cell_by_key()
    wb.save(NEW_FILE_PATH_XLSX)
    return


def write_to_EN_cell_by_key():
    global OLD_KEY_LIST
    global NEW_KEY_LIST
    global wb
    global sheet

    for key in OLD_KEY_LIST:

        if key not in NEW_KEY_LIST:
            continue

        row = get_key_row_in_xlsx(key, sheet)

        if row == None:
            continue

        EN_cell_position = EN_VALUE_COLUMN_IN_XLSX + str(row)

        sheet[EN_cell_position] = OLD_EN_MAP[key]

    # wb.save(NEW_FILE_PATH_XLSX)
    return


def get_key_row_in_xlsx(key, sheet):
    for i in range(1, sheet.max_row):
        key_position = ANDROID_KEY_COLUMN_IN_XLSX + str(i)
        if sheet[key_position].value == key:
            return i
        else:
            continue


if (__name__ == '__main__'):
    write_old_file_to_new(OLD_FILE_PATH_XLSX, NEW_FILE_PATH_XLSX)
