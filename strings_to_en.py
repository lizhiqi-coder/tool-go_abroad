# coding=utf-8
import os
from xml.dom import minidom
from xml.dom.minidom import parse

import openpyxl
from openpyxl.writer.excel import ExcelWriter

CN_map = {}
EN_map = {}

FILE_PATH_XLSX = 'strings.xlsx'
FILE_PATH_XML = 'strings.xml'
GENERATE_ANDROID_STRINGS_FILE = "strings_en.xml"

ANDROID_KEY_COLUMN_IN_XLSX = 'B'
CN_VALUE_COLUMN_IN_XLSX = 'C'
EN_VALUE_COLUMN_IN_XLSX = 'D'

FIRST_LINE_IN_XLSX = '1'
LAST_LINE_IN_XLSX = '50'

CONTENT_ARANGE_IN_XLSX = ANDROID_KEY_COLUMN_IN_XLSX + FIRST_LINE_IN_XLSX + ':' + EN_VALUE_COLUMN_IN_XLSX + LAST_LINE_IN_XLSX


def read_xlsx_file(file_path):
    cleanBuffer()
    wb = openpyxl.load_workbook(filename=file_path, use_iterators=True)
    sheet_names = wb.get_sheet_names()
    head_sheet = wb.get_sheet_by_name(sheet_names[0])
    # LAST_LINE_IN_XLSX = head_sheet.max_row
    print CONTENT_ARANGE_IN_XLSX
    for row in head_sheet.iter_rows(CONTENT_ARANGE_IN_XLSX):
        key = row[0].value
        CN_value = row[1].value
        EN_value = row[2].value
        CN_map[key] = CN_value.encode('gb2312')
        EN_map[key] = EN_value.encode('gb2312')
    return CN_map, EN_map


def write_to_xlsx_file(content, file_path):
    if os.path.exists(file_path) == False:
        createFile(file_path)

    wb = openpyxl.Workbook()
    ew = ExcelWriter(workbook=wb)
    wsheet = wb.worksheets[0]

    is_close = raw_input('please be sure to close your xlsx file (y/n)?')
    if is_close != 'y':
        print '未关闭文件，退出程序'
        return

    rows = [row for row in wsheet.iter_rows(CONTENT_ARANGE_IN_XLSX)]
    rows[0][0].value = 'android_key'
    rows[0][1].value = 'CN_String_value'
    rows[0][2].value = 'EN_String_value'

    keys = [key for key in content]
    for i in range(0, keys.__len__()):
        rows[i + 1][0].value = keys[i]
        rows[i + 1][1].value = content[keys[i]]

        # print rows[i + 1][0].value, rows[i + 1][1].value

    ew.save(file_path)
    return


def read_xml_file(file_path):
    cleanBuffer()
    xml_tree = parse(file_path)
    resources = xml_tree.documentElement
    StringKeyMaps = resources.getElementsByTagName('string')

    for item in StringKeyMaps:
        key = item.getAttribute('name')
        value = item.childNodes[0].nodeValue
        CN_map[key] = value
        # print key, '  :  ', value, '\n'

    return CN_map


def write_xml_file(content, file_path):
    if os.path.exists(file_path) == False:
        createFile(file_path)

    imp = minidom.getDOMImplementation()
    dom = imp.createDocument(None, None, None)
    root = dom.createElement('resourse')

    for key, value in content.items():
        item = dom.createElement('string')
        item_value = dom.createTextNode(value)
        item.setAttribute('name', key)
        item.appendChild(item_value)

        root.appendChild(item)

        continue
    dom.appendChild(root)

    file_handle = open(file_path, 'w')
    dom.writexml(file_handle, '', '', '\n', None)
    file_handle.close()

    return


def generate_android_strings_file(content):
    write_xml_file(content, GENERATE_ANDROID_STRINGS_FILE)
    return


def cleanBuffer():
    EN_map.clear()
    CN_map.clear()
    return


def createFile(file_name):
    f = open(file_name, 'w')
    f.close()
    return


# content = {'one': '1', 'two': '2', 'three': '3'}

if (__name__ == "__main__"):
    # read_xml_file(FILE_PATH_XML)
    read_xlsx_file(FILE_PATH_XLSX)
    # write_to_xlsx_file(CN_map, FILE_PATH_XLSX)
    print 'this is cn_map', CN_map
    generate_android_strings_file(CN_map)
