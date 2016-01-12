# coding=utf-8
import xlrd
from xml.dom.minidom import parse
import xml.dom.minidom
import openpyxl

CN_map = {}
EN_map = {}

FILE_PATH_XLSX = 'strings.xlsx'
FILE_PATH_XML = 'strings.xml'


def read_xlsx_file(file_path):
    wb = openpyxl.load_workbook(filename=file_path, use_iterators=True)
    sheet_names = wb.get_sheet_names()
    head_sheet = wb.get_sheet_by_name(sheet_names[0])
    it = head_sheet.iter_rows()
    for i in next(it):
        print i.value

    return


def write_to_xlsx_file(path):
    return


def read_xml_file(file_path):
    xml_tree = xml.dom.minidom.parse(file_path)
    resources = xml_tree.documentElement
    StringKeyMaps = resources.getElementsByTagName('string')

    for item in StringKeyMaps:
        key = item.getAttribute('name')
        value = item.childNodes[0].nodeValue
        CN_map[key] = value
        print key, '  :  ', value, '\n'

    return


def write_xml_file(path):
    return


def extract_kay_and_value_by_string(string):
    Key = 1
    value = 'v'
    return


def chinese_to_english_by_key(key, en_value):
    return


if (__name__ == "__main__"):
    # read_xml_file(FILE_PATH_XML)
    read_xlsx_file(FILE_PATH_XLSX)
