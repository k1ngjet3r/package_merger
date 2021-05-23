from openpyxl import load_workbook, Workbook
import json

def open_json(filename):
    # For Windows
    # with open('json\\' + filename) as f:
    # For MacOS
    with open('json/' + filename) as f:
        return json.load(f)

conflict_dict = open_json('pkg_conflict.json')
pkg_def_dict = open_json('pkg_def.json')

