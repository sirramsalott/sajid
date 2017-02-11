import xlrd
from os.path import join, isfile
from os import listdir
import json


document_home = "/home/joe/sajid/"
acquisitions_path = join(document_home, "M&As Data/Zephyr_Export_3.xls")
loans_path = join(document_home, "Syndicated Loan Data/")
loans_sheets_paths = []


for subdir in ["GlobalminusUSUK", "UK", "US"]:
    p = join(loans_path, subdir)
    loans_sheets_paths += [join(p, f) for f in listdir(p) if isfile(join(p, f)) and not "~" in f and not "rev" in f]


BORROWER = ("Borrower", 0)
DATE = ("Announcement Date", 15)
DESCRS = ("Role Descriptions", 44)
MANAGERS = ("All Managers", 45)


class Loan(object):
    def __init__(self, date, borrower, role_descrs, managers):
        self.date = date
        self.borrower = borrower
        self.role_descrs = role_desrcs
        self.managers = managers


def get_sheet(fname):
    return xlrd.open_workbook(fname).sheet_by_index(0)


def get_sheet_data(sheet):
    data = []
    for row in range(3, sheet.nrows):
        i = row - 3
        data.append({})
        data[i][BORROWER[0]] = sheet.cell_value(rowx=row, colx=BORROWER[1])
        data[i][DATE[0]] = sheet.cell_value(rowx=row, colx=DATE[1])
        data[i]["Managers"] = list(zip(sheet.cell_value(rowx=row, colx=MANAGERS[1]).split("\n"), sheet.cell_value(rowx=row, colx=DESCRS[1]).split("\n")))
    return data


def all_loans_data():
    data = []
    for i in range(len(loans_sheets_paths)):
        data += get_sheet_data(get_sheet(loans_sheets_paths[i]))
        print("File " + str(i) + " of " + str(len(loans_sheets_paths)) + " loaded")
        
    return data


def all_role_descriptions(loans_data):
    descrs = set()
    for d in loans_data:
        for m in d["Managers"]:
            descrs.add(m[1])
    return descrs
