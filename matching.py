import xlrd, json, openpyxl
from os.path import join, isfile
from os import listdir
from pprint import pprint as pp


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
LEAD_ROLES = {"CO-MANAGER", "LEAD MANAGER", "CO-LEAD MANAGER"}
LEAD = "LEAD"
PART = "PART"
ID = "ID"
N_PARTS = "N Parts"
N_LEADS = "N Leads"


def get_sheet(fname):
    return xlrd.open_workbook(fname).sheet_by_index(0)


def get_sheet_data(sheet):
    data = []
    for row in range(3, sheet.nrows):
        i = row - 3
        data.append({})
        data[i][ID] = i
        data[i][BORROWER[0]] = sheet.cell_value(rowx=row, colx=BORROWER[1])
        data[i][DATE[0]] = sheet.cell_value(rowx=row, colx=DATE[1])
        data[i][MANAGERS[0]] = list(zip(sheet.cell_value(rowx=row, colx=MANAGERS[1]).split("\n"), map(lambda x: LEAD if x in LEAD_ROLES else PART, sheet.cell_value(rowx=row, colx=DESCRS[1]).split("\n")), sheet.cell_value(rowx=row, colx=DESCRS[1]).split("\n")))
        data[i][N_PARTS] = sum(p[1] == PART for p in data[i][MANAGERS[0]])
        data[i][N_LEADS] = sum(p[1] == LEAD for p in data[i][MANAGERS[0]])
        
    return data


def all_loans_data(n = None):
    if n is None:
        n = len(loans_sheets_paths)
    data = []
    for i in range(n):
        data += get_sheet_data(get_sheet(loans_sheets_paths[i]))
        print("File " + str(i + 1) + " of " + str(len(loans_sheets_paths)) + " loaded")
        
    return data


def all_role_descriptions(loans_data):
    descrs = set()
    for d in loans_data:
        for m in d[MANAGERS[0]]:
            descrs.add(m[1])
    return descrs


def make_sheet(loans):
    wb = openpyxl.Workbook()
    wb.name = "Loans data"
    sheet = wb.active
    sheet.title = "Loans data"

    max_leads = max(l[N_LEADS] for l in loans)
    max_parts = max(l[N_PARTS] for l in loans)

    sheet.cell(row = 1, column = 1).value = ID
    sheet.cell(row = 1, column = 2).value = "Date"

    for i in range(max_leads):
        sheet.cell(row = 1, column = i + 3).value  = "Lead " + str(i + 1)

    for i in range(max_parts):
        sheet.cell(row = 1, column = i + 3 + max_leads).value =  "Part " + str(i + 1)

    for y, loan in enumerate(loans):
        if (y % 10000 == 0):
            print(str(y) + " loans written of " + str(len(loans)))
        sheet.cell(row = y + 2, column = 1).value = loan[ID]
        sheet.cell(row = y + 2, column = 2).value = loan[DATE[0]]

        for i, lead in enumerate(filter(lambda x: x[1] == LEAD, loan[MANAGERS[0]])):
            sheet.cell(row = y + 2, column = i + 3).value = lead[0]

        for i, part in enumerate(filter(lambda x: x[1] == PART, loan[MANAGERS[0]])):
            sheet.cell(row = y + 2, column = i + 3 + max_leads).value = part[0]
            
        
    wb.save("Loans.xlsx")
