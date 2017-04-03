import xlrd, json, openpyxl
from string import punctuation
from os.path import join, isfile
from os import listdir
from pprint import pprint as pp
from functools import reduce
from nltk import edit_distance
from numpy import array
from collections import Counter
from company_name_similarity import CompanyNameSimilarity


document_home = "/home/joe/sajid/"
acquisitions_path = join(document_home, "M&As Data/Zephyr_Export_3.xls")
loans_path = join(document_home, "Syndicated Loan Data/")
loans_sheets_paths = []
cm = CompanyNameSimilarity()


with open("wc.txt") as f:
    places = f.read().split("\n")


for subdir in ["GlobalminusUSUK", "UK", "US"]:
    p = join(loans_path, subdir)
    loans_sheets_paths += [join(p, f) for f in listdir(p)
                           if isfile(join(p, f)) and not "~" in f and not "rev" in f]


BORROWER = ("BORROWER", 0)
DATE = ("ANNOUNCEMENT DATE", 15)
DESCRS = ("ROLE DESCRIPTIONS", 44)
MANAGERS = ("ALL MANAGERS", 45)
BOOKRUNNERS = ("BOOKRUNNERS", 41)
MANDATED_MANAGERS = ("MANDATED MANAGERS", 54)
LEAD_ROLES = {"CO-MANAGER", "LEAD MANAGER", "CO-LEAD MANAGER"}
LEAD = "LEAD"
PART = "PART"
ID = "ID"
N_PARTS = "N Parts"
N_LEADS = "N Leads"
NORMALISED_MANAGERS = "NORMALISED_MANAGERS"
ACQUIROR_NAME = ("ACQUIROR NAME", 2)
TARGET_NAME = ("TARGET NAME", 4)
ACQUISITION_DATE = ("COMPLETED DATE", 13)
DEAL_STATUS = ("DEAL STATUS", 7)
MATCHES = "MATCHES"
THRESH = 0.01
STOPWORDS = ["bank", "sa", "ltd", "inc", "plc", "of", "the", "ag", "oao", "group", "de", "do", "di", "and", "banking", "banca"]
THRESH = 0.2


class Loan(object):

    matches = []
    
    def __init__(self, num, date, leads, parts, borrower, raw_leads, raw_parts):
        self.num = num
        self.date = date
        self.leads = leads
        self.parts = parts
        self.borrower = borrower
        self.raw_leads = raw_leads
        self.raw_parts = raw_parts

    def __repr__(self):
        return str(self.__dict__)


class Acquisition(object):
    def __init__(self, num, acquiror, target, date, status, acquiror_set, target_set):
        self.num = num
        self.acquiror = acquiror
        self.target = target
        self.date = date
        self.status = status
        self.acquiror_set = acquiror_set
        self.target_set = target_set

    def __repr__(self):
        return str(self.__dict__)


def get_sheet(fname):
    return xlrd.open_workbook(fname).sheet_by_index(0)


def normalise(name):
    #return [x for x in name.lower().translate(str.maketrans("", "", punctuation)).split(" ")]
    return [x for x in name.lower().translate(str.maketrans("", "", punctuation)).split(" ") if x not in STOPWORDS]


def get_acquisitions_data():
    sheet = xlrd.open_workbook(acquisitions_path).sheet_by_index(1)    
    data = []

    for row in range(1, sheet.nrows):
        data.append(Acquisition(sheet.cell_value(rowx=row, colx=1),
                                sheet.cell_value(rowx=row, colx=ACQUIROR_NAME[1]),
                                sheet.cell_value(rowx=row, colx=TARGET_NAME[1]),
                                sheet.cell_value(rowx=row, colx=ACQUISITION_DATE[1]),
                                sheet.cell_value(rowx=row, colx=DEAL_STATUS[1]),
                                set(normalise(sheet.cell_value(rowx=row, colx=ACQUIROR_NAME[1]))),
                                set(normalise(sheet.cell_value(rowx=row, colx=TARGET_NAME[1])))))

    return array(data)


def jaccard(a, b):
    return len(a.intersection(b)) / len(a.union(b))


def a_stopwords(n):
    acs = get_acquisitions_data()
    c = Counter()
    for a in acs:
        for word in a.target_set:
            c[word] += 1
        for word in a.acquiror_set:
            c[word] += 1
    return [x[0] for x in c.most_common(n)]


def similarity_score(s1, s2):
    dif = s1.symmetric_difference(s2)
    if len(dif) > 0 and all({x in places for x in dif}):
        return 0
    return cm.match_score(" ".join(s1), " ".join(s2))


def compare_acquisitions():
    wb = openpyxl.Workbook()
    wb.name = "Name comparison"
    acs = get_acquisitions_data()[1000:]
    sheet = wb.active
    sheet.title = "Name comparison"
    i = 0
    cm = CompanyNameSimilarity()
    matches = set()
    
    for a in acs:
        ja = " ".join(a.acquiror_set)
        
        for b in acs:
            jb = " ".join(b.acquiror_set)    
            match = (a.acquiror, b.acquiror)
            j = similarity_score(a.acquiror_set, b.acquiror_set)

            if 0.51 < j < 1 and match not in matches:
                matches.add(match)
                i += 1
                sheet.cell(row = i, column = 1).value = a.acquiror
                sheet.cell(row = i, column = 2).value = ja
                sheet.cell(row = i, column = 3).value = b.acquiror
                sheet.cell(row = i, column = 4).value = jb
                sheet.cell(row = i, column = 5).value = j
                sheet.cell(row = i, column = 6).value = " ".join(a.acquiror_set.symmetric_difference(b.acquiror_set))
            
    wb.save("Names.xlsx")


def is_lead(manager):
    return manager[1] in LEAD_ROLES


def partition(xs, p):
    return ([x for x in xs if p(x)], [x for x in xs if not p(x)])


def get_sheet_data(sheet, n, acquisitions, lmax):
    data = []
    for row in range(3, min(lmax + 3, sheet.nrows)):
        i = row - 3
        print(i)
        borrower = sheet.cell_value(rowx=row, colx=BORROWER[1])
        date = sheet.cell_value(rowx=row, colx=DATE[1])
        managers = list(zip(sheet.cell_value(rowx=row, colx=MANAGERS[1]).split("\n"),
                            sheet.cell_value(rowx=row, colx=DESCRS[1]).split("\n")))
        (leads, parts) = partition(managers, lambda x: is_lead(x))
        loan = Loan(n + i, date, [(normalise(l[0]), l[1]) for l in leads], [(normalise(p[0]), p[1]) for p in parts], borrower, [l[0] for l in leads], [p[0] for p in parts])
        
        ms = [normalise(m[0]) for m in managers]
        
        for a in acquisitions:
            for m1 in ms:
                sm1 = set(m1)
                for m2 in ms:
                    sm2 = set(m2)
                    if sm1 != sm2 and (jaccard(sm1, a.acquiror_set) < THRESH and jaccard(sm2, a.target_set) < THRESH) or (jaccard(sm2, a.acquiror_set) < THRESH and jaccard(sm1, a.target_set) < THRESH):
                            loan.matches.append((m1, m2))
        data.append(loan)
        
        #data[i][BOOKRUNNERS[0]] = sheet.cell_value(rowx=row, colx=BOOKRUNNERS[1]).split("\n")
        #data[i][MANDATED_MANAGERS[0]] = sheet.cell_value(rowx=row, colx=MANDATED_MANAGERS[1]).split("\n")
        #data
        #data[i][NORMALISED_MANAGERS] = [normalise(m[0]) for m in data[i][MANAGERS[0]]]

    return (data, n + i)


def all_loans_data(n = None):
    a = get_acquisitions_data()
    if n is None:
        n = sys.maxsize
    data = []
    tot = 0
    i = 0
    while i < len(loans_sheets_paths) and n > 0:
        (temp, tot) = get_sheet_data(get_sheet(loans_sheets_paths[i]), tot, a, n)
        data += temp
        n -= tot
        print("File " + str(i + 1) + " of " + str(len(loans_sheets_paths)) + " loaded")
        i += 1
        
    return data


def all_role_descriptions(loans_data):
    descrs = set()
    for d in loans_data:
        for m in d[MANAGERS[0]]:
            descrs.add(m[1])
    return descrs
            

def compare_managers(loans):
    wb = openpyxl.Workbook()
    wb.name = "Managers"
    sheet = wb.active
    sheet.title = "Managers"

    sheet.cell(row = 1, column = 1).value = ID
    sheet.cell(row = 1, column = 2).value = MANAGERS[0]
    sheet.cell(row = 1, column = 3).value = "ROLE"
    sheet.cell(row = 1, column = 4).value = "IDENTIFIED ROLE"
    sheet.cell(row = 1, column = 5).value = "BOOKRUNNER"
    sheet.cell(row = 1, column = 6).value = "MANDATED MANAGER"

    y = 2

    for loan in loans:
        sheet.cell(row = y, column = 1).value = loan[ID]
        
        for m in loan[MANAGERS[0]]:
            sheet.cell(row = y, column = 2).value = m[0]
            sheet.cell(row = y, column = 3).value = m[2]
            sheet.cell(row = y, column = 4).value = m[1]
            sheet.cell(row = y, column = 5).value = str(m[0] in loan[BOOKRUNNERS[0]])
            sheet.cell(row = y, column = 6).value = str(m[0] in loan[MANDATED_MANAGERS[0]])
            y += 1

    wb.save("Managers.xlsx")


def make_sheet(loans):
    wb = openpyxl.Workbook()
    wb.name = "Loans data"
    sheet = wb.active
    sheet.title = "Loans data"

    max_leads = max(len(l.raw_leads) for l in loans)
    max_parts = max(len(l.raw_parts) for l in loans)

    sheet.cell(row = 1, column = 1).value = ID
    sheet.cell(row = 1, column = 2).value = DATE[0]

    for i in range(max_leads):
        sheet.cell(row = 1, column = i + 3).value  = "Lead " + str(i + 1)

    for i in range(max_parts):
        sheet.cell(row = 1, column = i + 3 + max_leads).value =  "Part " + str(i + 1)

    for y, loan in enumerate(loans):
        sheet.cell(row = y + 2, column = 1).value = loan.num
        sheet.cell(row = y + 2, column = 2).value = loan.date

        for i, lead in enumerate(loan.raw_leads):
            sheet.cell(row = y + 2, column = i + 3).value = lead

        for i, part in enumerate(loan.raw_parts):
            sheet.cell(row = y + 2, column = i + 3 + max_leads).value = part
        
    wb.save("Loans.xlsx")
