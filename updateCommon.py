#
# updateCommon.py -- common methods and variables for updates
#
# @author Mark Sattolo <epistemik@gmail.com>
# @version Python 3.6
# @created 2019-04-07
# @updated 2019-04-07

from sys import stdout
from datetime import date, timedelta
from bisect import bisect_right
from decimal import Decimal
from math import log10
import csv
from gnucash import GncNumeric
import json

# constant strings
QTR   = 'Quarter'
YR    = 'Year'
REV   = 'Revenue'
INV   = 'Invest'
OTH   = 'Other'
SAL   = 'Salary'
BAL   = 'Balance'
CONT  = 'Contingent'
NEC   = 'Necessary'
DEDNS = 'Sal_Dedns'
AU    = 'Gold'
AG    = 'Silver'
CASH  = 'Cash'
BANK  = 'Bank'
RWRDS = 'Rewards'
OPEN  = 'OPEN'
RRSP  = 'RRSP'
TFSA  = 'TFSA'
HOUSE = 'House'

# find the proper path to the accounts in the gnucash file
REV_ACCTS = {
    INV : ["REV_Invest"],
    OTH : ["REV_Other"],
    SAL : ["REV_Salary"]
}
EXP_ACCTS = {
    BAL   : ["EXP_Balance"],
    CONT  : ["EXP_CONTINGENT"],
    NEC   : ["EXP_NECESSARY"]
}
DEDN_ACCTS = {
    "Mark" : ["Mk-Dedns"],
    "Lulu" : ["Lu-Dedns"],
    "ML"   : ["ML-Dedns"]
}

# find the proper path to the accounts in the gnucash file
ASSET_ACCTS = {
    AU    : ["FAMILY", "Precious Metals", "Au"],
    AG    : ["FAMILY", "Precious Metals", "Ag"],
    CASH  : ["FAMILY", "LIQUID", "$&"],
    BANK  : ["FAMILY", "LIQUID", BANK],
    RWRDS : ["FAMILY", RWRDS],
    OPEN  : ["FAMILY", "INVEST", OPEN],
    RRSP  : ["FAMILY", "INVEST", RRSP],
    TFSA  : ["FAMILY", "INVEST", TFSA],
    HOUSE : ["FAMILY", HOUSE]
}

# either for One Quarter or for Four Quarters if updating an entire Year
results = list()
# store the values needed to update the document
REV_EXP_RESULTS = {
    QTR   : '0',
    REV   : '0',
    BAL   : '0',
    CONT  : '0',
    NEC   : '0',
    DEDNS : '0'
}
# store the values needed to update the document
ASSET_RESULTS = {
    QTR   : '0',
    AU    : '0',
    AG    : '0',
    CASH  : '0',
    BANK  : '0',
    RWRDS : '0',
    OPEN  : '0',
    RRSP  : '0',
    TFSA  : '0',
    HOUSE : '0'
}

# number of months in the period
PERIOD_QTR = 3

NUM_MONTHS = 12
ONE_DAY = timedelta(days=1)
ZERO = Decimal(0)

CREDENTIALS = 'secrets/credentials.json'

SHEETS_RW_SCOPE = ['https://www.googleapis.com/auth/spreadsheets']

SHEETS_EPISTEMIK_RW_TOKEN = {
    'P2' : 'secrets/token.sheets.epistemik.rw.pickle2' ,
    'P3' : 'secrets/token.sheets.epistemik.rw.pickle3' ,
    'P4' : 'secrets/token.sheets.epistemik.rw.pickle4'
}

# Spreadsheet ID
BUDGET_QTRLY_SPRD_SHEET = '1YbHb7RjZUlA2gyaGDVgRoQYhjs9I8gndKJ0f1Cn-Zr0'
# sheet names in Budget Quarterly
ALL_INC_SHEET       = 'All Inc Quarterly'
ALL_INC_PRAC_SHEET  = 'All Inc Practice'
NEC_INC_SHEET       = 'Nec Inc Quarterly'
NEC_INC_PRAC_SHEET  = 'Nec Inc Practice'
BALANCE_SHEET       = 'Balance Sheet'
QTR_ASTS_SHEET      = 'Quarterly Assets'
QTR_ASTS_PRAC_SHEET = 'Qtrly Assets Practice'
ML_WORK_SHEET       = 'ML Work'
CALCULNS_SHEET      = 'Calculations'

TOKEN = SHEETS_EPISTEMIK_RW_TOKEN['P4']

# for One Quarter or for Four Quarters if updating an entire Year
data = list()
cell_data = {
    # sample data
    'range': 'Calculations!P47',
    'values': [ [ '$9033.66' ] ]
}

# base cell (Q1) locations in Budget-qtrly.gsht
BASE_ROW = 3

REV_EXP_COLS = {
    REV   : 'D',
    BAL   : 'P',
    CONT  : 'O',
    NEC   : 'G',
    DEDNS : 'D'
}
RE_BASE_YEAR = 2012
# number of rows between quarters in the same year
RE_QTR_SPAN = 2
# number of rows between years
RE_YEAR_SPAN = 11

ASSET_COLS = {
    AU    : 'U',
    AG    : 'T',
    CASH  : 'R',
    BANK  : 'Q',
    RWRDS : 'O',
    OPEN  : 'L',
    RRSP  : 'M',
    TFSA  : 'N',
    HOUSE : 'I'
}
AST_BASE_YEAR = 2007
# number of rows between quarters in the same year
AST_QTR_SPAN = 1
# number of rows between years
AST_BASE_YEAR_SPAN = 4


def year_span(year):
    """
    For Asset rows, have to factor in the header row placed every three years
    :param year: int: year to calculate for
    :return: int: year span to use in figuring out which row to update
    """
    diff = year - AST_BASE_YEAR
    return diff + (diff // 3)


# noinspection PyUnresolvedReferences
def gnc_numeric_to_python_decimal(numeric):
    negative = numeric.negative_p()
    sign = 1 if negative else 0

    val = GncNumeric(numeric.num(), numeric.denom())
    result = val.to_decimal(None)
    if not result:
        raise Exception("GncNumeric value '{}' CANNOT be converted to decimal!".format(val.to_string()))

    digit_tuple = tuple(int(char) for char in str(val.num()) if char != '-')
    denominator = val.denom()
    exponent = int(log10(denominator))
    assert( (10 ** exponent) == denominator )
    return Decimal((sign, digit_tuple, -exponent))


def next_period_start(start_year, start_month):
    # add number of months for a Quarter
    end_month = start_month + PERIOD_QTR

    # use integer division to find out if the new end month is in a different year,
    # what year it is, and what the end month number should be changed to.
    end_year = start_year + ((end_month - 1) // NUM_MONTHS)
    end_month = ((end_month - 1) % NUM_MONTHS) + 1

    return end_year, end_month


def period_end(start_year, start_month):
    end_year, end_month = next_period_start(start_year, start_month)

    # last step, the end date is one day back from the start of the next period
    # so we get a period end like 2010-03-31 instead of 2010-04-01
    return date(end_year, end_month, 1) - ONE_DAY


def generate_period_boundaries(start_year, start_month, periods):
    for i in range(periods):
        yield( date(start_year, start_month, 1), period_end(start_year, start_month) )
        start_year, start_month = next_period_start(start_year, start_month)


def account_from_path(top_account, account_path, original_path=None):
    # print("top_account = %s, account_path = %s, original_path = %s" % (top_account, account_path, original_path))
    if original_path is None:
        original_path = account_path
    account, account_path = account_path[0], account_path[1:]
    # print("account = %s, account_path = %s" % (account, account_path))

    account = top_account.lookup_by_name(account)
    # print("account = " + str(account))
    if account is None:
        raise Exception("Path '" + str(original_path) + "' could NOT be found!")
    if len(account_path) > 0:
        return account_from_path(account, account_path, original_path)
    else:
        return account


def save_to_json(title, time, json_data):
    """
    print json data to a file -- add a time string to get a unique file name each run
    :param title: string
    :param time: string
    :param json_data: json compatible struct
    :return: file name
    """
    out_file = title + '.' + time + ".json"
    print("\njson file is '{}'".format(out_file))
    fp = open(out_file, 'w')
    json.dump(json_data, fp, indent=4)
    fp.close()
    return out_file


def get_splits(acct, period_starts, period_list):
    """
    get the splits for the account and each sub-account
    :param acct: Gnucash account
    :param period_starts: list: start date for each period
    :param period_list: list of structs: store the dates and amounts for each quarter
    :return: nil
    """
    # insert and add all splits in the periods of interest
    for split in acct.GetSplitList():
        trans = split.parent
        # GetDate() returns a datetime but need a date
        trans_date = trans.GetDate().date()

        # use binary search to find the period that starts before or on the transaction date
        period_index = bisect_right(period_starts, trans_date) - 1

        # ignore transactions with a date before the matching period start and after the last period_end
        if period_index >= 0 and trans_date <= period_list[len(period_list) - 1][1]:
            # get the period bucket appropriate for the split in question
            period = period_list[period_index]
            assert( period[1] >= trans_date >= period[0] )

            split_amount = gnc_numeric_to_python_decimal(split.GetAmount())

            # if the amount is negative this is a credit, else a debit
            debit_credit_offset = 1 if split_amount < ZERO else 0

            # add the debit or credit to the sum, using the offset to get in the right bucket
            period[2 + debit_credit_offset] += split_amount

            # add the debit or credit to the overall total
            period[4] += split_amount


def fill_splits(root_acct, target_path, period_starts, period_list):
    """
    fill the period list for each account
    :param root_acct: Gnucash account: from the Gnucash book
    :param target_path: list: account hierarchy from root account to target account
    :param period_starts: list: start date for each period
    :param period_list: list of structs: store the dates and amounts for each quarter
    :return: acct_name: string: name of target_acct
    """
    account_of_interest = account_from_path(root_acct, target_path)
    acct_name = account_of_interest.GetName()
    print("\naccount_of_interest = {}".format(acct_name))

    # get the split amounts for the parent account
    get_splits(account_of_interest, period_starts, period_list)
    descendants = account_of_interest.get_descendants()
    if len(descendants) > 0:
        # for EACH sub-account add to the overall total
        # print("Descendants of {}:".format(account_of_interest.GetName()))
        for subAcct in descendants:
            # print("{} balance = {}".format(subAcct.GetName(), gnc_numeric_to_python_decimal(subAcct.GetBalance())))
            get_splits(subAcct, period_starts, period_list)

    csv_write_period_list(period_list)
    return acct_name


def csv_write_period_list(period_list):
    """
    Write out the details of the submitted period list in csv format
    :param period_list: list of structs: store the dates and amounts for each quarter
    :return: nil
    """
    # write out the column headers
    csv_writer = csv.writer(stdout)
    # csv_writer.writerow('')
    csv_writer.writerow(('period start', 'period end', 'debits', 'credits', 'TOTAL'))

    # write out the overall totals for the account of interest
    for start_date, end_date, debit_sum, credit_sum, total in period_list:
        csv_writer.writerow((start_date, end_date, debit_sum, credit_sum, total))
