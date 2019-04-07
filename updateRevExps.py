#
# updateRevExps.py -- use the Gnucash and Google APIs to update the Revenue and Expenses
#                     in my BudgetQtrly document for a specified year or quarter
#
# some code from account_analysis.py by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
# some code from Google quickstart spreadsheets example
#
# @author Mark Sattolo <epistemik@gmail.com>
# @version Python 3.6
# @created 2019-03-30
# @updated 2019-04-07

from sys import argv
from datetime import datetime as dt
from gnucash import Session
import pickle
import os.path as osp
import copy
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from updateCommon import *

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

# store the values needed to update the document
REV_EXP_RESULTS = {
    QTR   : '0',
    REV   : '0',
    BAL   : '0',
    CONT  : '0',
    NEC   : '0',
    DEDNS : '0'
}

REV_EXP_COLS = {
    REV   : 'D',
    BAL   : 'P',
    CONT  : 'O',
    NEC   : 'G',
    DEDNS : 'D'
}
BASE_YEAR = 2012
# number of rows between quarters in the same year
QTR_SPAN = 2
# number of rows between years
YEAR_SPAN = 11

now = dt.now().strftime("%Y-%m-%dT%H-%M-%S")


def get_revenue(root_account, period_starts, period_list, re_year, qtr):
    """
    Get REVENUE data for the specified Quarter
    :param root_account: Gnucash account: from the Gnucash book
    :param period_starts: list: start date for each period
    :param period_list: list of structs: store the dates and amounts for each quarter
    :param re_year: int: year to read
    :param qtr: int: quarter to read: 1..4
    :return: string with revenue
    """
    str_rev = '= '
    for item in REV_ACCTS:
        # reset the debit and credit totals for each individual account
        period_list[0][2] = 0
        period_list[0][3] = 0

        acct_base = REV_ACCTS[item]
        acct_name = fill_splits(root_account, acct_base, period_starts, period_list)

        sum_revenue = (period_list[0][2] + period_list[0][3]) * (-1)
        str_rev += sum_revenue.to_eng_string() + (' + ' if item != SAL else '')
        print("{} Revenue for {}-Q{} = ${}".format(acct_name, re_year, qtr, sum_revenue))

    REV_EXP_RESULTS[REV] = str_rev
    return str_rev


def get_deductions(root_account, period_starts, period_list, re_year, qtr):
    """
    Get SALARY DEDUCTIONS data for the specified Quarter
    :param root_account: Gnucash account: from the Gnucash book
    :param period_starts: list: start date for each period
    :param period_list: list of structs: store the dates and amounts for each quarter
    :param re_year: int: year to read
    :param qtr: int: quarter to read: 1..4
    :return: string with deductions
    """
    str_dedns = '= '
    for item in DEDN_ACCTS:
        # reset the debit and credit totals for each individual account
        period_list[0][2] = 0
        period_list[0][3] = 0

        acct_base = DEDN_ACCTS[item]
        acct_name = fill_splits(root_account, acct_base, period_starts, period_list)

        sum_deductions = period_list[0][2] + period_list[0][3]
        str_dedns += sum_deductions.to_eng_string() + (' + ' if item != "ML" else '')
        print("{} Salary Deductions for {}-Q{} = ${}".format(acct_name, re_year, qtr, sum_deductions))

    REV_EXP_RESULTS[DEDNS] = str_dedns
    return str_dedns


def get_expenses(root_account, period_starts, period_list, re_year, qtr):
    """
    Get EXPENSE data for the specified Quarter
    :param root_account: Gnucash account: from the Gnucash book
    :param period_starts: list: start date for each period
    :param period_list: list of structs: store the dates and amounts for each quarter
    :param re_year: int: year to read
    :param qtr: int: quarter to read = 1..4
    :return: string with total expenses
    """
    str_total = ''
    for item in EXP_ACCTS:
        # reset the debit and credit totals for each individual account
        period_list[0][2] = 0
        period_list[0][3] = 0

        acct_base = EXP_ACCTS[item]
        acct_name = fill_splits(root_account, acct_base, period_starts, period_list)

        sum_expenses = period_list[0][2] + period_list[0][3]
        str_expenses = sum_expenses.to_eng_string()
        REV_EXP_RESULTS[item] = str_expenses
        print("{} Expenses for {}-Q{} = ${}".format(acct_name.split('_')[-1], re_year, qtr, str_expenses))
        str_total += str_expenses + ' + '

    get_deductions(root_account, period_starts, period_list, re_year, qtr)

    return str_total


# noinspection PyUnboundLocalVariable,PyUnresolvedReferences
def get_rev_exps(gnucash_file, re_year, re_quarter):
    """
    Get REVENUE and EXPENSE data for ONE specified Quarter or ALL four Quarters for the specified Year
    :param gnucash_file: string: name of file used to read the values
    :param re_year: int: year to update
    :param re_quarter: int: 1..4 for quarter to update or 0 if updating entire year
    :return: nil
    """
    num_quarters = 1 if re_quarter else 4
    print("find Revenue & Expenses in {} for {}{}".format(gnucash_file, re_year, ('-Q' + str(re_quarter)) if re_quarter else ''))

    try:
        gnucash_session = Session(gnucash_file, is_new=False)
        root_account = gnucash_session.book.get_root_account()
        # print("type root_account = {}".format(type(root_account)))

        for i in range(num_quarters):
            qtr = re_quarter if re_quarter else i + 1
            REV_EXP_RESULTS[QTR] = str(qtr)

            start_month = (qtr * 3) - 2

            # for each period keep the start date, end date, debits and credits sums and overall total
            period_list = [
                [
                    start_date, end_date,
                    ZERO, # debits sum
                    ZERO, # credits sum
                    ZERO  # TOTAL
                ]
                for start_date, end_date in generate_period_boundaries(re_year, start_month, 1)
            ]
            # print(period_list)
            # a copy of the above list with just the period start dates
            period_starts = [e[0] for e in period_list]
            # print(period_starts)

            get_revenue(root_account, period_starts, period_list, re_year, qtr)
            tot_revenue = period_list[0][4] * (-1)
            print("\n{} Revenue for {}-Q{} = ${}".format("TOTAL", re_year, qtr, tot_revenue))

            period_list[0][4] = 0
            get_expenses(root_account, period_starts, period_list, re_year, qtr)
            tot_expenses = period_list[0][4]
            print("\n{} Expenses for {}-Q{} = ${}\n".format("TOTAL", re_year, qtr, tot_expenses))

            results.append(copy.deepcopy(REV_EXP_RESULTS))
            period_list[0][4] = 0
            print(json.dumps(REV_EXP_RESULTS, indent=4))

        # no save needed, we're just reading...
        gnucash_session.end()

        save_to_json('out/updateRevExps_results', now, results)

    except Exception as ge:
        print("Exception: {}!".format(ge))
        if "gnucash_session" in locals() and gnucash_session is not None:
            gnucash_session.end()
        exit(223)


def fill_rev_exps_data(mode, re_year):
    """
    Fill the data list:
    for each item in results, either 1 for one quarter or 4 for four quarters:
    create 5 cell_data's, one each for REV, BAL, CONT, NEC, DEDNS:
    fill in the range based on the year and quarter
    range = SHEET_NAME + '!' + calculated cell
    fill in the values based on the sheet being updated and the type of cell_data
    REV string is '= ${INV} + ${OTH} + ${SAL}'
    DEDNS string is '= ${Mk-Dedns} + ${Lu-Dedns} + ${ML-Dedns}'
    others are just the string from the item
    :param mode: string: 'xxx[prod][send]'
    :param re_year: int: year to update
    :return: data list
    """
    print("\nfill_rev_exps_data({}, {})\n".format(mode, re_year))

    all_inc_dest = ALL_INC_PRAC_SHEET
    nec_inc_dest = NEC_INC_PRAC_SHEET
    if 'prod' in mode:
        all_inc_dest = ALL_INC_SHEET
        nec_inc_dest = NEC_INC_SHEET
    print("all_inc_dest = {}".format(all_inc_dest))
    print("nec_inc_dest = {}\n".format(nec_inc_dest))

    year_row = BASE_ROW + ((re_year - BASE_YEAR) * YEAR_SPAN)
    # get exact row from Quarter value in each item
    for item in results:
        print("{} = {}".format(QTR, item[QTR]))
        int_qtr = int(item[QTR])
        dest_row = year_row + ((int_qtr - 1) * QTR_SPAN)
        print("dest_row = {}\n".format(dest_row))
        for key in item:
            if key != QTR:
                cell = copy.copy(cell_data)
                dest = nec_inc_dest
                if key in (REV, BAL, CONT):
                    dest = all_inc_dest
                col = REV_EXP_COLS[key]
                val = item[key]
                cell_locn = dest + '!' + col + str(dest_row)
                cell['range']  = cell_locn
                cell['values'] = [[val]]
                print("cell = {}".format(cell))
                data.append(cell)
    return data


def send_rev_exps(mode, re_year):
    """
    Fill the data list and send to the document
    :param mode: string: 'xxx[prod][send]'
    :param re_year: int: year to update
    :return: server response
    """
    print("\nsend_rev_exps({}, {})".format(mode, re_year))
    print("cell_data['range'] = {}".format(cell_data['range']))
    print("cell_data['values'][0][0] = {}".format(cell_data['values'][0][0]))

    response = 'EMPTY'
    try:
        fill_rev_exps_data(mode, re_year)

        save_to_json('out/updateRevExps_data', now, data)

        creds = None
        if osp.exists(TOKEN):
            with open(TOKEN, 'rb') as token:
                creds = pickle.load(token)

        # if there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS, SHEETS_RW_SCOPE)
                creds = flow.run_local_server()
            # save the credentials for the next run
            with open(TOKEN, 'wb') as token:
                pickle.dump(creds, token, pickle.HIGHEST_PROTOCOL)

        service = build('sheets', 'v4', credentials=creds)
        srv_sheets = service.spreadsheets()

        my_body = {
            'valueInputOption': 'USER_ENTERED',
            'data': data
        }

        response = 'NO SEND.'
        if 'send' in mode:
            vals = srv_sheets.values()
            response = vals.batchUpdate(spreadsheetId=BUDGET_QTRLY_SPRD_SHEET, body=my_body).execute()

            print('\n{} cells updated!'.format(response.get('totalUpdatedCells')))
            save_to_json('out/updateRevExps_response', now, response)

    except Exception as se:
        print("Exception: {}!".format(se))
        exit(325)

    return response


def update_rev_exps_main():
    """
    Main: check command line and call functions to get the data from Gnucash book and send to Google document
    :return: nil
    """
    exe = argv[0].split('/')[-1]
    if len(argv) < 4:
        print("NOT ENOUGH parameters!")
        print("usage: {} <book url> <mode=xxx[prod][send]> <year> [quarter]".format(exe))
        print("PROGRAM EXIT!")
        return

    gnucash_file = argv[1]
    mode = argv[2].lower()
    print("\nrunning '{}' on '{}' in mode '{}' at run-time: {}\n".format(exe, gnucash_file, mode, now))

    re_year = int(argv[3])
    re_quarter = int(argv[4]) if len(argv) > 4 else 0

    get_rev_exps(gnucash_file, re_year, re_quarter)

    send_rev_exps(mode, re_year)

    print("\n >>> PROGRAM ENDED.")


if __name__ == "__main__":
    update_rev_exps_main()
