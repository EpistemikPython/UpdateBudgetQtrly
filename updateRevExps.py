##############################################################################################################################
# coding=utf-8
#
# updateRevExps.py -- use the Gnucash and Google APIs to update the Revenue and Expenses
#                     in my BudgetQtrly document for a specified year or quarter
#
# some code from account_analysis.py by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
# some code from Google quickstart spreadsheets examples
#
# Copyright (c) 2019 Mark Sattolo <epistemik@gmail.com>
#
__author__ = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__python_version__ = 3.6
__created__ = '2019-03-30'
__updated__ = '2019-07-07'

from gnucash import Session
from googleapiclient.discovery import build
from updateCommon import *

# path to the account in the Gnucash file
REV_ACCTS = {
    INV : ["REV_Invest"],
    OTH : ["REV_Other"],
    EMP : ["REV_Employment"]
}
EXP_ACCTS = {
    BAL   : ["EXP_Balance"],
    CONT  : ["EXP_CONTINGENT"],
    NEC   : ["EXP_NECESSARY"]
}
DEDNS_BASE = 'DEDNS_Employment'
DEDN_ACCTS = {
    "Mark" : [DEDNS_BASE, 'Mark'],
    "Lulu" : [DEDNS_BASE, 'Lulu'],
    "ML"   : [DEDNS_BASE, 'Marie-Laure']
}

# column index in the Google sheets
REV_EXP_COLS = {
    REV   : 'D',
    BAL   : 'P',
    CONT  : 'O',
    NEC   : 'G',
    DEDNS : 'D'
}

BASE_YEAR = 2012
# number of rows between same quarter in adjacent years
BASE_YEAR_SPAN = 11
# number of rows between quarters in the same year
QTR_SPAN = 2


def get_revenue(root_account, period_starts, period_list, re_year, qtr):
    """
    Get REVENUE data for the specified Quarter
    :param  root_account: Gnucash Account: from the Gnucash book
    :param period_starts:            list: start date for each period
    :param   period_list: list of structs: store the dates and amounts for each quarter
    :param       re_year:             int: year to read
    :param           qtr:             int: quarter to read: 1..4
    :return: dict with quarter data
    """
    data_quarter = {}
    str_rev = '= '
    for item in REV_ACCTS:
        # reset the debit and credit totals for each individual account
        period_list[0][2] = ZERO
        period_list[0][3] = ZERO

        acct_base = REV_ACCTS[item]
        acct_name = fill_splits(root_account, acct_base, period_starts, period_list)

        sum_revenue = (period_list[0][2] + period_list[0][3]) * (-1)
        str_rev += sum_revenue.to_eng_string() + (' + ' if item != EMP else '')
        print_info("{} Revenue for {}-Q{} = ${}".format(acct_name, re_year, qtr, sum_revenue))

    data_quarter[REV] = str_rev
    return data_quarter


def get_deductions(root_account, period_starts, period_list, re_year, data_quarter):
    """
    Get SALARY DEDUCTIONS data for the specified Quarter
    :param  root_account: Gnucash Account: from the Gnucash book
    :param period_starts:            list: start date for each period
    :param   period_list: list of structs: store the dates and amounts for each quarter
    :param       re_year:             int: year to read
    :param  data_quarter:            dict: collected data for the quarter
    :return: string with deductions
    """
    str_dedns = '= '
    for item in DEDN_ACCTS:
        # reset the debit and credit totals for each individual account
        period_list[0][2] = ZERO
        period_list[0][3] = ZERO

        acct_path = DEDN_ACCTS[item]
        acct_name = fill_splits(root_account, acct_path, period_starts, period_list)

        sum_deductions = period_list[0][2] + period_list[0][3]
        str_dedns += sum_deductions.to_eng_string() + (' + ' if item != "ML" else '')
        print_info("{} {} Deductions for {}-Q{} = ${}".format(acct_name, EMP, re_year, data_quarter[QTR], sum_deductions))

    data_quarter[DEDNS] = str_dedns
    return str_dedns


def get_expenses(root_account, period_starts, period_list, re_year, data_quarter):
    """
    Get EXPENSE data for the specified Quarter
    :param  root_account: Gnucash Account: from the Gnucash book
    :param period_starts:            list: start date for each period
    :param   period_list: list of structs: store the dates and amounts for each quarter
    :param       re_year:             int: year to read
    :param  data_quarter:            dict: collected data for the quarter
    :return: string with total expenses
    """
    str_total = ''
    for item in EXP_ACCTS:
        # reset the debit and credit totals for each individual account
        period_list[0][2] = ZERO
        period_list[0][3] = ZERO

        acct_base = EXP_ACCTS[item]
        acct_name = fill_splits(root_account, acct_base, period_starts, period_list)

        sum_expenses = period_list[0][2] + period_list[0][3]
        str_expenses = sum_expenses.to_eng_string()
        data_quarter[item] = str_expenses
        print_info("{} Expenses for {}-Q{} = ${}".format(acct_name.split('_')[-1], re_year, data_quarter[QTR], str_expenses))
        str_total += str_expenses + ' + '

    get_deductions(root_account, period_starts, period_list, re_year, data_quarter)

    return str_total


# noinspection PyUnboundLocalVariable,PyUnresolvedReferences
def get_gnucash_data(gnucash_file, re_year, re_quarter):
    """
    Get REVENUE and EXPENSE data for ONE specified Quarter or ALL four Quarters for the specified Year
    :param gnucash_file: string: name of file used to read the values
    :param      re_year:    int: year to update
    :param   re_quarter:    int: 1..4 for quarter to update or 0 if updating ALL FOUR quarters
    :return: list: Gnucash data
    """
    num_quarters = 1 if re_quarter else 4
    print_info("find Revenue & Expenses in {} for {}{}".format(gnucash_file, re_year, ('-Q' + str(re_quarter)) if re_quarter else ''))
    gnc_data = list()
    try:
        gnucash_session = Session(gnucash_file, is_new=False)
        root_account = gnucash_session.book.get_root_account()
        # print_info("type root_account = {}".format(type(root_account)))

        for i in range(num_quarters):
            qtr = re_quarter if re_quarter else i + 1
            start_month = (qtr * 3) - 2

            # for each period keep the start date, end date, debits and credits sums and overall total
            period_list = [
                [
                    start_date, end_date,
                    ZERO, # debits sum
                    ZERO, # credits sum
                    ZERO  # TOTAL
                ]
                for start_date, end_date in generate_quarter_boundaries(re_year, start_month, 1)
            ]
            # a copy of the above list with just the period start dates
            period_starts = [e[0] for e in period_list]

            data_quarter = get_revenue(root_account, period_starts, period_list, re_year, qtr)
            data_quarter[QTR] = str(qtr)
            print_info("\n{} Revenue for {}-Q{} = ${}".format("TOTAL", re_year, qtr, period_list[0][4] * (-1)))

            period_list[0][4] = ZERO
            get_expenses(root_account, period_starts, period_list, re_year, data_quarter)
            print_info("\n{} Expenses for {}-Q{} = ${}\n".format("TOTAL", re_year, qtr, period_list[0][4]))

            gnc_data.append(data_quarter)
            print_info(json.dumps(data_quarter, indent=4))

        # no save needed, we're just reading...
        gnucash_session.end()

        fname = "out/updateRevExps_gnc-data-{}{}".format(re_year, ('-Q' + str(re_quarter)) if re_quarter else '')
        save_to_json(fname, now, gnc_data)
        return gnc_data

    except Exception as ge:
        print_error("Exception: {}!".format(ge))
        if "gnucash_session" in locals() and gnucash_session is not None:
            gnucash_session.end()
        exit(201)


def fill_google_data(mode, re_year, gnc_data):
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
    :param     mode: string: '.?[send][1]'
    :param  re_year:    int: year to update
    :param gnc_data:   list: Gnucash data for each needed quarter
    :return: data list
    """
    all_inc_dest = ALL_INC_2_SHEET
    nec_inc_dest = NEC_INC_2_SHEET
    if '1' in mode:
        all_inc_dest = ALL_INC_SHEET
        nec_inc_dest = NEC_INC_SHEET
    print_info("all_inc_dest = {}".format(all_inc_dest))
    print_info("nec_inc_dest = {}\n".format(nec_inc_dest))

    google_data = list()
    year_row = BASE_ROW + ((re_year - BASE_YEAR) * BASE_YEAR_SPAN)
    # get exact row from Quarter value in each item
    for item in gnc_data:
        print_info("{} = {}".format(QTR, item[QTR]))
        dest_row = year_row + ( (get_quarter(item[QTR]) - 1) * QTR_SPAN )
        print_info("dest_row = {}\n".format(dest_row))
        for key in item:
            if key != QTR:
                dest = nec_inc_dest
                if key in (REV, BAL, CONT):
                    dest = all_inc_dest
                fill_cell(dest, REV_EXP_COLS[key], dest_row, item[key], google_data)

    str_qtr = None
    if len(gnc_data) == 1:
        str_qtr = gnc_data[0][QTR]
    fname = "out/updateRevExps_google-data-{}{}".format(str(re_year), ('-Q' + str_qtr) if str_qtr else '')
    save_to_json(fname, now, google_data)
    return google_data


def send_google_data(mode, re_year, gnc_data):
    """
    Fill the data list and send to the document
    :param     mode: string: '.?[send][1]'
    :param  re_year:    int: year to update
    :param gnc_data:   list: Gnucash data for each needed quarter
    :return: server response
    """
    response = None
    try:
        google_data = fill_google_data(mode, re_year, gnc_data)

        rev_exps_body = {
            'valueInputOption': 'USER_ENTERED',
            'data': google_data
        }

        if 'send' in mode:
            creds = get_credentials()
            service = build('sheets', 'v4', credentials=creds)
            vals = service.spreadsheets().values()
            response = vals.batchUpdate(spreadsheetId=get_budget_id(), body=rev_exps_body).execute()

            print_info('\n{} cells updated!'.format(response.get('totalUpdatedCells')))

    except Exception as se:
        print_error("Exception on Send: {}!".format(se))
        exit(277)

    return response


def update_rev_exps_main(args):
    """
    Main: check command line and call functions to get the data from Gnucash book and send to Google document
    :return: string
    """
    if len(args) < 3:
        print_error("NOT ENOUGH parameters!")
        usage = 'usage: py36 updateRevExps.py <book url> mode=<.?[send][1]> <year> [quarter]'
        print_info(usage, GREEN)
        print_info("PROGRAM EXIT!", MAGENTA)
        return usage

    gnucash_file = args[0]
    mode = args[1].lower()
    print_info("\nrunning in mode '{}' at run-time: {}\n".format(mode, now), CYAN)

    re_year = get_year(args[2], BASE_YEAR)
    re_quarter = get_quarter(args[3]) if len(args) > 3 else 0

    # either for One Quarter or for Four Quarters if updating an entire Year
    gnc_data = get_gnucash_data(gnucash_file, re_year, re_quarter)

    response = send_google_data(mode, re_year, gnc_data)

    print_info("\n >>> PROGRAM ENDED.")

    if response:
        fname = "out/updateRevExps_response-{}{}".format(re_year, ('-Q' + str(re_quarter)) if re_quarter else '')
        save_to_json(fname, now, response)
        return response
    else:
        return gnc_data


if __name__ == "__main__":
    import sys
    update_rev_exps_main(sys.argv[1:])
