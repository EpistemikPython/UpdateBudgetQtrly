##############################################################################################################################
# coding=utf-8
#
# updateBalance.py -- use the Gnucash and Google APIs to update the 'Balance' sheet
#                     in my BudgetQtrly document for today or for a specified year or years
#
# some code from account_analysis.py by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
# some code from Google quickstart spreadsheets examples
#
# Copyright (c) 2019 Mark Sattolo <epistemik@gmail.com>
#
__author__ = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__python_version__ = 3.6
__created__ = '2019-04-13'
__updated__ = '2019-08-03'

from gnucash import Session
from googleapiclient.discovery import build
from updateCommon import *
from updateAssets import QTR_SPAN, ASSET_COLS, BASE_YEAR as AST_BY, BASE_YEAR_SPAN as AST_BYS, HDR_SPAN as AST_HS

# path to the accounts in the Gnucash file
BALANCE_ACCTS = {
    HOUSE : [ASTS, HOUSE] ,
    LIAB  : [LIAB]  ,
    TRUST : [TRUST] ,
    CHAL  : [CHAL]  ,
    ASTS  : [ASTS]
}

BAL_MTHLY_COLS = {
    LIAB  : {YR: 'U', MTH: 'L'},
    DATE  : 'B' ,
    TODAY : 'C' ,
    ASTS  : 'K' ,
    MTH   : 'I'
}
BASE_MTHLY_ROW = 19

# cell locations in the Google file
BAL_TODAY_RANGES = {
    HOUSE : 26 ,
    LIAB  : 28 ,
    TRUST : 21 ,
    CHAL  : 22 ,
    ASTS  : 27
}

BASE_YEAR = 2008
# number of rows between same quarter in adjacent years
BASE_YEAR_SPAN = 1
# number of year groups between header rows
HDR_SPAN = 9


def get_acct_bal(acct, p_date:date, p_currency):
    """
    get the BALANCE in this account on this date in this currency
    :param       acct: Gnucash Account
    :param     p_date: required
    :param p_currency: Gnucash commodity
    :return: Decimal with balance
    """
    # CALLS ARE RETRIEVING ACCOUNT BALANCES FROM DAY BEFORE!!??
    p_date += ONE_DAY

    acct_bal = acct.GetBalanceAsOfDate(p_date)
    acct_comm = acct.GetCommodity()
    # check if account is already in the desired currency and convert if necessary
    acct_cur = acct_bal if acct_comm == p_currency else acct.ConvertBalanceToCurrencyAsOfDate(acct_bal, acct_comm, p_currency, p_date)

    return gnc_numeric_to_python_decimal(acct_cur)


def get_total_balance(root_acct, path:list, p_date:date, p_currency):
    """
    get the total BALANCE in the account and all sub-accounts on this path on this date in this currency
    :param  root_acct: Gnucash Account from the Gnucash book
    :param       path: path to the account
    :param     p_date: to get the balance
    :param p_currency: Gnucash Commodity: currency to use for the totals
    :return: string, int: account name and account sum
    """
    acct = account_from_path(root_acct, path)
    acct_name = acct.GetName()
    # get the split amounts for the parent account
    acct_sum = get_acct_bal(acct, p_date, p_currency)
    descendants = acct.get_descendants()
    if len(descendants) > 0:
        # for EACH sub-account add to the overall total
        for sub_acct in descendants:
            # ?? GETTING SLIGHT ROUNDING ERRORS WHEN ADDING MUTUAL FUND VALUES...
            acct_sum += get_acct_bal(sub_acct, p_date, p_currency)

    print_info("{} on {} = ${}".format(acct_name, p_date, acct_sum), MAGENTA)
    return acct_name, acct_sum


# noinspection PyUnboundLocalVariable
def fill_today(root_account, dest:str, p_currency):
    """
    Get Balance data for TODAY: LIABS, House, FAMILY, XCHALET, TRUST
    :param root_account: Gnucash Account from the Gnucash book
    :param         dest: Google sheet to update
    :param   p_currency: Gnucash Commodity: currency to use for the totals
    :return: list of cell(s) with location and value to update on Google sheet
    """
    data = []
    # calls using 'today' ARE NOT off by one day??
    tdate = today - ONE_DAY
    for item in BALANCE_ACCTS:
        path = BALANCE_ACCTS[item]
        acct_name, acct_sum = get_total_balance(root_account, path, tdate, p_currency)

        # need assets apart from house and liabilities which are reported separately
        if item == HOUSE:
            house_sum = acct_sum
        elif item == LIAB:
            liab_sum = acct_sum
        elif item == ASTS:
            acct_sum = acct_sum - house_sum - liab_sum
            print_info("Adjusted assets on {} = ${}".format(today, acct_sum.to_eng_string()), MAGENTA)

        fill_cell(dest, BAL_MTHLY_COLS[TODAY], BAL_TODAY_RANGES[item], acct_sum, data)

    return data


def fill_all_years(root_account, dest:str, p_currency):
    """
    LIABS for all years
    :param root_account: Gnucash Account from the Gnucash book
    :param         dest: Google sheet to update
    :param   p_currency: Gnucash Commodity: currency to use for the totals
    :return: list of cell(s) with location and value to update on Google sheet
    """
    data = []
    for i in range(today.year - BASE_YEAR - 1):
        year = BASE_YEAR + i
        # fill LIABS
        fill_year(year, root_account, dest, p_currency, data)

    return data


def fill_current_year(root_account, dest:str, p_currency):
    """
    CURRENT YEAR: fill_today() AND: LIABS for ALL completed month_ends; FAMILY for ALL non-3 completed month_ends in year
    :param root_account: Gnucash Account from the Gnucash book
    :param         dest: Google sheet to update
    :param   p_currency: Gnucash Commodity: currency to use for the totals
    :return: list of cell(s) with location and value to update on Google sheet
    """
    data = fill_today(root_account, dest, p_currency)

    for i in range(today.month - 1):
        month_end = date(today.year, i+2, 1)-ONE_DAY
        print_info("month_end = {}".format(month_end), BLUE)

        row = BASE_MTHLY_ROW + month_end.month
        # fill LIABS
        acct_name, liab_sum = get_total_balance(root_account, BALANCE_ACCTS[LIAB], month_end, p_currency)
        fill_cell(dest, BAL_MTHLY_COLS[LIAB][MTH], row, liab_sum, data)

        # fill ASSETS for months NOT covered by the Assets sheet
        if month_end.month % 3 != 0:
            acct_name, acct_sum = get_total_balance(root_account, BALANCE_ACCTS[ASTS], month_end, p_currency)
            adjusted_assets = acct_sum - liab_sum
            print_info("Adjusted assets on {} = ${}".format(month_end, adjusted_assets.to_eng_string()), MAGENTA)
            fill_cell(dest, BAL_MTHLY_COLS[ASTS], row, adjusted_assets, data)
        else:
            print_info('Update reference to Assets sheet for Mar, June, Sep or Dec', GREEN)
            # have to update the CELL REFERENCE to current year/qtr ASSETS
            year_row = BASE_ROW + year_span(today.year, AST_BY, AST_BYS, AST_HS)
            # print_info('year_row = {}'.format(year_row))
            int_qtr = (month_end.month // 3) - 1
            # print_info("int_qtr = {}".format(int_qtr))
            dest_row = year_row + (int_qtr * QTR_SPAN)
            # print_info("dest_row = {}".format(dest_row))
            val_num = '1' if '1' in dest else '2'
            # print_info("val_num = {}".format(val_num))
            value = "='Assets " + val_num + "'!" + ASSET_COLS[TOTAL] + str(dest_row)
            # print_info("value = {}".format(value))
            fill_cell(dest, BAL_MTHLY_COLS[ASTS], row, value, data)

        # fill DATE for month column
        fill_cell(dest, BAL_MTHLY_COLS[MTH], row, str(month_end), data)

    return data


def fill_previous_year(root_account, dest:str, p_currency):
    """
    PREVIOUS YEAR: LIABS for ALL NON-completed months; FAMILY for ALL non-3 NON-completed months in year
    :param root_account: Gnucash Account from the Gnucash book
    :param         dest: Google sheet to update
    :param   p_currency: Gnucash Commodity: currency to use for the totals
    :return: list of cell(s) with location and value to update on Google sheet
    """
    data = []
    year = today.year - 1
    for i in range(12-today.month):
        dte = date(year, i+today.month+1, 1)-ONE_DAY
        print_info("date = {}".format(dte), BLUE)

        row = BASE_MTHLY_ROW + dte.month
        # fill LIABS
        acct_name, liab_sum = get_total_balance(root_account, BALANCE_ACCTS[LIAB], dte, p_currency)
        fill_cell(dest, BAL_MTHLY_COLS[LIAB][MTH], row, liab_sum, data)

        # fill ASSETS for months NOT covered by the Assets sheet
        if dte.month % 3 != 0:
            acct_name, acct_sum = get_total_balance(root_account, BALANCE_ACCTS[ASTS], dte, p_currency)
            adjusted_assets = acct_sum - liab_sum
            print_info("Adjusted assets on {} = ${}".format(dte, adjusted_assets.to_eng_string()), MAGENTA)
            fill_cell(dest, BAL_MTHLY_COLS[ASTS], row, adjusted_assets, data)

    # LIABS entry for year end
    year_end = date(year, 12, 31)
    acct_name, liab_sum = get_total_balance(root_account, BALANCE_ACCTS[LIAB], year_end, p_currency)
    # month column
    fill_cell(dest, BAL_MTHLY_COLS[LIAB][MTH], BASE_MTHLY_ROW + 12, liab_sum, data)
    # year column
    fill_year(year, root_account, dest, p_currency, data)

    return data


# noinspection PyUnboundLocalVariable
def fill_year(year:int, root_account, dest:str, p_currency, data_list:list=None):
    """
    :param         year: to get data for
    :param root_account: Gnucash Account from the Gnucash book
    :param         dest: Google sheet to update
    :param   p_currency: Gnucash Commodity: currency to use for the totals
    :param    data_list: optional list to fill
    LIABS for year column
    :return: list of cell(s) with location and value to update on Google sheet
    """
    year_end = date(year, 12, 31)
    print_info("year_end = {}".format(year_end), BLUE)

    # fill LIABS
    acct_name, liab_sum = get_total_balance(root_account, BALANCE_ACCTS[LIAB], year_end, p_currency)
    yr_span = year_span(year, BASE_YEAR, BASE_YEAR_SPAN, HDR_SPAN)
    if data_list is None:
        data = fill_cell(dest, BAL_MTHLY_COLS[LIAB][YR], BASE_ROW + yr_span, liab_sum)
    else:
        fill_cell(dest, BAL_MTHLY_COLS[LIAB][YR], BASE_ROW + yr_span, liab_sum, data_list)

    return data if data_list is None else data_list


# noinspection PyUnboundLocalVariable,PyUnresolvedReferences
def fill_google_data(gnucash_file:str, domain:str, dest:str):
    """
    :param gnucash_file: name of file used to read the values
    :param       domain: to update
    :param         dest: Google sheet to update
    Get Balance data for TODAY:
      LIABS, House, FAMILY, XCHALET, TRUST
    OR for the specified year:
      IF CURRENT YEAR: TODAY & LIABS for ALL completed months; FAMILY for ALL non-3 completed months in year
      IF PREVIOUS YEAR: LIABS for ALL NON-completed months; FAMILY for ALL non-3 NON-completed months in year
      ELSE: LIABS for year
    :return: list of cell(s) with location and value to update on Google sheet
    """
    print_info("find Balances in {} for {}".format(gnucash_file, domain), GREEN)

    try:
        gnucash_session = Session(gnucash_file, is_new=False)
        root_account = gnucash_session.book.get_root_account()
        commod_tab = gnucash_session.book.get_table()
        # noinspection PyPep8Naming
        CAD = commod_tab.lookup("ISO4217", "CAD")

        if domain == 'today':
            google_data = fill_today(root_account, dest, CAD)
        elif domain == 'allyears':
            google_data = fill_all_years(root_account, dest, CAD)
        else:
            year = get_int_year(domain, BASE_YEAR)
            if year == today.year:
                google_data = fill_current_year(root_account, dest, CAD)
            elif today.year - year == 1:
                google_data = fill_previous_year(root_account, dest, CAD)
            else:
                google_data = fill_year(year, root_account, dest, CAD)

        # fill today's date
        fill_cell(dest, BAL_MTHLY_COLS[DATE], BASE_MTHLY_ROW, today.strftime(DAY_STR), google_data)

        # no save needed, we're just reading...
        gnucash_session.end()

        if google_data:
            fname = "out/updateBalance_{}".format(domain)
            save_to_json(fname, now, google_data)

        return google_data

    except Exception as ge:
        print_error("Exception: {}!".format(ge))
        if "gnucash_session" in locals() and gnucash_session is not None:
            gnucash_session.end()
        exit(279)


def send_google_data(mode:str, data:list):
    """
    Send the data list to the document
    :param mode: '.?[send][1]'
    :param data: Gnucash data for each needed quarter
    :return: server response
    """
    print_info("send_google_data({})\n".format(mode))

    response = None
    try:
        assets_body = {
            'valueInputOption': 'USER_ENTERED',
            'data': data
        }

        if 'send' in mode:
            creds = get_credentials()
            service = build('sheets', 'v4', credentials=creds)
            vals = service.spreadsheets().values()
            response = vals.batchUpdate(spreadsheetId=get_budget_id(), body=assets_body).execute()

            print_info('{} cells updated!\n'.format(response.get('totalUpdatedCells')))

    except Exception as se:
        print_error("Exception: {}!".format(se))
        exit(308)

    return response


# TODO: fill in date column for previous month when updating 'today', check to update 'today' or 'tomorrow'
def update_balance_main(args:list):
    """
    Main: check command line and call functions to get the data from Gnucash book and send to Google document
    :return: string
    """
    if len(args) < 3:
        print_error("NOT ENOUGH parameters!")
        usage = "usage: py36 updateBalance.py <book url> mode=<.?[send][1]> <year | 'today' | 'allyears'>"
        print_info(usage, GREEN)
        print_error("PROGRAM EXIT!")
        return usage

    gnucash_file = args[0]

    mode = args[1].lower()
    dest = BAL_2_SHEET
    if '1' in mode:
        dest = BAL_1_SHEET

    print_info("\nrunning in mode '{}' at run-time: {}\n".format(mode, now), CYAN)

    domain = args[2].lower()

    # get the requested data from Gnucash and package in the update format required by Google spreadsheets
    data = fill_google_data(gnucash_file, domain, dest)

    response = send_google_data(mode, data)

    print_info(" >>> PROGRAM ENDED.\n", GREEN)

    if response:
        fname = "out/updateBalance_{}-response".format(domain)
        save_to_json(fname, now, response)
        return response
    else:
        return data


if __name__ == "__main__":
    import sys
    update_balance_main(sys.argv[1:])
