#
# updateBalance.py -- use the Gnucash and Google APIs to update the 'Balance' sheet
#                     in my BudgetQtrly document for today or for a specified year or years
#
# some code from account_analysis.py by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
# some code from Google quickstart spreadsheets examples
#
# @author Mark Sattolo <epistemik@gmail.com>
# @version Python 3.6
# @created 2019-04-13
# @updated 2019-04-20

from sys import argv
from gnucash import Session
from googleapiclient.discovery import build
from updateCommon import *

# constant strings
ASTS  = 'FAMILY'
LIAB  = 'LIABS'
TRUST = 'TRUST'
CHAL  = 'XCHALET'
HOUSE = 'House'
TODAY = 'Today'

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
    TODAY : 'C',
    ASTS  : 'K'
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


def get_acct_bal(acct, idate, cur):
    """
    get the BALANCE in this account on this date in this currency
    :param  acct: Gnucash Account
    :param idate: Date
    :param   cur: Gnucash commodity
    :return: Decimal with balance
    """
    # CALLS ARE RETRIEVING ACCOUNT BALANCES FROM DAY BEFORE!!??
    idate += ONE_DAY

    acct_bal = acct.GetBalanceAsOfDate(idate)
    acct_comm = acct.GetCommodity()
    # check if account is already in the desired currency and convert if necessary
    acct_cur = acct_bal if acct_comm == cur else acct.ConvertBalanceToCurrencyAsOfDate(acct_bal, acct_comm, cur, idate)

    return gnc_numeric_to_python_decimal(acct_cur)


def get_total_balance(root_acct, path, tdate, cur):
    """
    get the total BALANCE in the account and all sub-accounts on this path on this date in this currency
    :param root_acct:   Gnucash Account: from the Gnucash book
    :param      path:              list: path to the account
    :param     tdate:              date: to get the balance
    :param       cur: Gnucash Commodity: currency to use for the totals
    :return: string, int: account name and account sum
    """
    acct = account_from_path(root_acct, path)
    acct_name = acct.GetName()
    # get the split amounts for the parent account
    acct_sum = get_acct_bal(acct, tdate, cur)
    descendants = acct.get_descendants()
    if len(descendants) > 0:
        # for EACH sub-account add to the overall total
        for sub_acct in descendants:
            # ?? GETTING SLIGHT ROUNDING ERRORS WHEN ADDING MUTUAL FUND VALUES...
            acct_sum += get_acct_bal(sub_acct, tdate, cur)

    print_info("{} on {} = ${}".format(acct_name, tdate, acct_sum), MAGENTA)
    return acct_name, acct_sum


# noinspection PyUnboundLocalVariable
def fill_today(root_account, dest, cur):
    """
    Get Balance data for TODAY: LIABS, House, FAMILY, XCHALET, TRUST
    :param root_account:   Gnucash Account: from the Gnucash book
    :param         dest:            string: Google sheet to update
    :param          cur: Gnucash Commodity: currency to use for the totals
    :return: list: cell(s) with location and value to update on Google sheet
    """
    data = []
    for item in BALANCE_ACCTS:
        path = BALANCE_ACCTS[item]
        acct_name, acct_sum = get_total_balance(root_account, path, today, cur)

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


def fill_all_years(root_account, dest, cur):
    """
    LIABS for all years
    :param root_account:   Gnucash Account: from the Gnucash book
    :param         dest:            string: Google sheet to update
    :param          cur: Gnucash Commodity: currency to use for the totals
    :return: list: cell(s) with location and value to update on Google sheet
    """
    data = []
    for i in range(today.year-BASE_YEAR-1):
        year_end = date(BASE_YEAR+i, 12, 31)
        print("year_end = {}".format(year_end))
        # fill LIABS
        acct_name, liab_sum = get_total_balance(root_account, BALANCE_ACCTS[LIAB], year_end, cur)
        yr_span = year_span(year_end.year - BASE_YEAR, BASE_YEAR_SPAN, HDR_SPAN)
        fill_cell(dest, BAL_MTHLY_COLS[LIAB][YR], BASE_ROW + yr_span, liab_sum, data)

    return data


def fill_current_year(root_account, dest, cur):
    """
    CURRENT YEAR: fill_today() AND: LIABS for ALL completed month_ends; FAMILY for ALL non-3 completed month_ends in year
    :param root_account: Gnucash Account: from the Gnucash book
    :param         dest: Google sheet to update
    :param          cur: Gnucash Commodity: currency to use for the totals
    :return: list: cell(s) with location and value to update on Google sheet
    """
    data = fill_today(root_account, dest, cur)

    for i in range(today.month-1):
        month_end = date(today.year, i+2, 1)-ONE_DAY
        print("month_end = {}".format(month_end))

        # fill LIABS
        acct_name, liab_sum = get_total_balance(root_account, BALANCE_ACCTS[LIAB], month_end, cur)
        fill_cell(dest, BAL_MTHLY_COLS[LIAB][MTH], BASE_MTHLY_ROW + month_end.month, liab_sum, data)

        # fill ASSETS for months NOT covered by the Assets sheet
        if month_end.month % 3 != 0:
            acct_name, acct_sum = get_total_balance(root_account, BALANCE_ACCTS[ASTS], month_end, cur)
            adjusted_assets = acct_sum - liab_sum
            print_info("Adjusted assets on {} = ${}".format(month_end, adjusted_assets.to_eng_string()), MAGENTA)
            fill_cell(dest, BAL_MTHLY_COLS[ASTS], BASE_MTHLY_ROW + month_end.month, adjusted_assets, data)

    return data


def fill_previous_year(root_account, dest, cur):
    """
    PREVIOUS YEAR: LIABS for ALL NON-completed months; FAMILY for ALL non-3 NON-completed months in year
    :param root_account:   Gnucash Account: from the Gnucash book
    :param         dest:            string: Google sheet to update
    :param          cur: Gnucash Commodity: currency to use for the totals
    :return: list: cell(s) with location and value to update on Google sheet
    """
    data = []
    year = today.year - 1
    for i in range(12-today.month):
        dte = date(year, i+5, 1)-ONE_DAY
        print("date = {}".format(dte))

        # fill LIABS
        acct_name, liab_sum = get_total_balance(root_account, BALANCE_ACCTS[LIAB], dte, cur)
        fill_cell(dest, BAL_MTHLY_COLS[LIAB][MTH], BASE_MTHLY_ROW + dte.month, liab_sum, data)

        # fill ASSETS for months NOT covered by the Assets sheet
        if dte.month % 3 != 0:
            acct_name, acct_sum = get_total_balance(root_account, BALANCE_ACCTS[ASTS], dte, cur)
            adjusted_assets = acct_sum - liab_sum
            print_info("Adjusted assets on {} = ${}".format(dte, adjusted_assets.to_eng_string()), MAGENTA)
            fill_cell(dest, BAL_MTHLY_COLS[ASTS], BASE_MTHLY_ROW + dte.month, adjusted_assets, data)

    # LIABS entry for year end
    year_end = date(year, 12, 31)
    acct_name, liab_sum = get_total_balance(root_account, BALANCE_ACCTS[LIAB], year_end, cur)
    fill_cell(dest, BAL_MTHLY_COLS[LIAB][MTH], BASE_MTHLY_ROW + 12, liab_sum, data)

    return data


def fill_year(year, root_account, dest, cur):
    """
    LIABS for year
    :param         year:               int: get data for this year
    :param root_account:   Gnucash Account: from the Gnucash book
    :param         dest:            string: Google sheet to update
    :param          cur: Gnucash Commodity: currency to use for the totals
    :return: list: cell(s) with location and value to update on Google sheet
    """
    data = []
    year_end = date(year, 12, 31)
    print("year_end = {}".format(year_end))

    # fill LIABS
    acct_name, liab_sum = get_total_balance(root_account, BALANCE_ACCTS[LIAB], year_end, cur)
    yr_span = year_span(year - BASE_YEAR, BASE_YEAR_SPAN, HDR_SPAN)
    fill_cell(dest, BAL_MTHLY_COLS[LIAB][YR], BASE_ROW + yr_span, liab_sum, data)

    return data


# noinspection PyUnboundLocalVariable,PyUnresolvedReferences
def get_gnucash_data(gnucash_file, domain, dest):
    """
    Get Balance data for TODAY:
      LIABS, House, FAMILY, XCHALET, TRUST
    OR for the specified year:
      IF CURRENT YEAR: TODAY & LIABS for ALL completed months; FAMILY for ALL non-3 completed months in year
      IF PREVIOUS YEAR: LIABS for ALL NON-completed months; FAMILY for ALL non-3 NON-completed months in year
      ELSE: LIABS for year
    :param gnucash_file: string: name of file used to read the values
    :param       domain: string: what to update
    :param         dest: string: Google sheet to update
    :return: list: cell(s) with location and value to update on Google sheet
    """
    print_info("find Balances in {} for {}".format(gnucash_file, domain), GREEN)

    try:
        gnucash_session = Session(gnucash_file, is_new=False)
        root_account = gnucash_session.book.get_root_account()
        commod_tab = gnucash_session.book.get_table()
        # noinspection PyPep8Naming
        CAD = commod_tab.lookup("ISO4217", "CAD")

        if domain == 'today':
            data = fill_today(root_account, dest, CAD)
        elif domain == 'allyears':
            data = fill_all_years(root_account, dest, CAD)
        else:
            year = get_year(domain, BASE_YEAR)
            if year == today.year:
                data = fill_current_year(root_account, dest, CAD)
            elif today.year - year == 1:
                data = fill_previous_year(root_account, dest, CAD)
            else:
                data = fill_year(year, root_account, dest, CAD)

        # no save needed, we're just reading...
        gnucash_session.end()

        if data:
            fname = "out/updateBalance_{}".format(domain)
            save_to_json(fname, now, data)

        return data

    except Exception as ge:
        print_error("Exception: {}!".format(ge))
        if "gnucash_session" in locals() and gnucash_session is not None:
            gnucash_session.end()
        exit(356)


def send_google_data(mode, data):
    """
    Send the data list to the document
    :param mode: string: '.?[send][1]'
    :param data:   list: Gnucash data for each needed quarter
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
        exit(385)

    return response


def update_balance_main():
    """
    Main: check command line and call functions to get the data from Gnucash book and send to Google document
    :return: nil
    """
    exe = argv[0].split('/')[-1]
    if len(argv) < 4:
        print_error("NOT ENOUGH parameters!")
        print_info("usage: {} <book url> mode=<.?[send][1]> <year|'today'|'allyears'>".format(exe), GREEN)
        print_error("PROGRAM EXIT!")
        return

    gnucash_file = argv[1]

    mode = argv[2].lower()
    dest = BAL_2_SHEET
    if '1' in mode:
        dest = BAL_1_SHEET

    print_info("\nrunning '{}' on '{}' in mode '{}' at run-time: {}\n".format(exe, gnucash_file, mode, now), GREEN)

    domain = argv[3].lower()

    # get the requested data from Gnucash and package in the update format required by Google spreadsheets
    data = get_gnucash_data(gnucash_file, domain, dest)

    response = send_google_data(mode, data)
    if response:
        fname = "out/updateBalance_{}-response".format(domain)
        save_to_json(fname, now, response)

    print_info(" >>> PROGRAM ENDED.\n", GREEN)


if __name__ == "__main__":
    update_balance_main()
