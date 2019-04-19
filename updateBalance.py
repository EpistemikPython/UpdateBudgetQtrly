#
# updateBalance.py -- use the Gnucash and Google APIs to update the 'Balance' sheet
#                     in my BudgetQtrly document for today or for a specified year
#
# some code from account_analysis.py by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
# some code from Google quickstart spreadsheets example
#
# @author Mark Sattolo <epistemik@gmail.com>
# @version Python 3.6
# @created 2019-04-13
# @updated 2019-04-15

from sys import argv
from datetime import datetime as dt
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
    HOUSE : '26' ,
    LIAB  : '28' ,
    TRUST : '21' ,
    CHAL  : '22' ,
    ASTS  : '27'
}

BASE_YEAR = 2008
# number of rows between years
BASE_YEAR_SPAN = 1

today = dt.now()
now = today.strftime("%Y-%m-%dT%H-%M-%S")


def year_span(year):
    """
    For Balance rows, have to factor in the header row placed every NINE years
    :param year: int: year to calculate for
    :return: int: value to use to calculate which row to update
    """
    year_diff = year - BASE_YEAR
    return (year_diff * BASE_YEAR_SPAN) + (year_diff // 9)


def get_year(str_year):
    return int(str_year)


def get_acct_bal(acct, idate, cur):
    """
    get the BALANCE in the account on this date in this currency
    :param acct: Gnucash account
    :param idate: Date
    :param cur: Gnucash commodity
    :return: python Decimal with balance
    """
    # CALLS ARE RETRIEVING ASSET VALUES FROM DAY BEFORE!!??
    idate += ONE_DAY

    acct_bal = acct.GetBalanceAsOfDate(idate)
    acct_comm = acct.GetCommodity()
    # print_info("acct_comm = {}".format(acct_comm))
    if acct_comm == cur:
        acct_cad = acct_bal
    else:
        acct_cad = acct.ConvertBalanceToCurrencyAsOfDate(acct_bal, acct_comm, cur, idate)
    # print_info("{} balance on {} = {}".format(acct.GetName(), idate, acct_cad))

    return gnc_numeric_to_python_decimal(acct_cad)


def get_period_sum(root_account, path, pdate, cur):
    acct = account_from_path(root_account, path)
    acct_name = acct.GetName()
    # get the split amounts for the parent account
    acct_sum = get_acct_bal(acct, pdate, cur)
    descendants = acct.get_descendants()
    if len(descendants) > 0:
        # for EACH sub-account add to the overall total
        # print("Descendants of {}:".format(acct_name))
        for sub_acct in descendants:
            # ?? GETTING SLIGHT ROUNDING ERRORS WHEN ADDING MUTUAL FUND VALUES...
            acct_sum += get_acct_bal(sub_acct, pdate, cur)

    return acct_name, acct_sum


# noinspection PyDictCreation
def fill_today(root_account, dest, cur):
    """
    Get Balance data for TODAY: LIABS, House, FAMILY, XCHALET, TRUST
    :param root_account: Gnucash account: from the Gnucash book
    :param dest: Google sheet to update
    :param cur: Gnucash Commodity: currency to use for the totals
    :return: list with values in Google format
    """
    data = []
    for item in BALANCE_ACCTS:
        path = BALANCE_ACCTS[item]
        acct_name, acct_sum = get_period_sum(root_account, path, today, cur)
        # need assets apart from house and liabilities
        if item == HOUSE:
            house_sum = acct_sum
        elif item == LIAB:
            liab_sum = acct_sum
        if item == ASTS:
            print_info("TOTAL for {} on {} = ${}".format(acct_name, today, acct_sum), MAGENTA)
            acct_sum = acct_sum - house_sum - liab_sum
        str_sum = acct_sum.to_eng_string()
        print_info("Balance for {} on {} = ${}".format(acct_name, today, str_sum), MAGENTA)

        cell = {}
        cell['range'] = dest + '!' + BAL_MTHLY_COLS[TODAY] + BAL_TODAY_RANGES[item]
        cell['values'] = [[str_sum]]
        print_info("cell = {}\n".format(cell))
        data.append(cell)

    return data


# noinspection PyDictCreation,PyDictCreation
def fill_current_year(root_account, dest, cur):
    """
    CURRENT YEAR: LIABS for ALL completed month_ends; FAMILY for ALL non-3 completed month_ends in year
    :param root_account: Gnucash account: from the Gnucash book
    :param dest: Google sheet to update
    :param cur: Gnucash Commodity: currency to use for the totals
    :return: list with values in Google format
    """
    data = []
    month = today.month
    print_info("It is now month '{}'".format(month))

    year = today.year
    end_prev_month = date(year, month, 1)

    month_ends = []
    for i in range(month-1):
        print("range = {}".format(i))
        month_ends.append(date(year, i+2, 1)-ONE_DAY)

    for dte in month_ends:
        print("date = {}-{}-{}".format(dte.year, dte.month, dte.day))

        # fill LIABS
        path = BALANCE_ACCTS[LIAB]
        acct_name, liab_sum = get_period_sum(root_account, path, dte, cur)
        str_sum = liab_sum.to_eng_string()
        print_info("{} on {} = ${}".format(acct_name, dte, str_sum), MAGENTA)
        cell = {}
        cell['range'] = dest + '!' + BAL_MTHLY_COLS[LIAB][MTH] + str(BASE_MTHLY_ROW + dte.month)
        cell['values'] = [[str_sum]]
        print_info("cell = {}\n".format(cell))
        data.append(cell)

        # fill ASSETS for months NOT covered by the Assets sheet
        if dte.month % 3 != 0:
            path = BALANCE_ACCTS[ASTS]
            acct_name, acct_sum = get_period_sum(root_account, path, dte, cur)
            corrected_assets = acct_sum - liab_sum
            str_sum = corrected_assets.to_eng_string()
            print_info("Assets on {} = ${}".format(dte, str_sum), MAGENTA)
            cell = {}
            cell['range'] = dest + '!' + BAL_MTHLY_COLS[ASTS] + str(BASE_MTHLY_ROW + dte.month)
            cell['values'] = [[str_sum]]
            print_info("cell = {}\n".format(cell))
            data.append(cell)

    return data


# noinspection PyDictCreation,PyDictCreation,PyDictCreation
def fill_previous_year(root_account, dest, cur):
    """
    CURRENT YEAR: LIABS for ALL completed months; FAMILY for ALL non-3 completed months in year
    :param root_account: Gnucash account: from the Gnucash book
    :param dest: Google sheet to update
    :param cur: Gnucash Commodity: currency to use for the totals
    :return: list with values in Google format
    """
    data = []
    month = today.month
    print_info("It is now month '{}'".format(month))

    year = today.year
    end_prev_month = date(year, month, 1)

    month_ends = []
    for i in range(12-month):
        print("range = {}".format(i))
        month_ends.append(date(year-1, i+5, 1)-ONE_DAY)
    year_end = date(year-1, 12, 31)
    month_ends.append(year_end)

    for dte in month_ends:
        print("date = {}-{}-{}".format(dte.year, dte.month, dte.day))

        # fill LIABS
        path = BALANCE_ACCTS[LIAB]
        acct_name, liab_sum = get_period_sum(root_account, path, dte, cur)
        str_sum = liab_sum.to_eng_string()
        print_info("{} on {} = ${}".format(acct_name, dte, str_sum), MAGENTA)
        cell = {}
        cell['range'] = dest + '!' + BAL_MTHLY_COLS[LIAB][MTH] + str(BASE_MTHLY_ROW + dte.month)
        cell['values'] = [[str_sum]]
        print_info("cell = {}\n".format(cell))
        data.append(cell)

        # fill ASSETS for months NOT covered by the Assets sheet
        if dte.month % 3 != 0:
            path = BALANCE_ACCTS[ASTS]
            acct_name, acct_sum = get_period_sum(root_account, path, dte, cur)
            corrected_assets = acct_sum - liab_sum
            str_sum = corrected_assets.to_eng_string()
            print_info("Assets on {} = ${}".format(dte, str_sum), MAGENTA)
            cell = {}
            cell['range'] = dest + '!' + BAL_MTHLY_COLS[ASTS] + str(BASE_MTHLY_ROW + dte.month)
            cell['values'] = [[str_sum]]
            print_info("cell = {}\n".format(cell))
            data.append(cell)

        # extra LIABS entry for year end
        if dte == year_end:
            cell = {}
            cell['range'] = dest + '!' + BAL_MTHLY_COLS[LIAB][YR] + str(BASE_ROW + year_span(year-1))
            cell['values'] = [[str_sum]]
            print_info("cell = {}\n".format(cell))
            data.append(cell)

    return data


# noinspection PyUnboundLocalVariable,PyUnresolvedReferences
def get_gnucash_data(gnucash_file, domain, dest):
    """
    Get Balance data for TODAY:
      LIABS, House, FAMILY, XCHALET, TRUST
    OR for the specified year:
      IF CURRENT YEAR: LIABS for ALL completed months; FAMILY for ALL non-3 completed months in year
      IF PREVIOUS YEAR: LIABS for year, for ALL NON-completed months; FAMILY for ALL non-3 NON-completed months in year
      ELSE: LIABS for year
    :param gnucash_file: string: name of file used to read the values
    :param domain: string: what to update
    :param dest: Google sheet to update
    :return: list: cells with location and value to update on Google sheet
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
        else:
            year = get_year(domain)
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
        exit(167)


def send_google_data(mode, bal_data):
    """
    Fill the data list and send to the document
    :param mode: string: '.?[send][1]'
    :param bal_data: list: Gnucash data for each needed quarter
    :return: server response
    """
    print_info("\nsend_google_data({})".format(mode))

    response = None
    try:
        assets_body = {
            'valueInputOption': 'USER_ENTERED',
            'data': bal_data
        }

        if 'send' in mode:
            creds = get_credentials()
            service = build('sheets', 'v4', credentials=creds)
            vals = service.spreadsheets().values()
            response = vals.batchUpdate(spreadsheetId=get_budget_id(), body=assets_body).execute()

            print_info('\n{} cells updated!'.format(response.get('totalUpdatedCells')))

    except Exception as se:
        print_error("Exception: {}!".format(se))
        exit(275)

    return response


def update_balance_main():
    """
    Main: check command line and call functions to get the data from Gnucash book and send to Google document
    :return: nil
    """
    exe = argv[0].split('/')[-1]
    if len(argv) < 4:
        print_error("NOT ENOUGH parameters!")
        print_info("usage: {} <book url> mode=<.?[send][1]> <year|'today'>".format(exe), GREEN)
        print_error("PROGRAM EXIT!")
        return

    gnucash_file = argv[1]

    mode = argv[2].lower()
    dest = BAL_2_SHEET
    if '1' in mode:
        dest = BAL_1_SHEET

    print_info("\nrunning '{}' on '{}' in mode '{}' at run-time: {}\n".format(exe, gnucash_file, mode, now), GREEN)

    domain = argv[3].lower()

    data = get_gnucash_data(gnucash_file, domain, dest)

    response = send_google_data(mode, data)
    if response:
        fname = "out/updateBalance_{}-response".format(domain)
        save_to_json(fname, now, response)

    print_info("\n >>> PROGRAM ENDED.", GREEN)


if __name__ == "__main__":
    update_balance_main()
