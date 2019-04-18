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
    ASTS  : [ASTS],
    LIAB  : [LIAB],
    TRUST : [TRUST],
    CHAL  : [CHAL],
    HOUSE : [ASTS, HOUSE]
}

# path to the accounts in the Gnucash file
BALANCE_RANGES = {
    ASTS  : 'Balance 1!C27' ,
    LIAB  : 'Balance 1!C28' ,
    TRUST : 'Balance 1!C21' ,
    CHAL  : 'Balance 1!C22' ,
    HOUSE : 'Balance 1!C26'
}

BALANCE_COLS = {
    LIAB  : {'Yrs': 'U', 'Mths': 'L'},
    TODAY : 'C',
    ASTS  : 'K'
}

BASE_YEAR = 2008
# number of rows between years
BASE_YEAR_SPAN = 1

BASE_MTHLY_ROW = 20

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


def get_acct_totals(root_account, bdate, cur):
    """
    Get BALANCE data for the specified account for the specified quarter
    :param root_account: Gnucash account: from the Gnucash book
    :param bdate: date: read the account total at the end of the quarter
    :param cur: Gnucash Commodity: currency to use for the totals
    :return: string with sum of totals
    """
    data = {}
    for item in BALANCE_ACCTS:
        path = BALANCE_ACCTS[item]
        acct = account_from_path(root_account, path)
        acct_name = acct.GetName()

        # get the split amounts for the parent account
        acct_sum = get_acct_bal(acct, bdate, cur)
        descendants = acct.get_descendants()
        if len(descendants) > 0:
            # for EACH sub-account add to the overall total
            # print("Descendants of {}:".format(acct_name))
            for sub_acct in descendants:
                # ?? GETTING SLIGHT ROUNDING ERRORS WHEN ADDING MUTUAL FUND VALUES...
                acct_sum += get_acct_bal(sub_acct, bdate, cur)

        str_sum = acct_sum.to_eng_string()
        print_info("Assets for {} on {} = ${}\n".format(acct_name, bdate, str_sum), MAGENTA)
        data[item] = str_sum

    return data


def fill_today(root_account, cur):
    """
    Get Balance data for TODAY: LIABS, House, FAMILY, XCHALET, TRUST
    :param root_account: Gnucash account: from the Gnucash book
    :param cur: Gnucash Commodity: currency to use for the totals
    :return: list with values in Google format
    """
    data = []
    for item in BALANCE_ACCTS:
        path = BALANCE_ACCTS[item]
        acct = account_from_path(root_account, path)
        acct_name = acct.GetName()

        # get the split amounts for the parent account
        acct_sum = get_acct_bal(acct, today, cur)
        descendants = acct.get_descendants()
        if len(descendants) > 0:
            # for EACH sub-account add to the overall total
            # print("Descendants of {}:".format(acct_name))
            for sub_acct in descendants:
                # ?? GETTING SLIGHT ROUNDING ERRORS WHEN ADDING MUTUAL FUND VALUES...
                acct_sum += get_acct_bal(sub_acct, today, cur)

        str_sum = acct_sum.to_eng_string()
        print_info("Assets for {} on {} = ${}".format(acct_name, today, str_sum), MAGENTA)

        cell = {}
        cell['range'] = BALANCE_RANGES[item]
        cell['values'] = [[str_sum]]
        print_info("cell = {}\n".format(cell))
        data.append(cell)

    return data


def fill_current_year(root_account, cur):
    """
    CURRENT YEAR: LIABS for ALL completed months; FAMILY for ALL non-3 completed months in year
    :param root_account: Gnucash account: from the Gnucash book
    :param cur: Gnucash Commodity: currency to use for the totals
    :return: list with values in Google format
    """
    data = []
    month = today.month
    print_info("It is now month '{}'".format(month))

    year = today.year
    end_prev_month = date(year, month, 1)

    months = []
    for i in range(month):
        print("range = {}".format(i))
        months.append(date(year, i+1, 1)-ONE_DAY)
    for it in months:
        print("date = {}-{}-{}".format(it.year, it.month, it.day))
    if year == 2019:
        return

    for item in BALANCE_ACCTS:
        path = BALANCE_ACCTS[item]
        acct = account_from_path(root_account, path)
        acct_name = acct.GetName()

        # get the split amounts for the parent account
        acct_sum = get_acct_bal(acct, today, cur)
        descendants = acct.get_descendants()
        if len(descendants) > 0:
            # for EACH sub-account add to the overall total
            # print("Descendants of {}:".format(acct_name))
            for sub_acct in descendants:
                # ?? GETTING SLIGHT ROUNDING ERRORS WHEN ADDING MUTUAL FUND VALUES...
                acct_sum += get_acct_bal(sub_acct, today, cur)

        str_sum = acct_sum.to_eng_string()
        print_info("Assets for {} on {} = ${}\n".format(acct_name, today, str_sum), MAGENTA)

        cell = {}
        cell['range'] = BALANCE_RANGES[item]
        cell['values'] = [[str_sum]]
        print_info("cell = {}".format(cell))
        data.append(cell)

    return data


# noinspection PyUnboundLocalVariable,PyUnresolvedReferences
def get_gnucash_data(gnucash_file, domain):
    """
    Get Balance data for TODAY:
      LIABS, House, FAMILY, XCHALET, TRUST
    OR for the specified year:
      IF CURRENT YEAR: LIABS for ALL completed months; FAMILY for ALL non-3 completed months in year
      IF PREVIOUS YEAR: LIABS for year, for ALL NON-completed months; FAMILY for ALL non-3 NON-completed months in year
      ELSE: LIABS for year
    :param gnucash_file: string: name of file used to read the values
    :param domain: string: what to update
    :return: nil
    """
    print_info("find Balances in {} for {}".format(gnucash_file, domain), GREEN)

    try:
        gnucash_session = Session(gnucash_file, is_new=False)
        root_account = gnucash_session.book.get_root_account()
        commod_tab = gnucash_session.book.get_table()
        # noinspection PyPep8Naming
        CAD = commod_tab.lookup("ISO4217", "CAD")

        if domain == 'today':
            data = fill_today(root_account, CAD)
        else:
            year = get_year(domain)
            if year == today.year:
                data = fill_current_year(root_account, CAD)
            elif year - today.year == 1:
                data = fill_previous_year(root_account)
            else:
                data = fill_year(year, root_account)

        # no save needed, we're just reading...
        gnucash_session.end()

        if data:
            fname = "out/updateBalance_data-{}".format(domain)
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
    :param mode: string: 'xxx[prod][send]'
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
        print_info("usage: {} <book url> mode=<.?[send]> <year|'today'>".format(exe), GREEN)
        print_error("PROGRAM EXIT!")
        return

    gnucash_file = argv[1]
    mode = argv[2].lower()
    print_info("\nrunning '{}' on '{}' in mode '{}' at run-time: {}\n".format(exe, gnucash_file, mode, now), GREEN)

    domain = argv[3].lower()

    bal_data = get_gnucash_data(gnucash_file, domain)

    response = send_google_data(mode, bal_data)
    if response:
        fname = "out/updateBalance_response-{}".format(domain)
        save_to_json(fname, now, response)

    print_info("\n >>> PROGRAM ENDED.", GREEN)


if __name__ == "__main__":
    update_balance_main()
